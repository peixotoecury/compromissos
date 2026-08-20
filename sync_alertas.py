# -*- coding: utf-8 -*-
"""
sync_alertas.py — LAWgico Compromissos

Lê Alertas.xlsx (aba "Alerta de Atividades") — a lista curada manualmente
pela coordenadora de quais alertas devem efetivamente ser disparados, já com
e-mail e horário definidos linha a linha — e grava em compromissos_definicoes
com origem='alerta'.

Não toca nas definições origem='atividade'/'pauta'/'relatorio' (essas
continuam vindo do sync_definicoes.py, que lê o .xlsb e alimenta o catálogo
geral mostrado no painel). Este script só gerencia o subconjunto 'alerta',
que é o que enviar_alertas_horario.py usa pra disparar no horário certo.

Pré-requisito: rodar add_horario.sql no SQL Editor do Supabase antes da
primeira execução (adiciona a coluna 'horario').

Rodar manualmente sempre que Alertas.xlsx for atualizado.
"""
import re
import sys
import unicodedata
import logging
from datetime import datetime, timezone

import requests
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

CAMINHO_ALERTAS = r"C:\Users\ach\OneDrive - Peixoto e Cury Advogados\Pastas - Time B\Controladoria\Compromissos\Alertas.xlsx"
ABA = "Alerta de Atividades"

SUPABASE_URL = "https://rpibvjcnrseuugpkfmdj.supabase.co"
SUPABASE_KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJwaWJ2amNucnNldXVncGtmbWRqIiwi"
                "cm9sZSI6ImFub24iLCJpYXQiOjE3ODE1NTc3MTcsImV4cCI6MjA5NzEzMzcxN30.ecihol8JESMH7cgFSvWKIzp-OwoPRFqdK3aCDwpCeg8")

LOG = logging.getLogger("sync_alertas")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

WEEKDAY_MAP = {
    "segunda": 0, "terca": 1, "quarta": 2, "quinta": 3,
    "sexta": 4, "sabado": 5, "domingo": 6,
}


def norm(s):
    if not s:
        return ""
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s)


def extrair_dias_semana(texto_norm):
    return sorted({num for chave, num in WEEKDAY_MAP.items() if chave in texto_norm})


def normalizar_quando_alerta(texto):
    """'Todo dia' | 'Todo dia N' | 'Todo dia N do mês' | 'Toda <dia da semana>'."""
    t = norm(texto)
    m = re.match(r"^todo dia (\d+)", t)
    if m:
        return {"regra_tipo": "dia_mes", "regra_dia_mes": int(m.group(1))}
    if t == "todo dia":
        return {"regra_tipo": "diaria"}
    dias = extrair_dias_semana(t)
    if dias:
        return {"regra_tipo": "dias_semana", "regra_dias_semana": dias}
    LOG.warning(f"Nao reconheci a recorrencia {texto!r} -- marcando como 'indefinida'.")
    return {"regra_tipo": "indefinida"}


def normalizar_horario(texto):
    """'16h' -> '16:00', '9h15' -> '09:15', '9h30min' -> '09:30'. Casos ambíguos são
    logados como warning e tratados com o melhor palpite (revisar na planilha)."""
    t = str(texto).strip().lower()
    m = re.match(r"^(\d{1,2})h(\d{0,2})", t)
    if m:
        h = int(m.group(1))
        mm = int(m.group(2)) if m.group(2) else 0
        return f"{h:02d}:{mm:02d}"
    m2 = re.match(r"^(\d{1,2})min$", t)
    if m2:
        LOG.warning(f"Horario ambiguo {texto!r} (sem 'h') -- tratando como {int(m2.group(1)):02d}:00. "
                    f"Revisar essa celula na planilha original.")
        return f"{int(m2.group(1)):02d}:00"
    LOG.warning(f"Nao consegui normalizar o horario {texto!r} -- item ficara sem horario definido.")
    return None


COLUNAS_DEFINICAO = [
    "origem", "responsavel_nome", "responsavel_sigla", "responsavel_email",
    "grupo", "cliente", "item", "etapa", "periodicidade_original",
    "regra_tipo", "regra_dias_semana", "regra_dia_mes", "regra_meses",
    "horario", "ativo",
]


def extrair_alertas():
    wb = openpyxl.load_workbook(CAMINHO_ALERTAS, data_only=True)
    ws = wb[ABA]
    linhas = list(ws.iter_rows(values_only=True))
    definicoes = []
    for r in linhas[1:]:
        if not r or not r[0]:
            continue
        nome, email, sigla, quando, horario_txt, funcao, grupo, cliente, periodicidade, atividade = (list(r) + [None] * 10)[:10]
        regra = normalizar_quando_alerta(quando)
        horario = normalizar_horario(horario_txt) if horario_txt else None
        definicoes.append({
            "origem": "alerta",
            "responsavel_nome": (nome or "").strip(),
            "responsavel_sigla": (sigla or "").strip().upper() or None,
            "responsavel_email": (email or "").strip() or None,
            "grupo": (grupo or "").strip() or None,
            "cliente": (cliente or "").strip() or None,
            "item": (atividade or "").strip(),
            "etapa": None,
            "periodicidade_original": (quando or "").strip(),
            **regra,
            "horario": horario,
            "ativo": True,
        })
        if not definicoes[-1]["responsavel_email"]:
            LOG.warning(f"Linha sem e-mail: {nome!r} / {atividade!r}")
    return definicoes


def normalizar_chaves(definicoes):
    return [{col: d.get(col) for col in COLUNAS_DEFINICAO} for d in definicoes]


def gravar_supabase(definicoes):
    headers = {
        "apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json", "Prefer": "return=minimal",
    }
    agora = datetime.now(timezone.utc).isoformat()
    # marca inativo só o que já era origem='alerta' (não toca em atividade/pauta/relatorio)
    r = requests.patch(f"{SUPABASE_URL}/rest/v1/compromissos_definicoes?origem=eq.alerta&ativo=eq.true",
                        headers=headers, json={"ativo": False}, timeout=30)
    r.raise_for_status()
    definicoes = normalizar_chaves(definicoes)
    for d in definicoes:
        d["atualizado_em"] = agora
    r = requests.post(f"{SUPABASE_URL}/rest/v1/compromissos_definicoes",
                       headers={**headers, "Prefer": "resolution=merge-duplicates,return=minimal"},
                       json=definicoes, timeout=60)
    if not r.ok:
        LOG.error(f"Resposta do Supabase ({r.status_code}): {r.text}")
    r.raise_for_status()


def main():
    LOG.info(f"Lendo {CAMINHO_ALERTAS}")
    definicoes = extrair_alertas()
    LOG.info(f"Total de alertas curados: {len(definicoes)}")
    horarios = sorted({d["horario"] for d in definicoes if d["horario"]})
    LOG.info(f"Horarios distintos encontrados: {horarios}")
    sem_horario = [d for d in definicoes if not d["horario"]]
    if sem_horario:
        LOG.warning(f"{len(sem_horario)} item(ns) sem horario valido — nao serao disparados por "
                     f"enviar_alertas_horario.py em nenhum horario (rever normalizar_horario ou a planilha).")

    gravar_supabase(definicoes)
    LOG.info("Gravado no Supabase com sucesso (origem='alerta').")


if __name__ == "__main__":
    main()
