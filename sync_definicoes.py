# -*- coding: utf-8 -*-
"""
sync_definicoes.py — LAWgico Compromissos

Le a planilha "Alertas.xlsx" (aba "Alerta de Atividades") -- fonte unica de
verdade dos compromissos monitorados (decisao da usuaria em 24/08/2026: so
usar o que esta na planilha, em vez de tambem puxar responsaveis da planilha
"Divisao do time B e Controladoria*.xlsb", que podia incluir gente que nao
deveria receber lembrete, como advogados de processo). A planilha ja traz o
e-mail de cada responsavel direto (colunas "Integrante do time"/"E-mail"),
entao nao precisa mais resolver por roster/apelido.

Grava o catalogo normalizado na tabela compromissos_definicoes (Supabase).
Nao calcula "vence hoje" aqui -- isso e feito todo dia por
enviar_lembretes_diarios.py, que le esta tabela.

Rodar manualmente sempre que o Alertas.xlsx for atualizado com atividades/
pessoas/regras novas (nao e um watcher automatico).
"""
import re
import sys
import unicodedata
import logging
from datetime import datetime, timezone

import requests
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

ALERTAS_PATH = r"C:\Users\ach\OneDrive - Peixoto e Cury Advogados\Pastas - Time B\Controladoria\Compromissos\Alertas.xlsx"
ALERTAS_ABA = "Alerta de Atividades"

SUPABASE_URL = "https://rpibvjcnrseuugpkfmdj.supabase.co"
# Chave publica "anon" do mesmo projeto Supabase ja usado por outras
# ferramentas da Controladoria (Controle de Inserção de Cálculos, LAWgico
# Cadastro etc.) -- mesma chave ja embutida em HTML publico dessas
# ferramentas, ver "Base diária de atualização sócios/index.html".
SUPABASE_KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJwaWJ2amNucnNldXVncGtmbWRqIiwi"
                "cm9sZSI6ImFub24iLCJpYXQiOjE3ODE1NTc3MTcsImV4cCI6MjA5NzEzMzcxN30.ecihol8JESMH7cgFSvWKIzp-OwoPRFqdK3aCDwpCeg8")

LOG = logging.getLogger("sync_definicoes")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

WEEKDAY_MAP = {
    "segunda": 0, "terca": 1, "quarta": 2, "quinta": 3,
    "sexta": 4, "sabado": 5, "domingo": 6,
}
MESES_TRIMESTRE = [3, 6, 9, 12]

# Itens que ja mandavam lembrete DIARIO de verdade antes desta planilha virar
# fonte unica (confirmado funcionando em producao) mas que no Alertas.xlsx
# estao com Periodicidade "Mensal" + "Quando fica o Alerta?" = "Todo dia"
# generico, sem dia especifico -- o que cairia em "indefinida" pela regra
# geral. Decisao confirmada com a usuaria em 24/08/2026: manter diarias.
FORCAR_DIARIA = {
    ("giullia leme", "cobrar designacao de audiencia pendente"),
    ("giullia leme", "cobrar pos audiencia"),
    ("giullia leme", "enviar invite da pauta"),
    ("giullia leme", "atualizar o dash de baixados"),
    ("giullia leme", "rodar o prompt injection nos cadastros do dia"),
    ("beatriz luna", "cobrar as justificativas dos fatais"),
}


def norm(s):
    if not s:
        return ""
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s)


def extrair_dias_semana(texto):
    t = norm(texto)
    return sorted({num for chave, num in WEEKDAY_MAP.items() if chave in t})


def normalizar_regra_alerta(quando, periodicidade):
    """Deriva regra_tipo/regra_dias_semana/regra_dia_mes/regra_meses a partir
    das colunas 'Quando fica o Alerta?' e 'Periodicidade' do Alertas.xlsx.

    Dado real: as duas colunas as vezes se contradizem (ex: Periodicidade diz
    'Mensal' mas Quando so diz 'Todo dia', sem dia especifico -- ou o dia da
    semana foi digitado em Periodicidade em vez de em Quando, como nas linhas
    de Embraer/Bimbo/CBC). Por isso cruza as duas colunas em vez de confiar
    só numa."""
    tq = norm(quando)
    tp = norm(periodicidade)
    combinado = f"{tq} {tp}"

    m = re.search(r'todo dia (\d{1,2})', tq)
    if m:
        dia = int(m.group(1))
        meses = MESES_TRIMESTRE if "trimestral" in tp else None
        return {"regra_tipo": "dia_mes", "regra_dia_mes": dia, "regra_meses": meses}

    dias = extrair_dias_semana(combinado)
    if dias:
        return {"regra_tipo": "dias_semana", "regra_dias_semana": dias}

    if "diari" in tp and "todo dia" in tq:
        return {"regra_tipo": "diaria"}

    # Periodicidade Mensal/Trimestral/Semanalmente/"Quando solicitado"/etc
    # sem dia especifico na planilha -- nao da pra calcular sozinho quando
    # vence; fica cadastrado mas sem lembrete automatico ate a usuaria
    # preencher um dia (decisao confirmada em 24/08/2026, pra nao mandar
    # lembrete diario de tarefa que na verdade e mensal/trimestral).
    return {"regra_tipo": "indefinida"}


def extrair_alertas(caminho):
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    ws = wb[ALERTAS_ABA]
    rows = list(ws.iter_rows(values_only=True))
    definicoes = []
    for r in rows[1:]:
        if not r or not r[0] or not r[1]:
            continue
        nome = str(r[0]).strip()
        email = str(r[1]).strip()
        sigla = (str(r[2]).strip() if r[2] else None) or None
        quando, horario, funcao = r[3], r[4], r[5]
        grupo, cliente, periodicidade, item = r[6], r[7], r[8], r[9]
        item = (item or "").strip()

        regra = normalizar_regra_alerta(quando, periodicidade)
        if (norm(nome), norm(item)) in FORCAR_DIARIA:
            regra = {"regra_tipo": "diaria"}

        definicoes.append({
            "origem": "alerta",
            "responsavel_nome": nome,
            "responsavel_sigla": sigla,
            "responsavel_email": email,
            "responsavel_papel": (str(funcao).strip() if funcao else None),
            "grupo": grupo,
            "cliente": (cliente or "").strip() if cliente else cliente,
            "item": item,
            "etapa": None,
            "periodicidade_original": (periodicidade or "").strip() if periodicidade else periodicidade,
            **regra,
            "ativo": True,
        })
    return definicoes


COLUNAS_DEFINICAO = [
    "origem", "responsavel_nome", "responsavel_sigla", "responsavel_email", "responsavel_papel",
    "grupo", "cliente", "item", "etapa", "periodicidade_original",
    "regra_tipo", "regra_dias_semana", "regra_dia_mes", "regra_meses", "ativo",
]


def normalizar_chaves(definicoes):
    """O insert em lote do PostgREST exige que todo objeto tenha as mesmas
    chaves -- garante isso independente do que cada linha do Alertas.xlsx
    tiver preenchido."""
    return [{col: d.get(col) for col in COLUNAS_DEFINICAO} for d in definicoes]


def gravar_supabase(definicoes):
    headers = {
        "apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json", "Prefer": "return=minimal",
    }
    agora = datetime.now(timezone.utc).isoformat()
    # marca tudo como inativo antes de re-inserir o que a planilha atual tem
    # (assim item removido da planilha some do painel sem deletar historico de entregas)
    r = requests.patch(f"{SUPABASE_URL}/rest/v1/compromissos_definicoes?ativo=eq.true",
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
    LOG.info(f"Lendo {ALERTAS_PATH}")
    definicoes = extrair_alertas(ALERTAS_PATH)

    por_regra = {}
    for d in definicoes:
        por_regra[d["regra_tipo"]] = por_regra.get(d["regra_tipo"], 0) + 1
    LOG.info(f"Total de definicoes: {len(definicoes)} — {por_regra}")

    indefinidas = [d for d in definicoes if d["regra_tipo"] == "indefinida"]
    if indefinidas:
        LOG.info(f"{len(indefinidas)} sem regra automatica computavel (periodicidade "
                  f"mensal/trimestral/sob-demanda sem dia especifico na planilha) — "
                  f"ficam cadastradas mas sem lembrete automatico:")
        for d in indefinidas:
            LOG.info(f"  indefinida: {d['responsavel_nome']!r} — {d['item']!r} "
                      f"(periodicidade original: {d['periodicidade_original']!r})")

    pessoas_distintas = {(d["responsavel_nome"], d["responsavel_email"]) for d in definicoes}
    LOG.info(f"Pessoas distintas na planilha: {len(pessoas_distintas)}")

    gravar_supabase(definicoes)
    LOG.info("Gravado no Supabase com sucesso.")


if __name__ == "__main__":
    main()
