# -*- coding: utf-8 -*-
"""
sync_definicoes.py — LAWgico Compromissos
Le a planilha "Divisao do time B e Controladoria - DD.MM.AAAA.xlsx" (abas
"Atividades do time", "Datas - Pauta", "Relatorios - Data"), normaliza a regra
de recorrencia de cada linha e resolve o e-mail de cada responsavel via o
roster profissionais_peixoto(editado).xlsx (por Sigla/matricula quando
disponivel; por nome, com mapa de apelidos, nas abas que so tem nome livre).

Grava o catalogo normalizado na tabela compromissos_definicoes (Supabase).
Nao calcula "vence hoje" aqui -- isso e feito todo dia por
enviar_lembretes_diarios.py, que le esta tabela.

Rodar manualmente sempre que a planilha for atualizada com clientes/pessoas/
regras novas (nao e um watcher automatico).
"""
import glob
import os
import re
import sys
import unicodedata
import logging
from datetime import datetime, timezone

import requests
from pyxlsb import open_workbook
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

PASTA_DIVISAO = r"C:\Users\ach\OneDrive - Peixoto e Cury Advogados\Pastas - Time B"
ROSTER_PATH = r"C:\Users\ach\OneDrive - Peixoto e Cury Advogados\Pastas - Time B\Controladoria\Baixados\profissionais_peixoto(editado).xlsx"

SUPABASE_URL = "https://rpibvjcnrseuugpkfmdj.supabase.co"
# Chave publica "anon" do mesmo projeto Supabase ja usado por outras
# ferramentas da Controladoria (Controle de Inserção de Cálculos, LAWgico
# Cadastro etc.) -- mesma chave ja embutida em HTML publico dessas
# ferramentas, ver "Base diária de atualização sócios/index.html".
SUPABASE_KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJwaWJ2amNucnNldXVncGtmbWRqIiwi"
                "cm9sZSI6ImFub24iLCJpYXQiOjE3ODE1NTc3MTcsImV4cCI6MjA5NzEzMzcxN30.ecihol8JESMH7cgFSvWKIzp-OwoPRFqdK3aCDwpCeg8")

LOG = logging.getLogger("sync_definicoes")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

# ── Apelidos que nao batem por aproximacao de nome com o roster ──────────────
# (confirmado manualmente cruzando com profissionais_peixoto(editado).xlsx)
ALIAS_EMAIL = {
    "gabi": "cicera.vicente@peixotoecury.com.br",
    "carol": "anacarolina.martini@peixotoecury.com.br",
    "carol honorio": "anacarolina.martini@peixotoecury.com.br",
    "bruna": "bruna.fazenda@peixotoecury.com.br",
    "bruna cembranelli": "bruna.fazenda@peixotoecury.com.br",
    "rafaella": "rafaella.domenicis@peixotoecury.com.br",
    "rafaella domenicis": "rafaella.domenicis@peixotoecury.com.br",
    "rafaella wittmann": "rafaella.domenicis@peixotoecury.com.br",
    "maria carolina": "mariacarolina.dias@peixotoecury.com.br",
    "evandro": "evandro.vessoni@peixotoecury.com.br",
    "evandro vessoni": "evandro.vessoni@peixotoecury.com.br",
    "anna carolina": "annacarolina.baani@peixotoecury.com.br",
    "anna carolina baani": "annacarolina.baani@peixotoecury.com.br",
    "anna carolina balbino": "annacarolina.baani@peixotoecury.com.br",
    "brenna": "brenna.malta@peixotoecury.com.br",
    "brenna malta": "brenna.malta@peixotoecury.com.br",
    "leticia": "leticia.fontes@peixotoecury.com.br",
    "leticia meneghini": "leticia.fontes@peixotoecury.com.br",
    "vanessa": "vanessa.rosa@peixotoecury.com.br",
    "vanessa rosa": "vanessa.rosa@peixotoecury.com.br",
    "nicolau guimaraes": "nicolau.coelho@peixotoecury.com.br",
    "rodrigo souichi": "rodrigo.nishijima@peixotoecury.com.br",
    "guilherme morimoto": "guilherme.morimoto@peixotoecury.com.br",
    "geovanna rodrigues": "geovanna.rodrigues@peixotoecury.com.br",
    "raica kwok": "raica.kwok@peixotoecury.com.br",
    "rafael rocha ferreira": "rafael.ferreira@peixotoecury.com.br",
    "juliana perucci": "juliana.perucci@peixotoecury.com.br",
    # Confirmados diretamente pela usuaria (nao estavam no roster
    # profissionais_peixoto(editado).xlsx -- mesmo gap ja conhecido de
    # JNE/Jennifer em outras ferramentas desta base, ver memoria
    # dashboard_baixados_analises).
    "jennifer nogueira": "jennifer.nogueira@peixotoecury.com.br",
    "lucas gusthavo": "lucas.caliman@peixotoecury.com.br",
    "mariana arruda terberges": "mariana.teberges@peixotoecury.com.br",
}

# Mesmos 3 casos acima, mas pela SIGLA -- usado por extrair_atividades(), que
# resolve e-mail via matricula (sigla) do roster, nao por nome.
SIGLA_EMAIL_OVERRIDE = {
    "JNE": "jennifer.nogueira@peixotoecury.com.br",
    "LUC": "lucas.caliman@peixotoecury.com.br",
    "MTE": "mariana.teberges@peixotoecury.com.br",
}

WEEKDAY_MAP = {
    "segunda": 0, "terca": 1, "quarta": 2, "quinta": 3,
    "sexta": 4, "sabado": 5, "domingo": 6,
}
MESES_TRIMESTRE = [3, 6, 9, 12]


def norm(s):
    if not s:
        return ""
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s)


def localizar_xlsb_mais_recente():
    candidatos = glob.glob(os.path.join(PASTA_DIVISAO, "Divisão do time B e Controladoria*.xlsb"))
    if not candidatos:
        raise FileNotFoundError("Nenhum arquivo 'Divisão do time B e Controladoria*.xlsb' encontrado.")
    return max(candidatos, key=os.path.getmtime)


def carregar_roster():
    """matricula -> {nome, email}; nome_normalizado -> email (para busca por aproximacao)."""
    wb = openpyxl.load_workbook(ROSTER_PATH, data_only=True, read_only=True)
    ws = wb["Planilha2"]
    rows = list(ws.iter_rows(values_only=True))
    idx = {h: i for i, h in enumerate(rows[0])}
    por_matricula, por_nome = {}, []
    for r in rows[1:]:
        matricula = r[idx["matricula"]]
        nome = (r[idx["nome"]] or "").strip()
        email = r[idx["email_pc"]]
        if not matricula or not email:
            continue
        por_matricula[str(matricula).strip().upper()] = {"nome": nome, "email": email}
        por_nome.append({"nome": nome, "email": email})
    return por_matricula, por_nome


def resolver_por_nome(nome_livre, por_nome):
    chave = norm(nome_livre)
    if chave in ALIAS_EMAIL:
        return ALIAS_EMAIL[chave]
    tokens_curto = [t for t in chave.split(" ") if len(t) > 1]
    candidatos = []
    for r in por_nome:
        tokens_roster = norm(r["nome"]).split(" ")
        if all(any(tr.startswith(tc) or tc.startswith(tr) for tr in tokens_roster) for tc in tokens_curto):
            candidatos.append(r)
    if len(candidatos) == 1:
        return candidatos[0]["email"]
    LOG.warning(f"Nao foi possivel resolver e-mail para nome livre {nome_livre!r} "
                f"({'ambiguo, ' + str(len(candidatos)) + ' candidatos' if candidatos else 'sem match'}) — "
                f"adicionar em ALIAS_EMAIL.")
    return None


def extrair_dias_semana(texto):
    t = norm(texto)
    dias = sorted({num for chave, num in WEEKDAY_MAP.items() if chave in t})
    return dias


def normalizar_periodicidade_atividade(texto):
    t = norm(texto)
    if "diari" in t:
        return {"regra_tipo": "diaria"}
    dias = extrair_dias_semana(t)
    if dias:
        return {"regra_tipo": "dias_semana", "regra_dias_semana": dias}
    return {"regra_tipo": "indefinida"}


def normalizar_dia_pauta(texto):
    t = norm(texto)
    if "final de mes" in t or "fim de mes" in t:
        return {"regra_tipo": "fim_de_mes"}
    dias = extrair_dias_semana(t)
    if dias:
        return {"regra_tipo": "dias_semana", "regra_dias_semana": dias}
    return {"regra_tipo": "indefinida"}


def meses_elegiveis_relatorio(periodicidade_texto):
    t = norm(periodicidade_texto)
    if "trimestral" in t:
        return MESES_TRIMESTRE
    if "mensal" in t:
        return list(range(1, 13))
    return None  # semanal / quando solicitado / outros -> nao computavel por dia do mes


def to_int_dia(v):
    try:
        n = int(float(v))
        return n if 1 <= n <= 31 else None
    except (TypeError, ValueError):
        return None


def extrair_atividades(caminho, por_matricula):
    with open_workbook(caminho) as wb:
        with wb.get_sheet("Atividades do time") as sheet:
            linhas = [[c.v for c in r] for r in sheet.rows()]
    definicoes = []
    for r in linhas[1:]:
        if not r or not r[0]:
            continue
        # algumas linhas tem responsavel composto ("ANB E MCD" / "Fulano e Beltrana")
        siglas = [s.strip().upper() for s in re.split(r"\s+e\s+", r[1] or "", flags=re.IGNORECASE) if s.strip()]
        nomes = [n.strip() for n in re.split(r"\s+e\s+", r[0] or "", flags=re.IGNORECASE) if n.strip()]
        if not siglas:
            siglas = [""]
        regra = normalizar_periodicidade_atividade(r[5])
        for i, sigla in enumerate(siglas):
            pessoa = por_matricula.get(sigla)
            email = pessoa["email"] if pessoa else SIGLA_EMAIL_OVERRIDE.get(sigla)
            nome = nomes[i] if i < len(nomes) else (nomes[0] if nomes else "")
            definicoes.append({
                "origem": "atividade",
                "responsavel_nome": nome,
                "responsavel_sigla": sigla or None,
                "responsavel_email": email,
                "grupo": r[3], "cliente": r[4],
                "item": (r[6] or "").strip(),
                "etapa": None,
                "periodicidade_original": (r[5] or "").strip(),
                **regra,
                "ativo": True,
            })
            if sigla and not email:
                LOG.warning(f"Sigla {sigla!r} (atividade: {r[6]!r}) nao encontrada no roster nem em SIGLA_EMAIL_OVERRIDE.")
    return definicoes


def extrair_pauta(caminho, por_nome):
    with open_workbook(caminho) as wb:
        with wb.get_sheet("Datas - Pauta") as sheet:
            linhas = [[c.v for c in r] for r in sheet.rows()]
    definicoes = []
    for r in linhas[1:]:
        if not r or not r[1]:
            continue
        email = resolver_por_nome(r[1], por_nome)
        for etapa, col in [("Elaboração da Pauta", 3), ("Envio da Pauta ao Cliente", 4)]:
            regra = normalizar_dia_pauta(r[col])
            definicoes.append({
                "origem": "pauta",
                "responsavel_nome": (r[1] or "").strip(),
                "responsavel_sigla": None,
                "responsavel_email": email,
                "grupo": r[0], "cliente": (r[2] or "").strip(),
                "item": f"Pauta — {(r[2] or '').strip()}",
                "etapa": etapa,
                "periodicidade_original": (r[col] or "").strip(),
                **regra,
                "ativo": True,
            })
    return definicoes


def extrair_relatorios(caminho, por_nome):
    with open_workbook(caminho) as wb:
        with wb.get_sheet("Relatórios - Data") as sheet:
            linhas = [[c.v for c in r] for r in sheet.rows()]
    definicoes = []
    etapas_colunas = [("Corte", 7), ("Início da Elaboração", 8), ("Revisão", 9), ("Envio ao cliente", 10)]
    for r in linhas[1:]:
        if not r or not r[1]:
            continue
        if (r[3] or "").strip().lower() != "sim":
            continue  # "Tem relatório?" == Não
        meses = meses_elegiveis_relatorio(r[6])
        if meses is None:
            continue  # periodicidade sem dia-do-mes computavel (Semanal, Quando solicitado etc.)
        assistentes = [a.strip() for a in re.split(r"\s+e\s+", r[11] or "") if a.strip()]
        for etapa, col in etapas_colunas:
            dia = to_int_dia(r[col])
            if dia is None:
                continue
            for assistente in assistentes:
                email = resolver_por_nome(assistente, por_nome)
                definicoes.append({
                    "origem": "relatorio",
                    "responsavel_nome": assistente,
                    "responsavel_sigla": None,
                    "responsavel_email": email,
                    "grupo": r[0], "cliente": (r[1] or "").strip(),
                    "item": (r[4] or "").strip(),
                    "etapa": etapa,
                    "periodicidade_original": (r[6] or "").strip(),
                    "regra_tipo": "dia_mes",
                    "regra_dia_mes": dia,
                    "regra_meses": meses,
                    "ativo": True,
                })
    return definicoes


# ── Compromissos informados diretamente pela usuária (ainda não estão na
# planilha "Atividades do time" -- Giullia e Bia Luna são Equipe Controladoria
# Geral, um time diferente do Time B que a planilha cobre). Revisar/mover pra
# a planilha quando ela for atualizada; até lá, manter aqui. ──────────────────
COMPROMISSOS_MANUAIS = [
    {"responsavel_nome": "Giullia Peleje Leme", "responsavel_email": "giullia.leme@peixotoecury.com.br",
     "item": "Envio da Pauta — manhã (6h)", "regra_tipo": "diaria"},
    {"responsavel_nome": "Giullia Peleje Leme", "responsavel_email": "giullia.leme@peixotoecury.com.br",
     "item": "Envio da Pauta — tarde (17h)", "regra_tipo": "diaria"},
    {"responsavel_nome": "Giullia Peleje Leme", "responsavel_email": "giullia.leme@peixotoecury.com.br",
     "item": "Atualizar planilha de Baixados", "regra_tipo": "dias_semana", "regra_dias_semana": [0]},
    {"responsavel_nome": "Giullia Peleje Leme", "responsavel_email": "giullia.leme@peixotoecury.com.br",
     "item": "Cobrança de designações não feitas", "regra_tipo": "diaria"},
    {"responsavel_nome": "Giullia Peleje Leme", "responsavel_email": "giullia.leme@peixotoecury.com.br",
     "item": "Cobrança do Pós-Audiência", "regra_tipo": "diaria"},
    {"responsavel_nome": "Beatriz Rodrigues Luna", "responsavel_email": "beatriz.rodrigues@peixotoecury.com.br",
     "item": "Cobrança de justificativas dos fatais", "regra_tipo": "diaria"},
]


def extrair_manuais():
    definicoes = []
    for m in COMPROMISSOS_MANUAIS:
        definicoes.append({
            "origem": "atividade",
            "responsavel_sigla": None,
            "grupo": "Controladoria", "cliente": "TODOS",
            "etapa": None,
            "periodicidade_original": None,
            "regra_dias_semana": None, "regra_dia_mes": None, "regra_meses": None,
            "ativo": True,
            **m,
        })
    return definicoes


COLUNAS_DEFINICAO = [
    "origem", "responsavel_nome", "responsavel_sigla", "responsavel_email",
    "grupo", "cliente", "item", "etapa", "periodicidade_original",
    "regra_tipo", "regra_dias_semana", "regra_dia_mes", "regra_meses", "ativo",
]


def normalizar_chaves(definicoes):
    """O insert em lote do PostgREST exige que todo objeto tenha as mesmas
    chaves -- as 3 funcoes de extracao produzem dicts com conjuntos de
    chaves diferentes (nem toda origem usa regra_dia_mes/regra_meses etc)."""
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
    caminho = localizar_xlsb_mais_recente()
    LOG.info(f"Lendo {caminho}")
    por_matricula, por_nome = carregar_roster()

    definicoes = []
    definicoes += extrair_atividades(caminho, por_matricula)
    definicoes += extrair_pauta(caminho, por_nome)
    definicoes += extrair_relatorios(caminho, por_nome)
    definicoes += extrair_manuais()

    sem_email = [d for d in definicoes if not d["responsavel_email"]]
    LOG.info(f"Total de definicoes: {len(definicoes)} "
             f"(atividade={sum(1 for d in definicoes if d['origem']=='atividade')}, "
             f"pauta={sum(1 for d in definicoes if d['origem']=='pauta')}, "
             f"relatorio={sum(1 for d in definicoes if d['origem']=='relatorio')})")
    LOG.info(f"Sem e-mail resolvido: {len(sem_email)}")
    for d in sem_email:
        LOG.info(f"  sem e-mail: {d['responsavel_nome']!r} ({d['origem']}, {d['item']!r})")

    gravar_supabase(definicoes)
    LOG.info("Gravado no Supabase com sucesso.")


if __name__ == "__main__":
    main()
