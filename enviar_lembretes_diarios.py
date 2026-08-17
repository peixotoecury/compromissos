# -*- coding: utf-8 -*-
"""
enviar_lembretes_diarios.py — LAWgico Compromissos

Roda 1x por dia (Tarefa Agendada do Windows, dias uteis, ex: 07h):
1. Le compromissos_definicoes (Supabase) e calcula quais valem HOJE
   (diaria / dias_semana / fim_de_mes / dia_mes-com-meses-elegiveis).
2. Cria (upsert, sem sobrescrever quem ja marcou entregue) 1 linha
   'pendente' em compromissos_entregas por definicao valida hoje.
3. Agrupa TODAS as pendencias em aberto (hoje + atrasadas) por pessoa e
   manda 1 e-mail por pessoa via automacao do Outlook (.Send() de verdade,
   nao .Display() -- ninguem precisa clicar em nada).

MODO_TESTE=True (padrao): manda tudo pra TESTE_EMAIL em vez do destinatario
real, com o nome de quem receberia no assunto -- usar assim até a usuária
validar as regras por alguns dias. Depois, mudar pra False.
"""
import calendar
import logging
import sys
from collections import defaultdict
from datetime import date, timedelta

import requests
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

SUPABASE_URL = "https://rpibvjcnrseuugpkfmdj.supabase.co"
SUPABASE_KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJwaWJ2amNucnNldXVncGtmbWRqIiwi"
                "cm9sZSI6ImFub24iLCJpYXQiOjE3ODE1NTc3MTcsImV4cCI6MjA5NzEzMzcxN30.ecihol8JESMH7cgFSvWKIzp-OwoPRFqdK3aCDwpCeg8")
HEADERS = {"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}", "Content-Type": "application/json"}

COMPROMISSOS_URL = "https://peixotoecury.github.io/compromissos/"  # atualizar apos publicar o repo

# ── MODO TESTE — deixar True até a usuaria validar as regras por alguns dias ──
MODO_TESTE = True
TESTE_EMAIL = "claude.controladoria@peixotoecury.com.br"

LOG = logging.getLogger("enviar_lembretes_diarios")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

DIAS_SEMANA_PT = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado", "domingo"]


def buscar_definicoes_ativas():
    r = requests.get(f"{SUPABASE_URL}/rest/v1/compromissos_definicoes?ativo=eq.true&select=*",
                      headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.json()


def ultimo_dia_do_mes(ano, mes):
    return calendar.monthrange(ano, mes)[1]


def vale_hoje(d, hoje):
    tipo = d.get("regra_tipo")
    if tipo == "diaria":
        return True
    if tipo == "dias_semana":
        dias = d.get("regra_dias_semana") or []
        return hoje.weekday() in dias
    if tipo == "fim_de_mes":
        return hoje.day == ultimo_dia_do_mes(hoje.year, hoje.month)
    if tipo == "dia_mes":
        dia_alvo = d.get("regra_dia_mes")
        meses = d.get("regra_meses")
        if dia_alvo is None:
            return False
        if meses and hoje.month not in meses:
            return False
        dia_alvo_ajustado = min(dia_alvo, ultimo_dia_do_mes(hoje.year, hoje.month))
        return hoje.day == dia_alvo_ajustado
    return False  # 'indefinida' -- nunca dispara sozinho


def upsert_pendencias_hoje(definicoes, hoje):
    aplicaveis_hoje = [d for d in definicoes if vale_hoje(d, hoje)]
    LOG.info(f"Definições válidas hoje ({hoje.isoformat()}): {len(aplicaveis_hoje)}")
    if not aplicaveis_hoje:
        return
    linhas = [{"definicao_id": d["id"], "data": hoje.isoformat(), "status": "pendente"} for d in aplicaveis_hoje]
    r = requests.post(
        f"{SUPABASE_URL}/rest/v1/compromissos_entregas",
        headers={**HEADERS, "Prefer": "resolution=ignore-duplicates,return=minimal"},
        json=linhas, timeout=30,
    )
    r.raise_for_status()


def buscar_pendencias_em_aberto():
    """Tudo que ainda esta 'pendente' (hoje ou atrasado), com a definicao embutida."""
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/compromissos_entregas"
        f"?status=eq.pendente&select=*,compromissos_definicoes(*)&order=data.asc",
        headers=HEADERS, timeout=30,
    )
    r.raise_for_status()
    return [e for e in r.json() if e.get("compromissos_definicoes")]


def montar_corpo(nome, itens, hoje):
    itens_hoje = [i for i in itens if i["data"] == hoje.isoformat()]
    itens_atrasados = [i for i in itens if i["data"] < hoje.isoformat()]

    def linha(i):
        d = i["compromissos_definicoes"]
        partes = [d.get("item") or ""]
        if d.get("etapa"):
            partes.append(f"({d['etapa']})")
        if d.get("cliente"):
            partes.append(f"— {d['cliente']}")
        return " ".join(partes)

    corpo = f"Olá, {nome.split(' ')[0]}!<br><br>"
    corpo += f"Segue o que está pendente pra você hoje ({DIAS_SEMANA_PT[hoje.weekday()]}, {hoje.strftime('%d/%m/%Y')}):<br><br>"

    if itens_hoje:
        corpo += "<b>📋 Hoje:</b><ul>" + "".join(f"<li>{linha(i)}</li>" for i in itens_hoje) + "</ul>"
    if itens_atrasados:
        corpo += "<br><b style='color:#C0392B'>⚠ Atrasado(s):</b><ul>"
        for i in itens_atrasados:
            corpo += f"<li>{linha(i)} — venceu em {i['data'][8:10]}/{i['data'][5:7]}</li>"
        corpo += "</ul>"
    if not itens_hoje and not itens_atrasados:
        corpo += "Nada pendente — tudo em dia! 🎉<br>"

    corpo += (f"<br>Marque como entregue direto aqui: "
              f"<a href='{COMPROMISSOS_URL}?pessoa={nome.replace(' ', '%20')}'>Meus compromissos</a><br><br>"
              f"Atenciosamente,<br>Controladoria — Peixoto e Cury Advogados")
    return corpo


def enviar_email(destinatario, assunto, corpo_html):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = destinatario
    mail.Subject = assunto
    mail.HTMLBody = corpo_html
    mail.Send()


def main():
    hoje = date.today()
    definicoes = buscar_definicoes_ativas()
    LOG.info(f"Total de definições ativas: {len(definicoes)}")

    upsert_pendencias_hoje(definicoes, hoje)

    pendencias = buscar_pendencias_em_aberto()
    por_pessoa = defaultdict(list)
    sem_email = []
    for p in pendencias:
        d = p["compromissos_definicoes"]
        if not d.get("responsavel_email"):
            sem_email.append(d.get("responsavel_nome"))
            continue
        por_pessoa[(d["responsavel_nome"], d["responsavel_email"])].append(p)

    if sem_email:
        LOG.warning(f"Pendências sem e-mail resolvido (não notificadas): {sorted(set(sem_email))}")

    LOG.info(f"Pessoas com pendência em aberto: {len(por_pessoa)}")
    for (nome, email), itens in por_pessoa.items():
        corpo = montar_corpo(nome, itens, hoje)
        destinatario_real = email
        assunto = f"📋 Seus compromissos de hoje (Time B e Controladoria) — {hoje.strftime('%d/%m/%Y')}"
        if MODO_TESTE:
            destinatario_real = TESTE_EMAIL
            assunto = f"[TESTE — seria p/ {nome} <{email}>] {assunto}"
        LOG.info(f"Enviando pra {destinatario_real} ({'MODO TESTE, real=' + email if MODO_TESTE else 'real'}) "
                 f"— {len(itens)} item(ns)")
        enviar_email(destinatario_real, assunto, corpo)

    LOG.info("Concluído.")


if __name__ == "__main__":
    main()
