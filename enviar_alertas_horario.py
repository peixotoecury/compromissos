# -*- coding: utf-8 -*-
"""
enviar_alertas_horario.py — LAWgico Compromissos

Dispara só os itens curados em Alertas.xlsx (origem='alerta', ver
sync_alertas.py) que têm o horário passado por --horario e que valem hoje.
Pensado pra rodar várias vezes ao dia via Tarefas Agendadas do Windows, uma
por horário distinto (ex: 09:15, 09:30, 10:00, 11:00, 12:00, 13:00, 14:00,
15:00, 16:00 — os horários hoje presentes em Alertas.xlsx).

Não interfere no catálogo geral (origem='atividade'/'pauta'/'relatorio') nem
no envio único da manhã (enviar_lembretes_diarios.py) — os dois convivem:
o daquele é o resumo diário completo, este é o "toque" pontual no horário
certo pra quem tem alerta curado.

Uso:
    python enviar_alertas_horario.py --horario 11:00

MODO_TESTE=True (padrão, mesmo esquema do enviar_lembretes_diarios.py):
manda tudo pra TESTE_EMAIL com o destinatário real no assunto, até a
usuária validar. Mudar pra False só depois de confirmar.
"""
import argparse
import logging
import sys
from collections import defaultdict
from datetime import date
import calendar

import requests
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

SUPABASE_URL = "https://rpibvjcnrseuugpkfmdj.supabase.co"
SUPABASE_KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJwaWJ2amNucnNldXVncGtmbWRqIiwi"
                "cm9sZSI6ImFub24iLCJpYXQiOjE3ODE1NTc3MTcsImV4cCI6MjA5NzEzMzcxN30.ecihol8JESMH7cgFSvWKIzp-OwoPRFqdK3aCDwpCeg8")
HEADERS = {"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}", "Content-Type": "application/json"}

COMPROMISSOS_URL = "https://peixotoecury.github.io/compromissos/"  # atualizar apos publicar o repo

# ── MODO TESTE — validado pela usuaria em 21/08/2026, envio real ativado ──
MODO_TESTE = False
TESTE_EMAIL = "claude.controladoria@peixotoecury.com.br"

LOG = logging.getLogger("enviar_alertas_horario")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


def ultimo_dia_do_mes(ano, mes):
    return calendar.monthrange(ano, mes)[1]


def vale_hoje(d, hoje):
    tipo = d.get("regra_tipo")
    if tipo == "diaria":
        return True
    if tipo == "dias_semana":
        return hoje.weekday() in (d.get("regra_dias_semana") or [])
    if tipo == "dia_mes":
        dia_alvo = d.get("regra_dia_mes")
        if dia_alvo is None:
            return False
        meses = d.get("regra_meses")
        if meses and hoje.month not in meses:
            return False
        dia_alvo_ajustado = min(dia_alvo, ultimo_dia_do_mes(hoje.year, hoje.month))
        return hoje.day == dia_alvo_ajustado
    return False


def buscar_definicoes(horario):
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/compromissos_definicoes"
        f"?origem=eq.alerta&ativo=eq.true&horario=eq.{horario}&select=*",
        headers=HEADERS, timeout=30,
    )
    r.raise_for_status()
    return r.json()


def upsert_pendencias_hoje(definicoes_hoje, hoje):
    if not definicoes_hoje:
        return
    linhas = [{"definicao_id": d["id"], "data": hoje.isoformat(), "status": "pendente"} for d in definicoes_hoje]
    r = requests.post(
        f"{SUPABASE_URL}/rest/v1/compromissos_entregas",
        headers={**HEADERS, "Prefer": "resolution=ignore-duplicates,return=minimal"},
        json=linhas, timeout=30,
    )
    r.raise_for_status()


def montar_corpo(nome, itens, horario):
    def linha(d):
        partes = [d.get("item") or ""]
        if d.get("cliente"):
            partes.append(f"— {d['cliente']}")
        return " ".join(partes)

    corpo = f"Olá, {nome.split(' ')[0]}!<br><br>"
    corpo += f"Alerta das {horario} de hoje:<br><br><ul>"
    corpo += "".join(f"<li>{linha(d)}</li>" for d in itens)
    corpo += "</ul>"
    corpo += (f"<br>Marque como entregue direto aqui: "
              f"<a href='{COMPROMISSOS_URL}?pessoa={nome.replace(' ', '%20')}'>Meus compromissos</a><br><br>"
              f"Atenciosamente,<br>Controladoria — Peixoto e Cury Advogados")
    return corpo


def enviar_email(destinatario, assunto, corpo_html):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = destinatario
    mail.Subject = assunto
    mail.HTMLBody = corpo_html
    mail.Send()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--horario", required=True, help="Formato HH:MM, ex: 11:00")
    args = parser.parse_args()

    hoje = date.today()
    definicoes = buscar_definicoes(args.horario)
    LOG.info(f"Definições com horario={args.horario}: {len(definicoes)}")

    validas_hoje = [d for d in definicoes if vale_hoje(d, hoje)]
    LOG.info(f"Válidas hoje ({hoje.isoformat()}): {len(validas_hoje)}")
    if not validas_hoje:
        LOG.info("Nada pra disparar nesse horário hoje. Concluído.")
        return

    upsert_pendencias_hoje(validas_hoje, hoje)

    por_pessoa = defaultdict(list)
    sem_email = []
    for d in validas_hoje:
        if not d.get("responsavel_email"):
            sem_email.append(d.get("responsavel_nome"))
            continue
        por_pessoa[(d["responsavel_nome"], d["responsavel_email"])].append(d)

    if sem_email:
        LOG.warning(f"Sem e-mail resolvido (não notificados): {sorted(set(sem_email))}")

    for (nome, email), itens in por_pessoa.items():
        corpo = montar_corpo(nome, itens, args.horario)
        destinatario_real = email
        assunto = f"⏰ Alerta {args.horario} — {hoje.strftime('%d/%m/%Y')}"
        if MODO_TESTE:
            destinatario_real = TESTE_EMAIL
            assunto = f"[TESTE — seria p/ {nome} <{email}>] {assunto}"
        LOG.info(f"Enviando pra {destinatario_real} ({'MODO TESTE, real=' + email if MODO_TESTE else 'real'}) "
                 f"— {len(itens)} item(ns)")
        enviar_email(destinatario_real, assunto, corpo)

    LOG.info("Concluído.")


if __name__ == "__main__":
    main()
