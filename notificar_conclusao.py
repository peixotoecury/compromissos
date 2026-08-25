# -*- coding: utf-8 -*-
"""
notificar_conclusao.py — LAWgico Compromissos

Roda a cada poucos minutos (Tarefa Agendada do Windows, ex: a cada 15 min,
dias uteis, horario comercial): checa quem marcou compromisso(s) como
"entregue" desde a ultima checagem e manda 1 e-mail resumindo pra
controladoria@peixotoecury.com.br via automacao do Outlook (.Send() de
verdade, ninguem precisa clicar em nada).

Guarda o timestamp da ultima checagem em ULTIMO_CHECK_PATH (arquivo local,
fora do git) pra nunca notificar a mesma conclusao duas vezes nem perder
uma entre execucoes.

Se nada foi concluido desde a ultima checagem, nao manda e-mail nenhum
(evita spam de "nada aconteceu").

MODO_TESTE=True (padrao): manda pra TESTE_EMAIL em vez do destinatario
real, com aviso no assunto -- usar assim ate a usuaria validar por alguns
dias. Depois, mudar pra False (mesmo fluxo dos outros scripts desta pasta).
"""
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path

import requests
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

SUPABASE_URL = "https://rpibvjcnrseuugpkfmdj.supabase.co"
SUPABASE_KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJwaWJ2amNucnNldXVncGtmbWRqIiwi"
                "cm9sZSI6ImFub24iLCJpYXQiOjE3ODE1NTc3MTcsImV4cCI6MjA5NzEzMzcxN30.ecihol8JESMH7cgFSvWKIzp-OwoPRFqdK3aCDwpCeg8")
HEADERS = {"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}", "Content-Type": "application/json"}

COMPROMISSOS_URL = "https://peixotoecury.github.io/compromissos/"
DESTINATARIO_REAL = "controladoria@peixotoecury.com.br"

# ── MODO TESTE — validado pela usuaria em 25/08/2026, envio real ativado ──
MODO_TESTE = False
TESTE_EMAIL = "claude.controladoria@peixotoecury.com.br"

ULTIMO_CHECK_PATH = Path(__file__).parent / "_ultimo_check_conclusao.json"

LOG = logging.getLogger("notificar_conclusao")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


def ler_ultimo_check():
    if not ULTIMO_CHECK_PATH.exists():
        # Primeira execucao: nao varre o historico inteiro, comeca a contar a
        # partir de agora (senao manda 1 e-mail gigante com tudo que ja foi
        # entregue desde sempre).
        return datetime.now(timezone.utc).isoformat()
    return json.loads(ULTIMO_CHECK_PATH.read_text(encoding="utf-8"))["ultimo_check"]


def salvar_ultimo_check(timestamp_iso):
    ULTIMO_CHECK_PATH.write_text(json.dumps({"ultimo_check": timestamp_iso}), encoding="utf-8")


def buscar_concluidos_desde(ultimo_check):
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/compromissos_entregas",
        headers=HEADERS, timeout=30,
        params={
            "status": "eq.entregue",
            "entregue_em": f"gt.{ultimo_check}",
            "select": "*,compromissos_definicoes(*)",
            "order": "entregue_em.asc",
        },
    )
    r.raise_for_status()
    return [e for e in r.json() if e.get("compromissos_definicoes")]


def montar_corpo(itens):
    def linha(i):
        d = i["compromissos_definicoes"]
        partes = [f"<b>{d.get('responsavel_nome') or '—'}</b>", "—", d.get("item") or ""]
        if d.get("etapa"):
            partes.append(f"({d['etapa']})")
        if d.get("cliente"):
            partes.append(f"· {d['cliente']}")
        hora = (i.get("entregue_em") or "")[11:16]
        partes.append(f"<span style='color:#7E98AA'>[{hora}]</span>" if hora else "")
        return " ".join(p for p in partes if p)

    corpo = "Compromissos marcados como <b style='color:#117A65'>entregues</b> desde a última checagem:<br><br>"
    corpo += "<ul>" + "".join(f"<li>{linha(i)}</li>" for i in itens) + "</ul>"
    corpo += (f"<br><a href='{COMPROMISSOS_URL}'>Ver painel completo</a><br><br>"
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
    ultimo_check = ler_ultimo_check()
    agora = datetime.now(timezone.utc).isoformat()

    concluidos = buscar_concluidos_desde(ultimo_check)
    LOG.info(f"Concluídos desde {ultimo_check}: {len(concluidos)}")

    if not concluidos:
        salvar_ultimo_check(agora)
        LOG.info("Nada novo. Concluído.")
        return

    corpo = montar_corpo(concluidos)
    destinatario_real = DESTINATARIO_REAL
    assunto = f"✅ {len(concluidos)} compromisso(s) concluído(s) — Compromissos"
    if MODO_TESTE:
        destinatario_real = TESTE_EMAIL
        assunto = f"[TESTE — seria p/ {DESTINATARIO_REAL}] {assunto}"

    LOG.info(f"Enviando pra {destinatario_real} "
             f"({'MODO TESTE, real=' + DESTINATARIO_REAL if MODO_TESTE else 'real'}) — {len(concluidos)} item(ns)")
    enviar_email(destinatario_real, assunto, corpo)

    salvar_ultimo_check(agora)
    LOG.info("Concluído.")


if __name__ == "__main__":
    main()
