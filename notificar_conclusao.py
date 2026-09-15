# -*- coding: utf-8 -*-
"""
notificar_conclusao.py — LAWgico Painel de Tarefas

Roda a cada poucos minutos (Tarefa Agendada do Windows, ex: a cada 15 min,
dias uteis, horario comercial): checa quem marcou tarefa(s) como "concluido"
desde a ultima checagem e, pra cada AUTOMACAO ativa cujo escopo bater com o
item concluido, manda 1 e-mail resumindo pro destinatario da regra via
automacao do Outlook (.Send() de verdade, ninguem precisa clicar em nada).

Automacoes ficam na tabela compromissos_automacoes (criada/editada pela tela
"Automacoes" do index.html): cada regra tem escopo_tipo ('todos' | 'grupo' |
'pessoa'), escopo_valor e email_notificar. Uma regra 'todos' pega qualquer
conclusao; 'grupo'/'pessoa' so pegam itens cujo grupo/responsavel bate. Uma
mesma conclusao pode disparar varias regras (e-mails diferentes).

Guarda o timestamp da ultima checagem em ULTIMO_CHECK_PATH (arquivo local,
fora do git) pra nunca notificar a mesma conclusao duas vezes nem perder
uma entre execucoes.

Se nada foi concluido desde a ultima checagem (ou nenhuma regra ativa bate
com o que foi concluido), nao manda e-mail nenhum.
"""
import json
import logging
import sys
from collections import defaultdict
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

ULTIMO_CHECK_PATH = Path(__file__).parent / "_ultimo_check_conclusao.json"

LOG = logging.getLogger("notificar_conclusao")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


def ler_ultimo_check():
    if not ULTIMO_CHECK_PATH.exists():
        # Primeira execucao: nao varre o historico inteiro, comeca a contar a
        # partir de agora (senao manda 1 e-mail gigante com tudo que ja foi
        # concluido desde sempre).
        return datetime.now(timezone.utc).isoformat()
    return json.loads(ULTIMO_CHECK_PATH.read_text(encoding="utf-8"))["ultimo_check"]


def salvar_ultimo_check(timestamp_iso):
    ULTIMO_CHECK_PATH.write_text(json.dumps({"ultimo_check": timestamp_iso}), encoding="utf-8")


def buscar_concluidos_desde(ultimo_check):
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/compromissos_entregas",
        headers=HEADERS, timeout=30,
        params={
            "status": "eq.concluido",
            "entregue_em": f"gt.{ultimo_check}",
            "select": "*,compromissos_definicoes(*)",
            "order": "entregue_em.asc",
        },
    )
    r.raise_for_status()
    return [e for e in r.json() if e.get("compromissos_definicoes")]


def buscar_automacoes_ativas():
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/compromissos_automacoes",
        headers=HEADERS, timeout=30,
        params={"ativo": "eq.true", "select": "*"},
    )
    r.raise_for_status()
    return r.json()


def automacao_bate(regra, definicao):
    tipo = regra.get("escopo_tipo")
    if tipo == "todos":
        return True
    if tipo == "grupo":
        return (definicao.get("grupo") or "").strip() == (regra.get("escopo_valor") or "").strip()
    if tipo == "pessoa":
        return (definicao.get("responsavel_nome") or "").strip() == (regra.get("escopo_valor") or "").strip()
    return False


def montar_corpo(itens, nome_regra):
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

    corpo = f"Automação <b>{nome_regra}</b> — tarefas marcadas como <b style='color:#117A65'>concluídas</b> desde a última checagem:<br><br>"
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

    automacoes = buscar_automacoes_ativas()
    LOG.info(f"Automações ativas: {len(automacoes)}")

    if not automacoes:
        salvar_ultimo_check(agora)
        LOG.info("Nenhuma automação ativa cadastrada. Nada enviado (rode add_automacoes.sql se ainda não rodou).")
        return

    # Agrupa: pra cada regra que bater, acumula os itens que ela deve listar.
    por_regra = defaultdict(list)
    for item in concluidos:
        d = item["compromissos_definicoes"]
        for regra in automacoes:
            if automacao_bate(regra, d):
                por_regra[regra["id"]].append(item)

    if not por_regra:
        salvar_ultimo_check(agora)
        LOG.info("Nenhuma automação bateu com os itens concluídos. Nada enviado.")
        return

    regras_por_id = {r["id"]: r for r in automacoes}
    for regra_id, itens in por_regra.items():
        regra = regras_por_id[regra_id]
        corpo = montar_corpo(itens, regra["nome"])
        assunto = f"✅ {len(itens)} tarefa(s) concluída(s) — {regra['nome']}"
        LOG.info(f"Enviando pra {regra['email_notificar']} (regra '{regra['nome']}') — {len(itens)} item(ns)")
        enviar_email(regra["email_notificar"], assunto, corpo)

    salvar_ultimo_check(agora)
    LOG.info("Concluído.")


if __name__ == "__main__":
    main()
