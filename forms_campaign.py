import os
import csv
import time
import json
import argparse
import datetime as dt
from typing import List, Dict, Set
import requests
import msal
from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")

SCOPES = [
    "User.Read",
    "Mail.Send",
    "Files.Read.All",
    "Sites.Read.All",
]

GRAPH = "https://graph.microsoft.com/v1.0"

INPUT_CSV = "ConvitesFormulario_IMPORT_MIN.csv"
TRACKING_CSV = "tracking.csv"
RESPONSES_CSV = "respostas_forms.csv"
DAYS_DEADLINE = 7
SLEEP_SECONDS_BETWEEN_MAILS = 3

# domínios que NÃO devem ser tratados como “corporativos”
GENERIC_DOMAINS = {
    "gmail.com", "gmail.com.br",
    "outlook.com", "outlook.com.br",
    "hotmail.com", "hotmail.com.br",
    "live.com", "live.com.br",
    "yahoo.com", "yahoo.com.br",
    "icloud.com",
    "bol.com.br", "uol.com.br"
}


# -----------------------------
# AUTENTICAÇÃO
# -----------------------------
def get_token():
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError("Falha ao autenticar: ", result)
    return result["access_token"]


def graph_post(token, url, payload):
    r = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        data=json.dumps(payload),
    )
    if r.status_code >= 400:
        raise RuntimeError(r.text)
    return r.json() if r.text else {}


# -----------------------------
# ENVIO DE EMAIL
# -----------------------------
def send_mail(token, to_addr, subject, html_body):
    url = f"{GRAPH}/me/sendMail"
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": to_addr}}],
        },
        "saveToSentItems": True,
    }
    graph_post(token, url, payload)


# -----------------------------
# CSV DE TRACKING E LISTA
# -----------------------------
def load_csv_recipients():
    out = []
    with open(INPUT_CSV, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            title = (row.get("Title") or "").strip()
            email = (row.get("Email") or "").strip().lower()
            if title and email:
                out.append({"Title": title, "Email": email})
    return out


def load_tracking():
    if not os.path.exists(TRACKING_CSV):
        return []
    out = []
    with open(TRACKING_CSV, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            out.append(row)
    return out


def save_tracking(rows):
    fields = [
        "Title",
        "Email",
        "sent_at_iso",
        "due_at_iso",
        "responded_at_iso",
        "reminder_sent_at_iso",
    ]
    with open(TRACKING_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fields})


def merge_tracking(recipients, tracking):
    idx = {r["Email"]: r for r in tracking}
    for rec in recipients:
        if rec["Email"] not in idx:
            idx[rec["Email"]] = {
                "Title": rec["Title"],
                "Email": rec["Email"],
                "sent_at_iso": "",
                "due_at_iso": "",
                "responded_at_iso": "",
                "reminder_sent_at_iso": "",
            }
    return list(idx.values())


# -----------------------------
# LER CSV DE RESPOSTAS DO FORMS
# -----------------------------
def load_responses_from_csv(path: str):
    """
    Lê o CSV exportado do Forms e devolve:
      - set de e-mails que responderam
      - set de domínios que já têm alguém respondido
    Usa a coluna 'Informe um E-mail para contato' como prioridade.
    """
    if not os.path.exists(path):
        print(f"[Aviso] Arquivo de respostas '{path}' não encontrado.")
        return set(), set()

    with open(path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []

        colname = None

        # 1) prioriza exatamente "Informe um E-mail para contato"
        for h in headers:
            if not h:
                continue
            if h.strip().lower() == "informe um e-mail para contato":
                colname = h
                break

        # 2) se não achar, cai para qualquer coluna com "email" no nome
        if colname is None:
            for h in headers:
                if h and "email" in h.lower():
                    colname = h
                    break

        if colname is None:
            print("[Aviso] Nenhuma coluna de e-mail encontrada no CSV de respostas.")
            return set(), set()

        emails = set()
        domains = set()

        for row in reader:
            raw = (row.get(colname) or "").strip().lower()
            if "@" in raw and "." in raw:
                emails.add(raw)
                dom = raw.split("@")[-1]
                domains.add(dom)

        return emails, domains


def get_domains_from_tracking(trk_rows):
    domains = set()
    for row in trk_rows:
        email = (row.get("Email") or "").strip().lower()
        if "@" in email:
            dom = email.split("@")[-1]
            domains.add(dom)
    return domains


# -----------------------------
# TAREFA: ENVIAR CONVITES
# -----------------------------
def task_send(token, subject, form_link):
    rec = load_csv_recipients()
    trk = merge_tracking(rec, load_tracking())

    now = dt.datetime.utcnow()

    sent = 0
    for row in trk:
        if not row.get("sent_at_iso"):
            html = f"""
            <p>Prezados, <strong>{row['Title']}</strong>,</p>
            
            <p>Em virtude da <strong>Reforma Tributária</strong> em andamento no Brasil, estamos atualizando nosso cadastro de fornecedores para garantir a conformidade com as novas exigências fiscais.</p>
            
            <p>Solicitamos que preencha o formulário disponível no link abaixo com as informações atualizadas da sua empresa:</p>
            
            <p style="text-align: center; margin: 20px 0;">
                <a href="{form_link}" style="background-color: #0078D4; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; display: inline-block; font-weight: bold;">Preencher Formulário</a>
            </p>
            
            <p>Para mais informações sobre a Reforma Tributária, acesse:<br>
            <a href="https://www.gov.br/fazenda/pt-br/acesso-a-informacao/acoes-e-programas/reforma-tributaria">https://www.gov.br/fazenda/pt-br/acesso-a-informacao/acoes-e-programas/reforma-tributaria</a></p>
            
            <p>Contamos com sua colaboração para mantermos nossos registros atualizados.</p>
            
            <p>Atenciosamente,<br>
            <strong>Statomat Máquinas Especiais</strong></p>
            """
            send_mail(token, row["Email"], subject, html)
            row["sent_at_iso"] = now.isoformat() + "Z"
            row["due_at_iso"] = (
                now + dt.timedelta(days=DAYS_DEADLINE)
            ).isoformat() + "Z"
            save_tracking(trk)
            sent += 1
            time.sleep(SLEEP_SECONDS_BETWEEN_MAILS)

    print("Convites enviados:", sent)


# -----------------------------
# TAREFA: VERIFICAR RESPOSTAS + LEMBRETES
# -----------------------------
def task_check(token, subject_reminder, form_link):
    trk = load_tracking()
    now = dt.datetime.utcnow()

    # 1) Lê CSV de respostas exportado do Forms
    answered_emails, answered_domains = load_responses_from_csv(RESPONSES_CSV)

    print("E-mails respondidos encontrados no CSV:")
    for e in sorted(answered_emails):
        print(" -", e)

    print("Domínios com pelo menos 1 resposta:")
    for d in sorted(answered_domains):
        print(" -", d)

    # 2) Descobre automaticamente quais domínios são "corporativos"
    all_sent_domains = get_domains_from_tracking(trk)
    corporate_domains = {
        d for d in all_sent_domains if d not in GENERIC_DOMAINS
    }

    print("Domínios corporativos detectados (a partir da sua lista):")
    for d in sorted(corporate_domains):
        print(" -", d)

    # 3) Marca respondidos:
    #    - se o e-mail exato respondeu
    #    - OU se o domínio é corporativo e alguém desse domínio respondeu
    for row in trk:
        if row.get("responded_at_iso"):
            continue  # já marcado antes

        email = (row.get("Email") or "").strip().lower()
        if not email:
            continue

        domain = email.split("@")[-1] if "@" in email else ""

        # caso 1: match exato do e-mail
        if email in answered_emails:
            row["responded_at_iso"] = now.isoformat() + "Z"
            continue

        # caso 2: domínio corporativo inteiro válido
        if domain in corporate_domains and domain in answered_domains:
            row["responded_at_iso"] = now.isoformat() + "Z"
            continue

    save_tracking(trk)

    # 4) Envia lembretes IMEDIATAMENTE para quem não respondeu
    reminded = 0
    for row in trk:
        email = (row.get("Email") or "").strip().lower()
        if not email:
            continue

        # Se já respondeu, pula
        if row.get("responded_at_iso"):
            continue

        # Envia lembrete imediatamente para quem não respondeu
        html = f"""
<p>Prezados, <strong>{row.get('Title', '')}</strong>,</p>

<p>Este é um <strong>lembrete</strong> sobre a atualização cadastral solicitada anteriormente.</p>

<p>Até o momento, <strong>não identificamos sua resposta</strong> ao formulário de atualização de dados relacionado à Reforma Tributária.</p>

<p>Para facilitar, disponibilizamos novamente o link do formulário:</p>

<p style="text-align: center; margin: 20px 0;">
    <a href="{form_link}" style="background-color: #D83B01; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; display: inline-block; font-weight: bold;">Preencher Formulário Agora</a>
</p>

<p>Sua colaboração é fundamental para mantermos nosso cadastro em conformidade com as novas normas fiscais.</p>

<p>Atenciosamente,<br>
<strong>Statomat Máquinas Especiais</strong></p>

<hr style="margin-top: 30px; border: none; border-top: 1px solid #ccc;">

<p style="font-size: 12px; color: #666; font-style: italic;">
Se você já respondeu ao formulário, por favor desconsidere esta mensagem!
</p>
"""

        print(f"Enviando lembrete para: {email}")
        send_mail(token, email, subject_reminder, html)

        row["reminder_sent_at_iso"] = now.isoformat() + "Z"
        save_tracking(trk)
        reminded += 1
        time.sleep(SLEEP_SECONDS_BETWEEN_MAILS)

    print("Lembretes enviados:", reminded)


# -----------------------------
# CLI
# -----------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_send = sub.add_parser("send")
    p_send.add_argument("--subject", required=True)
    p_send.add_argument("--form-link", required=True)

    p_check = sub.add_parser("check")
    p_check.add_argument(
        "--subject", default="Lembrete: Atualização cadastral pendente"
    )
    p_check.add_argument("--form-link", required=True)

    args = parser.parse_args()

    token = get_token()

    if args.cmd == "send":
        task_send(token, args.subject, args.form_link)
    else:
        task_check(token, args.subject, args.form_link)
