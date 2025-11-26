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
FORM_RESPONSES_HINT = os.getenv("FORM_RESPONSES_HINT", "(Responses)")

SCOPES = [
    "User.Read",
    "Mail.Send",
    "Files.Read.All",
    "Sites.Read.All"
]


GRAPH = "https://graph.microsoft.com/v1.0"

INPUT_CSV = "ConvitesFormulario_IMPORT_MIN.csv"
TRACKING_CSV = "tracking.csv"
DAYS_DEADLINE = 7
SLEEP_SECONDS_BETWEEN_MAILS = 2


# -----------------------------
# AUTENTICAÇÃO
# -----------------------------
def get_token():
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}"
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


def graph_get(token, url, params=None):
    r = requests.get(
        url, headers={"Authorization": f"Bearer {token}"}, params=params)
    r.raise_for_status()
    return r.json()


def graph_post(token, url, payload):
    r = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}",
                 "Content-Type": "application/json"},
        data=json.dumps(payload)
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
            "toRecipients": [{"emailAddress": {"address": to_addr}}]
        },
        "saveToSentItems": True
    }
    graph_post(token, url, payload)


# -----------------------------
# BUSCAR ARQUIVO DE RESPOSTAS DO FORMS
# -----------------------------
def search_response_files(token, hint):
    found = []

    # OneDrive pessoal
    try:
        data = graph_get(token, f"{GRAPH}/me/drive/root/search(q='{hint}')")
        for item in data.get("value", []):
            if item.get("file"):
                found.append(
                    (item["parentReference"]["driveId"], item["id"], item["name"]))
    except:
        pass

    # Teams/Grupos
    try:
        groups = graph_get(token, f"{GRAPH}/me/joinedTeams")
        for g in groups.get("value", []):
            try:
                data = graph_get(
                    token, f"{GRAPH}/groups/{g['id']}/drive/root/search(q='{hint}')")
                for item in data.get("value", []):
                    if item.get("file"):
                        found.append(
                            (item["parentReference"]["driveId"], item["id"], item["name"]))
            except:
                continue
    except:
        pass

    # remover duplicados
    unique = {}
    for d, i, n in found:
        unique[(d, i)] = n
    return [(d, i, n) for (d, i), n in unique.items()]


def get_workbook_values(token, drive_id, item_id):
    sheets = graph_get(
        token, f"{GRAPH}/drives/{drive_id}/items/{item_id}/workbook/worksheets")
    ws_id = sheets["value"][0]["id"]
    rng = graph_get(
        token,
        f"{GRAPH}/drives/{drive_id}/items/{item_id}/workbook/worksheets/{ws_id}/usedRange(valuesOnly=true)"
    )
    return rng.get("values", [])


def extract_emails(matrix):
    if not matrix or len(matrix) < 2:
        return set()

    header = [str(h).strip() for h in matrix[0]]
    rows = matrix[1:]

    # tentar achar a coluna EMAIL primeiro
    try:
        col = header.index("Email")
    except ValueError:
        # fallback - pega a primeira coluna que contém a palavra email
        candidates = [i for i, h in enumerate(header) if "email" in h.lower()]
        if not candidates:
            return set()
        col = candidates[0]

    emails = set()
    for r in rows:
        if len(r) > col:
            val = str(r[col]).strip().lower()
            if "@" in val and "." in val:
                emails.add(val)
    return emails


# -----------------------------
# CSV DE TRACKING
# -----------------------------
def load_csv_recipients():
    out = []
    with open(INPUT_CSV, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            title = row["Title"].strip()
            email = row["Email"].strip().lower()
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
    fields = ["Title", "Email", "sent_at_iso", "due_at_iso",
              "responded_at_iso", "reminder_sent_at_iso"]
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
                "reminder_sent_at_iso": ""
            }
    return list(idx.values())


# -----------------------------
# TAREFA: ENVIAR CONVITES
# -----------------------------
def task_send(token, subject, form_link):
    rec = load_csv_recipients()
    trk = merge_tracking(rec, load_tracking())

    now = dt.datetime.utcnow()

    sent = 0
    for row in trk:
        if not row["sent_at_iso"]:
            html = f"""
            Olá, {row['Title']}<br><br>
            Precisamos de sua atualização cadastral.<br>
            <a href="{form_link}">Clique aqui para preencher</a><br><br>
            Obrigado!
            """
            send_mail(token, row["Email"], subject, html)
            row["sent_at_iso"] = now.isoformat()+"Z"
            row["due_at_iso"] = (
                now + dt.timedelta(days=DAYS_DEADLINE)).isoformat()+"Z"
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

    # buscar arquivos de resposta
    files = search_response_files(token, FORM_RESPONSES_HINT)
    answered = set()

    for d, i, n in files:
        values = get_workbook_values(token, d, i)
        answered |= extract_emails(values)

    # marcar respondidos
    for row in trk:
        if not row["responded_at_iso"] and row["Email"] in answered:
            row["responded_at_iso"] = now.isoformat()+"Z"

    save_tracking(trk)

    # enviar lembrete
    reminded = 0
    for row in trk:
        if row["responded_at_iso"]:
            continue
        if not row["due_at_iso"]:
            continue
        if row["reminder_sent_at_iso"]:
            continue

        due = dt.datetime.fromisoformat(row["due_at_iso"].replace("Z", ""))
        if now >= due:
            html = f"""
            Olá, {row['Title']}<br><br>
            Não registramos sua resposta no formulário.<br>
            <a href="{form_link}">Clique aqui para responder</a><br><br>
            Obrigado!
            """
            send_mail(token, row["Email"], subject_reminder, html)
            row["reminder_sent_at_iso"] = now.isoformat()+"Z"
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
        "--subject", default="Lembrete: Atualização cadastral pendente")
    p_check.add_argument("--form-link", required=True)

    args = parser.parse_args()

    token = get_token()

    if args.cmd == "send":
        task_send(token, args.subject, args.form_link)
    else:
        task_check(token, args.subject, args.form_link)
