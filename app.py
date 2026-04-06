"""
Online Version des Team Task Managers.
Speichert Daten in Google Sheets, Anhaenge in Google Drive.

Starten mit:
  streamlit run app.py
"""
import streamlit as st
import json
import os
import re
import urllib.parse
import base64
import io
from datetime import date, datetime

from google.oauth2.service_account import Credentials
import gspread
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# --- Translations (all keys from local version) --------------------

TRANSLATIONS = {
    "de": {
        "app_title": "Team Task Manager",
        "login_title": "Team Task Manager \u2014 Login",
        "setup_title": "Team Task Manager \u2014 Ersteinrichtung",
        "logout": "Ausloggen", "save": "Speichern", "saved": "Gespeichert!",
        "delete": "L\u00f6schen", "edit": "Bearbeiten", "save_changes": "\u00c4nderungen speichern",
        "cancel": "Abbrechen", "email": "Email", "password": "Passwort",
        "login": "Einloggen", "enter_email_pw": "Bitte Email und Passwort eingeben.",
        "no_member_found": "Kein Mitarbeiter mit dieser Email gefunden.",
        "wrong_password": "Falsches Passwort.",
        "setup_welcome": "Willkommen! Bitte lege den ersten Administrator an.",
        "your_name": "Dein Name *", "your_email": "Deine Email *",
        "create_admin": "Admin-Account erstellen", "admin_created": "Admin-Account erstellt!",
        "fill_all_fields": "Bitte alle Felder ausf\u00fcllen.",
        "my_team": "Mein Team", "invite_member": "Mitarbeiter einladen",
        "name": "Name", "invite": "Einladen",
        "email_exists": "Email existiert bereits.",
        "enter_name_email": "Bitte Name und Email eingeben.",
        "invited": "eingeladen! (Passwort: {pw})",
        "current_members": "Aktuelle Mitarbeiter",
        "change_password": "Passwort \u00e4ndern", "old_password": "Altes Passwort",
        "new_password": "Neues Passwort", "confirm_password": "Neues Passwort best\u00e4tigen",
        "change": "\u00c4ndern", "user_not_found": "User nicht gefunden.",
        "old_pw_wrong": "Altes Passwort ist falsch.",
        "enter_new_pw": "Bitte neues Passwort eingeben.",
        "pw_mismatch": "Neue Passw\u00f6rter stimmen nicht \u00fcberein.",
        "pw_changed": "Passwort ge\u00e4ndert!",
        "tab_new_task": "Neue Aufgabe", "tab_my_tasks": "Meine Aufgaben",
        "tab_team": "Aufgaben Team", "tab_import": "Import",
        "task_type": "Aufgabentyp",
        "team_task": "Team-Aufgabe (f\u00fcr Mitarbeiter)",
        "private_task": "Private Aufgabe (nur f\u00fcr mich)",
        "create_task": "Neue Aufgabe erstellen", "title": "Titel *",
        "description": "Beschreibung", "assign_to": "Zuweisen an *",
        "category": "Kategorie", "priority": "Priorit\u00e4t", "due_date": "F\u00e4llig am",
        "create_btn": "Aufgabe erstellen",
        "task_created": "Aufgabe '{title}' erstellt!",
        "private_created": "Private Aufgabe '{title}' erstellt!",
        "enter_title": "Bitte Titel eingeben.",
        "invite_first": "Bitte lade zuerst Mitarbeiter in der Sidebar ein.",
        "assigned_to_private": "Zugewiesen an: **{name}** (privat)",
        "no_tasks": "Du hast aktuell keine Aufgaben.",
        "status_filter": "Status-Filter", "category_filter": "Kategorie-Filter",
        "boss_tasks": "Aufgaben von Vorgesetzten ({count})",
        "private_tasks": "Meine privaten Aufgaben ({count})",
        "no_filtered": "Keine Aufgaben f\u00fcr die gew\u00e4hlten Filter.",
        "status": "Status", "comment": "Kommentar",
        "no_members": "Noch keine Mitarbeiter vorhanden.",
        "no_tasks_yet": "Noch keine Aufgaben vorhanden.",
        "tasks_count": "Aufgabe(n)",
        "send_email_to": "Aufgaben an {name} senden",
        "send_all_emails": "Alle Emails auf einmal \u00f6ffnen",
        "copy_email": "Email-Text kopieren",
        "import_title": "Email-Antwort importieren",
        "import_desc": "F\u00fcge die Antwort-Email eines Mitarbeiters hier ein.",
        "reply_from": "Antwort von Mitarbeiter",
        "paste_reply": "Email-Antwort hier einf\u00fcgen",
        "no_updates": "Keine Aufgaben-Updates erkannt.",
        "detected_changes": "Erkannte \u00c4nderungen",
        "task_not_found": "Aufgabe [{num}] existiert nicht.",
        "no_change": "keine \u00c4nderung",
        "unknown_status": "Unbekannter Status: '{status}'",
        "apply_changes": "{count} \u00c4nderung(en) \u00fcbernehmen",
        "changes_applied": "\u00c4nderungen \u00fcbernommen!",
        "no_valid_changes": "Keine g\u00fcltigen \u00c4nderungen.",
        "email_greeting": "Hallo {name},", "email_intro": "hier sind deine aktuellen Aufgaben:",
        "email_tasks_header": "--- AUFGABEN ---", "email_tasks_footer": "--- ENDE ---",
        "email_app_link": "Bitte aktualisiere deinen Fortschritt direkt in der App:",
        "email_reply_hint": "Oder antworte auf diese Email.",
        "email_closing": "Viele Gruesse",
        "email_subject": "Deine Aufgaben - {date}", "email_due": "Faellig: {date}",
        "email_desc": "Beschreibung: {desc}",
        "priority_high": "Hoch", "priority_medium": "Mittel", "priority_low": "Niedrig",
        "status_open": "Offen", "status_in_progress": "In Arbeit", "status_done": "Erledigt",
        "cat_marketing": "Marketing", "cat_it": "IT", "cat_purchasing": "Einkauf",
        "cat_sales": "Vertrieb", "cat_other": "Sonstiges",
        "search_placeholder": "Suche nach Aufgaben...",
        "archive": "Archivieren", "archived": "Archiviert",
        "show_archive": "Archiv anzeigen", "hide_archive": "Archiv ausblenden",
        "archive_empty": "Keine archivierten Aufgaben.",
        "task_archived": "Aufgabe archiviert!",
        "delete_member_title": "Mitarbeiter entfernen",
        "delete_member_confirm": "Passwort eingeben um {name} zu entfernen:",
        "delete_member_btn": "Entfernen best\u00e4tigen",
        "delete_member_wrong_pw": "Falsches Passwort.",
        "delete_member_success": "{name} wurde entfernt.",
        "attachments": "Anh\u00e4nge", "upload_files": "Dateien hochladen",
        "no_attachments": "Keine Anh\u00e4nge",
        "file_uploaded": "Datei '{name}' hochgeladen!", "open_file": "\u00d6ffnen",
        "hint_move_up": "Aufgabe nach oben verschieben",
        "hint_move_down": "Aufgabe nach unten verschieben",
        "hint_edit": "Aufgabe bearbeiten",
        "hint_archive": "Aufgabe archivieren",
        "remove_attachment": "Anhang entfernen",
    },
    "en": {
        "app_title": "Team Task Manager",
        "login_title": "Team Task Manager \u2014 Login",
        "setup_title": "Team Task Manager \u2014 Initial Setup",
        "logout": "Logout", "save": "Save", "saved": "Saved!",
        "delete": "Delete", "edit": "Edit", "save_changes": "Save changes",
        "cancel": "Cancel", "email": "Email", "password": "Password",
        "login": "Login", "enter_email_pw": "Please enter email and password.",
        "no_member_found": "No member found.", "wrong_password": "Wrong password.",
        "setup_welcome": "Welcome! Create the first administrator.",
        "your_name": "Your name *", "your_email": "Your email *",
        "create_admin": "Create admin", "admin_created": "Admin created!",
        "fill_all_fields": "Please fill all fields.",
        "my_team": "My Team", "invite_member": "Invite member",
        "name": "Name", "invite": "Invite", "email_exists": "Email exists.",
        "enter_name_email": "Enter name and email.",
        "invited": "invited! (Password: {pw})",
        "current_members": "Current members",
        "change_password": "Change password", "old_password": "Old password",
        "new_password": "New password", "confirm_password": "Confirm password",
        "change": "Change", "user_not_found": "User not found.",
        "old_pw_wrong": "Wrong old password.", "enter_new_pw": "Enter new password.",
        "pw_mismatch": "Passwords don't match.", "pw_changed": "Password changed!",
        "tab_new_task": "New Task", "tab_my_tasks": "My Tasks",
        "tab_team": "Team Tasks", "tab_import": "Import",
        "task_type": "Task type", "team_task": "Team task",
        "private_task": "Private task",
        "create_task": "Create task", "title": "Title *",
        "description": "Description", "assign_to": "Assign to *",
        "category": "Category", "priority": "Priority", "due_date": "Due date",
        "create_btn": "Create task", "task_created": "Task '{title}' created!",
        "private_created": "Private task '{title}' created!",
        "enter_title": "Enter a title.", "invite_first": "Invite members first.",
        "assigned_to_private": "Assigned to: **{name}** (private)",
        "no_tasks": "No tasks.", "status_filter": "Status filter",
        "category_filter": "Category filter",
        "boss_tasks": "Tasks from manager ({count})",
        "private_tasks": "My private tasks ({count})",
        "no_filtered": "No tasks match filters.", "status": "Status", "comment": "Comment",
        "no_members": "No members.", "no_tasks_yet": "No tasks.",
        "tasks_count": "task(s)", "send_email_to": "Send tasks to {name}",
        "send_all_emails": "Open all emails",
        "copy_email": "Copy email text",
        "import_title": "Import reply", "import_desc": "Paste email reply here.",
        "reply_from": "Reply from", "paste_reply": "Paste reply",
        "no_updates": "No updates.", "detected_changes": "Detected changes",
        "task_not_found": "Task [{num}] not found.", "no_change": "no change",
        "unknown_status": "Unknown status: '{status}'",
        "apply_changes": "Apply {count} change(s)", "changes_applied": "Applied!",
        "no_valid_changes": "No valid changes.",
        "email_greeting": "Hello {name},", "email_intro": "your current tasks:",
        "email_tasks_header": "--- TASKS ---", "email_tasks_footer": "--- END ---",
        "email_app_link": "Update progress in the app:",
        "email_reply_hint": "Or reply to this email.",
        "email_closing": "Best regards",
        "email_subject": "Your tasks - {date}", "email_due": "Due: {date}",
        "email_desc": "Description: {desc}",
        "priority_high": "High", "priority_medium": "Medium", "priority_low": "Low",
        "status_open": "Open", "status_in_progress": "In Progress", "status_done": "Done",
        "cat_marketing": "Marketing", "cat_it": "IT", "cat_purchasing": "Purchasing",
        "cat_sales": "Sales", "cat_other": "Other",
        "search_placeholder": "Search tasks...",
        "archive": "Archive", "archived": "Archived",
        "show_archive": "Show archive", "hide_archive": "Hide archive",
        "archive_empty": "No archived tasks.", "task_archived": "Archived!",
        "delete_member_title": "Remove member",
        "delete_member_confirm": "Enter password to remove {name}:",
        "delete_member_btn": "Confirm removal",
        "delete_member_wrong_pw": "Wrong password.",
        "delete_member_success": "{name} removed.",
        "attachments": "Attachments", "upload_files": "Upload files",
        "no_attachments": "No attachments",
        "file_uploaded": "File '{name}' uploaded!", "open_file": "Open",
        "hint_move_up": "Move task up",
        "hint_move_down": "Move task down",
        "hint_edit": "Edit task",
        "hint_archive": "Archive task",
        "remove_attachment": "Remove attachment",
    },
}


def get_lang():
    return st.session_state.get("lang", "de")

def t(key, **kwargs):
    text = TRANSLATIONS.get(get_lang(), TRANSLATIONS["de"]).get(key, key)
    return text.format(**kwargs) if kwargs else text

def get_priority_options():
    return [t("priority_high"), t("priority_medium"), t("priority_low")]
def get_status_options():
    return [t("status_open"), t("status_in_progress"), t("status_done")]
def get_category_options():
    return [t("cat_marketing"), t("cat_it"), t("cat_purchasing"), t("cat_sales"), t("cat_other")]

PRIORITY_KEYS = ["Hoch", "Mittel", "Niedrig"]
STATUS_KEYS = ["Offen", "In Arbeit", "Erledigt"]
CATEGORY_KEYS = ["Marketing", "IT", "Einkauf", "Vertrieb", "Sonstiges"]
STATUS_COLORS_MAP = {"Offen": "gray", "In Arbeit": "blue", "Erledigt": "green"}
CATEGORY_COLORS_MAP = {"Marketing": "#E91E63", "IT": "#2196F3", "Einkauf": "#FF9800", "Vertrieb": "#4CAF50", "Sonstiges": "#9E9E9E"}

def priority_to_key(d):
    o = get_priority_options(); return PRIORITY_KEYS[o.index(d)] if d in o else PRIORITY_KEYS[0]
def priority_to_display(k):
    return get_priority_options()[PRIORITY_KEYS.index(k)] if k in PRIORITY_KEYS else get_priority_options()[0]
def status_to_key(d):
    o = get_status_options(); return STATUS_KEYS[o.index(d)] if d in o else STATUS_KEYS[0]
def status_to_display(k):
    return get_status_options()[STATUS_KEYS.index(k)] if k in STATUS_KEYS else get_status_options()[0]
def category_to_key(d):
    o = get_category_options(); return CATEGORY_KEYS[o.index(d)] if d in o else CATEGORY_KEYS[0]
def category_to_display(k):
    return get_category_options()[CATEGORY_KEYS.index(k)] if k in CATEGORY_KEYS else get_category_options()[0]

DEFAULT_PASSWORD = "12345"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# --- Logo + Lang toggle (from local version) -------------------------

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")

def render_logo():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()
        st.markdown(f'<div style="margin-bottom:10px;"><img src="data:image/png;base64,{logo_b64}" width="250"></div>', unsafe_allow_html=True)

def render_lang_toggle():
    lang = get_lang()
    _, c2 = st.columns([6, 2])
    with c2:
        nl = st.selectbox("\U0001f310", ["Deutsch", "English"], index=0 if lang == "de" else 1, key="lang_sel", label_visibility="collapsed")
        nc = "de" if nl == "Deutsch" else "en"
        if nc != lang:
            st.session_state.lang = nc
            st.rerun()

# --- Google Sheets + Drive --------------------------------------------

@st.cache_resource
def get_gsheet_connection():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(st.secrets["spreadsheet"]["key"])
    return spreadsheet

@st.cache_resource
def get_drive_service():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)

def get_attachments_folder_id():
    return st.secrets.get("drive", {}).get("folder_id", "")

def upload_to_drive(drive_service, file_bytes, filename, mime_type, folder_id):
    file_meta = {"name": filename, "parents": [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
    uploaded = drive_service.files().create(body=file_meta, media_body=media, fields="id,webViewLink").execute()
    return {"id": uploaded["id"], "name": filename, "link": uploaded.get("webViewLink", "")}

def get_app_url():
    return st.secrets.get("app", {}).get("url", "")

# --- Worksheet headers + migration ------------------------------------

MEMBER_HEADERS = ["name", "email", "password", "role", "invited_by"]
TASK_HEADERS = [
    "title", "description", "assigned_to", "priority",
    "due", "created", "status", "comment", "category", "sort_order",
    "created_by", "private", "archived", "attachments",
]

def _migrate_worksheet(spreadsheet, ws_name, expected_headers):
    ws = spreadsheet.worksheet(ws_name)
    current_headers = ws.row_values(1)
    missing = [h for h in expected_headers if h not in current_headers]
    if not missing:
        return
    new_total = len(current_headers) + len(missing)
    if ws.col_count < new_total:
        ws.resize(cols=new_total)
    start_col = len(current_headers) + 1
    for i, header in enumerate(missing):
        ws.update_cell(1, start_col + i, header)

def ensure_worksheets(spreadsheet):
    existing = [ws.title for ws in spreadsheet.worksheets()]
    if "members" not in existing:
        ws = spreadsheet.add_worksheet(title="members", rows=100, cols=len(MEMBER_HEADERS))
        ws.update("A1:E1", [MEMBER_HEADERS])
    else:
        _migrate_worksheet(spreadsheet, "members", MEMBER_HEADERS)
    if "tasks" not in existing:
        ws = spreadsheet.add_worksheet(title="tasks", rows=500, cols=len(TASK_HEADERS))
        ws.update("A1:N1", [TASK_HEADERS])
    else:
        _migrate_worksheet(spreadsheet, "tasks", TASK_HEADERS)

# --- Load / Save ------------------------------------------------------

def load_members(spreadsheet):
    ws = spreadsheet.worksheet("members")
    records = ws.get_all_records()
    for m in records:
        m.setdefault("password", DEFAULT_PASSWORD)
        m.setdefault("role", "member")
        m.setdefault("invited_by", "")
        for k in m:
            m[k] = str(m[k]) if m[k] is not None else ""
    return records

def load_tasks(spreadsheet):
    ws = spreadsheet.worksheet("tasks")
    records = ws.get_all_records()
    for task in records:
        task.setdefault("status", "Offen")
        task.setdefault("comment", "")
        task.setdefault("category", "Sonstiges")
        task.setdefault("sort_order", 0)
        task.setdefault("created_by", "")
        task.setdefault("private", "")
        task.setdefault("archived", "")
        task.setdefault("attachments", "")
        for k in task:
            if k == "sort_order":
                try:
                    task[k] = int(task[k]) if task[k] else 0
                except (ValueError, TypeError):
                    task[k] = 0
            else:
                task[k] = str(task[k]) if task[k] is not None else ""
    return records

def save_members(spreadsheet, members):
    ws = spreadsheet.worksheet("members")
    ws.clear()
    ws.update("A1:E1", [MEMBER_HEADERS])
    if members:
        rows = [[m.get("name", ""), m.get("email", ""), m.get("password", DEFAULT_PASSWORD),
                 m.get("role", "member"), m.get("invited_by", "")] for m in members]
        ws.update(f"A2:E{len(rows)+1}", rows)

def save_tasks(spreadsheet, tasks):
    ws = spreadsheet.worksheet("tasks")
    ws.clear()
    ws.update("A1:N1", [TASK_HEADERS])
    if tasks:
        rows = []
        for task in tasks:
            att = task.get("attachments", "")
            if isinstance(att, list):
                att = json.dumps(att, ensure_ascii=False) if att else ""
            rows.append([
                task.get("title", ""), task.get("description", ""), task.get("assigned_to", ""),
                task.get("priority", ""), task.get("due", ""), task.get("created", ""),
                task.get("status", "Offen"), task.get("comment", ""), task.get("category", "Sonstiges"),
                task.get("sort_order", 0), task.get("created_by", ""), task.get("private", ""),
                task.get("archived", ""), att,
            ])
        ws.update(f"A2:N{len(rows)+1}", rows)

def get_task_attachments(task):
    raw = task.get("attachments", "")
    if not raw:
        return []
    if isinstance(raw, list):
        return raw
    try:
        return json.loads(raw)
    except (json.JSONDecodeError, TypeError):
        return []

def set_task_attachments(task, attachments):
    task["attachments"] = json.dumps(attachments, ensure_ascii=False) if attachments else ""

# --- Helper functions -------------------------------------------------

def get_my_team(members, ue):
    return [m for m in members if m.get("invited_by") == ue]
def get_assignable(members, ue, ur):
    return members if ur == "admin" else get_my_team(members, ue)
def get_visible(tasks, members, ue, ur):
    if ur == "admin":
        return [tk for tk in tasks if (tk.get("private") != "ja" or tk.get("created_by") == ue) and tk.get("archived") != "ja"]
    te = [m["email"] for m in get_my_team(members, ue)]
    return [tk for tk in tasks if tk["assigned_to"] in te and tk.get("private") != "ja" and tk.get("archived") != "ja"]
def get_archived(tasks, email):
    return [tk for tk in tasks if tk["assigned_to"] == email and tk.get("archived") == "ja"]
def search_tasks(tl, q):
    if not q: return tl
    q = q.lower()
    return [tk for tk in tl if q in tk.get("title", "").lower() or q in tk.get("description", "").lower() or q in tk.get("comment", "").lower()]

# --- Email ------------------------------------------------------------

def build_email_body(name, tl, app_url=""):
    lines = [t("email_greeting", name=name), "", t("email_intro"), "", t("email_tasks_header")]
    for i, tk in enumerate(tl, 1):
        lines.append(f"[{i}] {tk['title']} [{priority_to_display(tk.get('priority','Mittel'))}]")
        if tk.get("description"): lines.append(f"    {tk['description']}")
        lines.append(f"    Status: {status_to_display(tk.get('status','Offen'))}")
        lines.append(f"    Kommentar: {tk.get('comment','')}")
        lines.append("")
    lines.append(t("email_tasks_footer"))
    lines.append("")
    if app_url:
        lines += [t("email_app_link"), app_url, ""]
    lines += [t("email_reply_hint"), "", t("email_closing")]
    return "\n".join(lines)

def build_mailto(email, name, tl, app_url=""):
    subj = t("email_subject", date=date.today().strftime("%d.%m.%Y"))
    body = build_email_body(name, tl, app_url)
    p = urllib.parse.urlencode({"subject": subj, "body": body}, quote_via=urllib.parse.quote)
    return f"mailto:{email}?{p}"

# --- CSS (from local version - card layout) ---------------------------

COMPACT_CSS = """<style>
.task-card{display:flex;padding:10px 14px;border:1px solid rgba(128,128,128,0.2);border-radius:8px;margin-bottom:6px;gap:16px;font-size:0.9em;}
.task-card:hover{background:rgba(128,128,128,0.05);}
.task-left{flex:1;min-width:0;text-align:left;}
.task-heading{font-weight:700;font-size:1.05em;margin-bottom:4px;}
.task-desc-text{font-size:0.85em;line-height:1.5;margin-top:4px;opacity:0.75;word-wrap:break-word;}
.task-comment-text{font-size:0.82em;margin-top:5px;font-style:italic;opacity:0.6;}
.task-right{min-width:160px;display:flex;flex-direction:column;gap:6px;align-items:flex-end;padding-top:2px;}
.task-info-row{display:flex;align-items:center;gap:6px;font-size:0.82em;justify-content:flex-end;}
.task-info-label{opacity:0.5;min-width:55px;text-align:right;}
.task-info-value{font-weight:600;}
.badge{display:inline-block;padding:3px 10px;border-radius:12px;color:white;font-size:0.78em;white-space:nowrap;font-weight:600;}
</style>"""

def _esc(text):
    """Escape HTML special characters."""
    return str(text).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;").replace("\n"," ")

def task_row_html(tk):
    sk, ck, pk = tk.get("status","Offen"), tk.get("category","Sonstiges"), tk.get("priority","Mittel")
    sc = {"Offen":"#9e9e9e","In Arbeit":"#2196F3","Erledigt":"#4CAF50"}.get(sk,"#9e9e9e")
    cc = CATEGORY_COLORS_MAP.get(ck,"#9E9E9E")
    pc = {"Hoch":"#f44336","Mittel":"#ff9800","Niedrig":"#4caf50"}.get(pk,"#9e9e9e")
    priv = "🔒 " if tk.get("private") == "ja" else ""
    title = _esc(tk.get("title", ""))
    desc = _esc(tk.get("description", ""))
    desc_html = f'<div class="task-desc-text">{desc}</div>' if desc.strip() else ""
    cmt = _esc(tk.get("comment", ""))
    cmt_html = f'<div class="task-comment-text">💬 {cmt}</div>' if cmt.strip() else ""
    att = len(get_task_attachments(tk))
    att_html = f'<div class="task-info-row"><span class="task-info-label">Anhänge</span><span class="task-info-value">📎 {att}</span></div>' if att else ""
    due = _esc(tk.get("due", ""))
    due_html = f'<div class="task-info-row"><span class="task-info-label">Fällig</span><span class="task-info-value">{due}</span></div>' if due.strip() else ""

    return (
        '<div class="task-card">'
        '<div class="task-left">'
        f'<div class="task-heading">{priv}{title}</div>'
        f'{desc_html}{cmt_html}'
        '</div>'
        '<div class="task-right">'
        f'<div class="task-info-row"><span class="task-info-label">Status</span><span class="badge" style="background:{sc}">{status_to_display(sk)}</span></div>'
        f'<div class="task-info-row"><span class="task-info-label">Kategorie</span><span class="badge" style="background:{cc}">{category_to_display(ck)}</span></div>'
        f'<div class="task-info-row"><span class="task-info-label">Priorität</span><span class="badge" style="background:{pc}">{priority_to_display(pk)}</span></div>'
        f'{due_html}{att_html}'
        '</div>'
        '</div>'
    )

TABLE_HEADER = ""

# --- Login / Setup ----------------------------------------------------

def show_setup(spreadsheet):
    render_logo(); render_lang_toggle()
    st.title(f"\U0001f4cb {t('setup_title')}"); st.info(t("setup_welcome"))
    with st.form("setup"):
        n = st.text_input(t("your_name")); e = st.text_input(t("your_email"))
        p = st.text_input(t("password")+" *", type="password", value=DEFAULT_PASSWORD)
        if st.form_submit_button(t("create_admin")):
            if n and e and p:
                save_members(spreadsheet, [{"name":n,"email":e,"password":p,"role":"admin","invited_by":""}])
                st.session_state.user = {"email":e,"name":n,"role":"admin"}
                st.success(t("admin_created"))
                st.rerun()
            else: st.warning(t("fill_all_fields"))

def show_login(members):
    render_logo(); render_lang_toggle()
    st.title(f"\U0001f4cb {t('login_title')}")
    with st.form("login"):
        e = st.text_input(t("email")); p = st.text_input(t("password"), type="password")
        if st.form_submit_button(t("login")):
            if not e or not p: st.warning(t("enter_email_pw")); return
            m = next((x for x in members if x["email"].lower() == e.lower()), None)
            if not m: st.error(t("no_member_found")); return
            if m.get("password", DEFAULT_PASSWORD) != p: st.error(t("wrong_password")); return
            st.session_state.user = {"email":m["email"],"name":m["name"],"role":m.get("role","member")}
            st.rerun()

# --- Sidebar ----------------------------------------------------------

def render_sidebar(spreadsheet, members, tasks, user):
    ue, ur = user["email"], user["role"]
    st.sidebar.write(f"\U0001f464 **{user['name']}**"); st.sidebar.caption(f"{ue} ({ur.capitalize()})")
    if st.sidebar.button(f"\U0001f6aa {t('logout')}"): del st.session_state.user; st.rerun()
    st.sidebar.divider(); st.sidebar.header(f"\U0001f465 {t('my_team')}")
    with st.sidebar.form("add_m", clear_on_submit=True):
        nn = st.text_input(t("name")); ne = st.text_input(t("email"))
        if st.form_submit_button(t("invite")):
            if nn and ne:
                if any(m["email"].lower() == ne.lower() for m in members): st.error(t("email_exists"))
                else:
                    members.append({"name":nn,"email":ne,"password":DEFAULT_PASSWORD,"role":"member","invited_by":ue})
                    save_members(spreadsheet, members); st.success(f"{nn} {t('invited',pw=DEFAULT_PASSWORD)}"); st.rerun()
            else: st.warning(t("enter_name_email"))
    team = get_my_team(members, ue) if ur != "admin" else members
    if team:
        st.sidebar.subheader(t("current_members"))
        for i, m in enumerate(team):
            if m["email"] == ue: continue
            mi = members.index(m)
            st.sidebar.write(f"{'\U0001f451' if m.get('role')=='admin' else ''} **{m['name']}** ({m['email']})")
            dk = f"dc_{mi}"
            if dk not in st.session_state: st.session_state[dk] = False
            if not st.session_state[dk]:
                if st.sidebar.button(f"\u2715", key=f"db_{mi}"): st.session_state[dk] = True; st.rerun()
            else:
                with st.sidebar.form(f"df_{mi}"):
                    st.warning(t("delete_member_confirm", name=m["name"]))
                    dp = st.text_input(t("password"), type="password", key=f"dp_{mi}")
                    cc1, cc2 = st.columns(2)
                    if cc1.form_submit_button(f"\U0001f5d1 {t('delete_member_btn')}"):
                        cu = next((x for x in members if x["email"]==ue), None)
                        if cu and dp == cu.get("password", DEFAULT_PASSWORD):
                            tasks[:] = [tk for tk in tasks if tk["assigned_to"] != m["email"]]
                            members.pop(mi)
                            save_members(spreadsheet, members)
                            save_tasks(spreadsheet, tasks)
                            st.session_state[dk] = False; st.rerun()
                        else: st.error(t("delete_member_wrong_pw"))
                    if cc2.form_submit_button(t("cancel")): st.session_state[dk] = False; st.rerun()
    st.sidebar.divider(); st.sidebar.header(f"\U0001f511 {t('change_password')}")
    with st.sidebar.form("cpw"):
        op = st.text_input(t("old_password"), type="password")
        np = st.text_input(t("new_password"), type="password")
        np2 = st.text_input(t("confirm_password"), type="password")
        if st.form_submit_button(t("change")):
            m = next((x for x in members if x["email"]==ue), None)
            if not m: st.error(t("user_not_found"))
            elif op != m.get("password",DEFAULT_PASSWORD): st.error(t("old_pw_wrong"))
            elif not np: st.warning(t("enter_new_pw"))
            elif np != np2: st.error(t("pw_mismatch"))
            else: m["password"] = np; save_members(spreadsheet, members); st.success(t("pw_changed"))

# --- Attachments (Google Drive) ---------------------------------------

def render_attachments(task, tidx, spreadsheet, tasks, drive_service, prefix="a"):
    att = get_task_attachments(task)
    if att:
        for ai, a in enumerate(att):
            c1, c2 = st.columns([6, 1])
            with c1:
                link = a.get("link", "")
                fname = a.get("name", "file")
                if link:
                    st.markdown(f"\U0001f4ce [{fname}]({link})")
                else:
                    st.caption(f"\U0001f4ce {fname}")
            with c2:
                if st.button("\U0001f5d1", key=f"{prefix}_rm_{tidx}_{ai}", help=t("remove_attachment")):
                    att.pop(ai)
                    set_task_attachments(task, att)
                    save_tasks(spreadsheet, tasks)
                    st.rerun()

    upload_key = f"{prefix}_u_{tidx}"
    up = st.file_uploader(t("upload_files"), accept_multiple_files=True, key=upload_key, label_visibility="collapsed")
    saved_key = f"{prefix}_saved_{tidx}"
    if up:
        already_saved = st.session_state.get(saved_key, set())
        new_files = [f for f in up if f.name not in already_saved]
        if new_files:
            folder_id = get_attachments_folder_id()
            for f in new_files:
                file_info = upload_to_drive(drive_service, f.read(), f.name, f.type or "application/octet-stream", folder_id)
                att.append(file_info)
                already_saved.add(f.name)
            st.session_state[saved_key] = already_saved
            set_task_attachments(task, att)
            save_tasks(spreadsheet, tasks)
            st.rerun()

# --- Tab: Neue Aufgabe ------------------------------------------------

def tab_new(spreadsheet, members, tasks, user):
    asgn = get_assignable(members, user["email"], user["role"])
    tt = st.radio(t("task_type"), [f"\U0001f4e4 {t('team_task')}", f"\U0001f512 {t('private_task')}"], horizontal=True, key="tt")
    isp = "\U0001f512" in tt
    if not isp and not asgn: st.info(t("invite_first")); return
    with st.form("at", clear_on_submit=True):
        st.subheader(t("create_task"))
        ti = st.text_input(t("title")); de = st.text_area(t("description"), height=80)
        c1,c2,c3,c4 = st.columns(4)
        if isp: c1.info(f"\U0001f512 {user['name']}"); ae = user["email"]
        else:
            mn = [f"{m['name']} ({m['email']})" for m in asgn]; a = c1.selectbox(t("assign_to"), mn)
        ca = c2.selectbox(t("category"), get_category_options())
        pr = c3.selectbox(t("priority"), get_priority_options())
        du = c4.date_input(t("due_date"), value=None)
        if st.form_submit_button(f"\u2705 {t('create_btn')}", type="primary"):
            if ti:
                if not isp: ae = asgn[mn.index(a)]["email"]
                ms = max((tk.get("sort_order",0) for tk in tasks), default=0)
                tasks.append({"title":ti,"description":de,"assigned_to":ae,"priority":priority_to_key(pr),"due":du.strftime("%d.%m.%Y") if du else "","created":date.today().strftime("%d.%m.%Y"),"status":"Offen","comment":"","category":category_to_key(ca),"sort_order":ms+1,"created_by":user["email"],"private":"ja" if isp else "","archived":"","attachments":""})
                save_tasks(spreadsheet, tasks); st.success(t("private_created" if isp else "task_created", title=ti)); st.rerun()
            else: st.warning(t("enter_title"))

# --- Shared task list (card layout from local version) ----------------

def _task_row_label(tk):
    """Build a structured label: title + description on left, badges on right."""
    priv = "\U0001f512 " if tk.get("private") == "ja" else ""
    title = tk.get("title", "")
    desc = tk.get("description", "")
    cmt = tk.get("comment", "")
    sk = status_to_display(tk.get("status", "Offen"))
    ck = category_to_display(tk.get("category", "Sonstiges"))
    pk = priority_to_display(tk.get("priority", "Mittel"))
    due = tk.get("due", "")
    att = len(get_task_attachments(tk))

    line1 = f"{priv}**{title}**  \u00b7  {sk}  \u00b7  {ck}  \u00b7  {pk}"
    if due:
        line1 += f"  \u00b7  \U0001f4c5 {due}"
    if att:
        line1 += f"  \u00b7  \U0001f4ce{att}"
    lines = [line1]
    if desc:
        lines.append(desc)
    if cmt:
        lines.append(f"\U0001f4ac {cmt}")
    return "  \n".join(lines)


def render_task_list(task_list, tasks, members, spreadsheet, user, drive_service, prefix="tl"):
    """Renders a list of tasks with compact card rows, edit, sort, archive, attachments."""
    ue, ur = user["email"], user["role"]
    task_list.sort(key=lambda x: x.get("sort_order", 0))
    st.markdown(TABLE_HEADER, unsafe_allow_html=True)
    for j, tk in enumerate(task_list):
        tidx = tasks.index(tk)
        sk = tk.get("status", "Offen")
        # HTML card + action buttons
        card_col, btn_col = st.columns([10, 1])
        with card_col:
            st.markdown(task_row_html(tk), unsafe_allow_html=True)
        with btn_col:
            if st.button("\u270f", key=f"{prefix}_e_{tidx}", help=t("hint_edit")):
                st.session_state[f"eo_{prefix}_{tidx}"] = not st.session_state.get(f"eo_{prefix}_{tidx}", False)
                st.rerun()
            if st.button("\u2191", key=f"{prefix}_u_{tidx}", help=t("hint_move_up")):
                if j > 0:
                    p = task_list[j-1]
                    tk["sort_order"], p["sort_order"] = p["sort_order"], tk["sort_order"]
                    save_tasks(spreadsheet, tasks); st.rerun()
            if st.button("\u2193", key=f"{prefix}_d_{tidx}", help=t("hint_move_down")):
                if j < len(task_list)-1:
                    n = task_list[j+1]
                    tk["sort_order"], n["sort_order"] = n["sort_order"], tk["sort_order"]
                    save_tasks(spreadsheet, tasks); st.rerun()
            if sk == "Erledigt":
                if st.button("\U0001f4e6", key=f"{prefix}_a_{tidx}", help=t("hint_archive")):
                    tk["archived"] = "ja"; save_tasks(spreadsheet, tasks); st.rerun()
        if st.session_state.get(f"eo_{prefix}_{tidx}", False):
            with st.form(f"{prefix}_ef_{tidx}"):
                et = st.text_input(t("title"), value=tk["title"])
                ed = st.text_area(t("description"), value=tk.get("description", ""), height=120)
                e1, e2 = st.columns(2)
                asgn = get_assignable(members, ue, ur)
                an = [f"{m['name']} ({m['email']})" for m in asgn]
                ca = next((f"{m['name']} ({m['email']})" for m in asgn if m["email"] == tk["assigned_to"]), an[0] if an else "")
                ea = e1.selectbox(t("assign_to"), an, index=an.index(ca) if ca in an else 0, key=f"{prefix}_ea_{tidx}")
                ec = e2.selectbox(t("category"), get_category_options(), index=CATEGORY_KEYS.index(tk.get("category", "Sonstiges")) if tk.get("category", "Sonstiges") in CATEGORY_KEYS else 0, key=f"{prefix}_ec_{tidx}")
                e3, e4, e5 = st.columns(3)
                ep = e3.selectbox(t("priority"), get_priority_options(), index=PRIORITY_KEYS.index(tk.get("priority", "Mittel")) if tk.get("priority", "Mittel") in PRIORITY_KEYS else 0, key=f"{prefix}_ep_{tidx}")
                es = e4.selectbox(t("status"), get_status_options(), index=STATUS_KEYS.index(sk) if sk in STATUS_KEYS else 0, key=f"{prefix}_es_{tidx}")
                try:
                    dv = datetime.strptime(tk["due"], "%d.%m.%Y").date() if tk.get("due") else None
                except:
                    dv = None
                edue = e5.date_input(t("due_date"), value=dv, key=f"{prefix}_ed_{tidx}")
                sb1, sb2 = st.columns(2)
                save_clicked = sb1.form_submit_button(f"\U0001f4be {t('save_changes')}")
                cancel_clicked = sb2.form_submit_button(f"\u2715 {t('cancel')}")
                if save_clicked:
                    tk["title"] = et; tk["description"] = ed
                    tk["assigned_to"] = asgn[an.index(ea)]["email"]
                    tk["priority"] = priority_to_key(ep); tk["status"] = status_to_key(es)
                    tk["category"] = category_to_key(ec)
                    tk["due"] = edue.strftime("%d.%m.%Y") if edue else ""
                    st.session_state[f"eo_{prefix}_{tidx}"] = False
                    save_tasks(spreadsheet, tasks); st.rerun()
                if cancel_clicked:
                    st.session_state[f"eo_{prefix}_{tidx}"] = False; st.rerun()
            render_attachments(tk, tidx, spreadsheet, tasks, drive_service, prefix=f"{prefix}_att")

# --- Tab: Meine Aufgaben ---------------------------------------------

def tab_my(spreadsheet, tasks, user, drive_service, members):
    st.markdown(COMPACT_CSS, unsafe_allow_html=True)
    bt = [tk for tk in tasks if tk["assigned_to"] == user["email"] and tk.get("private") != "ja" and tk.get("archived") != "ja"]
    pt = [tk for tk in tasks if tk["assigned_to"] == user["email"] and tk.get("private") == "ja" and tk.get("created_by") == user["email"] and tk.get("archived") != "ja"]
    if not bt and not pt: st.info(t("no_tasks")); return

    s1, s2, s3 = st.columns([2, 1, 1])
    sq = s1.text_input("\U0001f50d", placeholder=t("search_placeholder"), key="ms", label_visibility="collapsed")
    fs = s2.multiselect(t("status_filter"), get_status_options(), default=get_status_options(), key="msf")
    fc = s3.multiselect(t("category_filter"), get_category_options(), default=get_category_options(), key="mcf")
    fsk = [status_to_key(s) for s in fs]; fck = [category_to_key(c) for c in fc]

    def fl(tl):
        return search_tasks([tk for tk in tl if tk.get("status", "Offen") in fsk and tk.get("category", "Sonstiges") in fck], sq)

    fb = fl(bt)
    if fb:
        with st.expander(f"\U0001f4e5 {t('boss_tasks', count=len(fb))}", expanded=True):
            render_task_list(fb, tasks, members, spreadsheet, user, drive_service, prefix="boss")
        my_arch = [tk for tk in tasks if tk["assigned_to"] == user["email"] and tk.get("archived") == "ja" and tk.get("private") != "ja"]
        if my_arch:
            _archive_inline(my_arch, "my_boss")

    fp = fl(pt)
    if fp:
        with st.expander(f"\U0001f512 {t('private_tasks', count=len(fp))}", expanded=True):
            render_task_list(fp, tasks, members, spreadsheet, user, drive_service, prefix="priv")
        priv_arch = [tk for tk in tasks if tk["assigned_to"] == user["email"] and tk.get("archived") == "ja" and tk.get("private") == "ja"]
        if priv_arch:
            _archive_inline(priv_arch, "my_priv")

    if not fb and not fp: st.info(t("no_filtered"))


def _archive_inline(archived, prefix):
    """Inline archive toggle."""
    ak = f"sa_{prefix}"
    if ak not in st.session_state: st.session_state[ak] = False
    lb = t("hide_archive") if st.session_state[ak] else t("show_archive")
    icon = "\U0001f4c2" if st.session_state[ak] else "\U0001f4e6"
    if st.button(f"{icon} {lb} ({len(archived)})", key=f"at_{prefix}"):
        st.session_state[ak] = not st.session_state[ak]; st.rerun()
    if st.session_state[ak]:
        for tk in archived:
            att = get_task_attachments(tk); ats = f" \U0001f4ce{len(att)}" if att else ""
            st.caption(f"~~{tk['title']}~~ | {category_to_display(tk.get('category', 'Sonstiges'))} | \U0001f4c5 {tk.get('due', '-')}{ats}")

# --- Tab: Aufgaben Team -----------------------------------------------

def tab_team(spreadsheet, members, tasks, user, drive_service):
    st.markdown(COMPACT_CSS, unsafe_allow_html=True)
    app_url = get_app_url()
    ue, ur = user["email"], user["role"]
    # Visible tasks excluding own tasks (those are in "Meine Aufgaben")
    vis = [tk for tk in get_visible(tasks, members, ue, ur) if tk["assigned_to"] != ue]
    team = [m for m in get_assignable(members, ue, ur) if m["email"] != ue]
    if not team: st.info(t("no_members")); return
    if not vis and not any(get_archived(tasks, m["email"]) for m in team): st.info(t("no_tasks_yet")); return

    s1, s2, s3 = st.columns([2, 1, 1])
    sq = s1.text_input("\U0001f50d", placeholder=t("search_placeholder"), key="ts", label_visibility="collapsed")
    fs = s2.multiselect(t("status_filter"), get_status_options(), default=get_status_options(), key="tsf")
    fc = s3.multiselect(t("category_filter"), get_category_options(), default=get_category_options(), key="tcf")
    fsk = [status_to_key(s) for s in fs]; fck = [category_to_key(c) for c in fc]

    for member in team:
        mt = search_tasks([tk for tk in vis if tk["assigned_to"] == member["email"] and tk.get("status", "Offen") in fsk and tk.get("category", "Sonstiges") in fck], sq)
        arch = get_archived(tasks, member["email"])
        if not mt and not arch: continue
        if not mt and arch:
            with st.expander(f"**{member['name']}** \u2014 0 {t('tasks_count')}", expanded=False):
                _archive_team(member, tasks, spreadsheet)
            continue

        with st.expander(f"**{member['name']}** \u2014 {len(mt)} {t('tasks_count')}", expanded=True):
            render_task_list(mt, tasks, members, spreadsheet, user, drive_service, prefix=f"tm_{member['email'][:8]}")
            amt = [tk for tk in tasks if tk["assigned_to"]==member["email"] and tk.get("private")!="ja" and tk.get("archived")!="ja"]
            if amt:
                ml = build_mailto(member["email"], member["name"], amt, app_url)
                mail_key = f"mail_{member['email']}"
                if st.button(f"📧 {t('send_email_to', name=member['name'])}", key=mail_key):
                    st.components.v1.html(f"<script>window.parent.location.href='{ml}';</script>", height=0)
            _archive_team(member, tasks, spreadsheet)
    st.divider()
    if st.button(f"📧 {t('send_all_emails')}"):
        js_parts = []
        for m in team:
            mtl = [tk for tk in vis if tk["assigned_to"] == m["email"] and tk.get("private") != "ja"]
            if mtl:
                ml = build_mailto(m["email"], m["name"], mtl, app_url)
                js_parts.append(f"window.open('{ml}');")
        if js_parts:
            st.components.v1.html(f"<script>{''.join(js_parts)}</script>", height=0)

def _archive_team(member, tasks, spreadsheet):
    ar = get_archived(tasks, member["email"])
    if not ar: return
    ak = f"sa_{member['email']}"
    if ak not in st.session_state: st.session_state[ak] = False
    lb = t("hide_archive") if st.session_state[ak] else t("show_archive")
    if st.button(f"{'\U0001f4c2' if st.session_state[ak] else '\U0001f4e6'} {lb} ({len(ar)})", key=f"at_{member['email']}"):
        st.session_state[ak] = not st.session_state[ak]; st.rerun()
    if st.session_state[ak]:
        for tk in ar:
            att = get_task_attachments(tk); ats = f" \U0001f4ce{len(att)}" if att else ""
            st.caption(f"~~{tk['title']}~~ | {category_to_display(tk.get('category','Sonstiges'))} | \U0001f4c5 {tk.get('due','-')}{ats}")

# --- Tab: Import ------------------------------------------------------

def tab_import(spreadsheet, members, tasks, user):
    st.subheader(f"\U0001f4e5 {t('import_title')}"); st.write(t("import_desc"))
    team = get_assignable(members, user["email"], user["role"])
    if not team: st.info(t("no_members")); return
    mn = [f"{m['name']} ({m['email']})" for m in team]
    sm = st.selectbox(t("reply_from"), mn, key="im"); me = team[mn.index(sm)]["email"]
    rt = st.text_area(t("paste_reply"), height=200, key="rt")
    if not rt: return
    pattern = re.compile(r"\[(\d+)\]\s*(.+?)(?:\n|\r\n?)(?:.*?(?:\n|\r\n?))*?\s*Status:\s*(.+?)(?:\n|\r\n?)\s*Kommentar:\s*(.*?)(?:\n|\r\n?|$)", re.DOTALL)
    updates = []
    for m in pattern.finditer(rt):
        sk = None
        for k in STATUS_KEYS:
            if k.lower() == m.group(3).strip().lower(): sk = k; break
        if not sk:
            for i, o in enumerate(get_status_options()):
                if o.lower() == m.group(3).strip().lower(): sk = STATUS_KEYS[i]; break
        updates.append({"index": int(m.group(1)), "status": sk or m.group(3).strip(), "comment": m.group(4).strip()})
    mt = [tk for tk in tasks if tk["assigned_to"]==me and tk.get("archived")!="ja"]
    if not updates: st.warning(t("no_updates")); return
    st.subheader(t("detected_changes")); valid = []
    for u in updates:
        tn = u["index"]
        if tn < 1 or tn > len(mt): st.warning(t("task_not_found", num=tn)); continue
        tk = mt[tn-1]; os_ = tk.get("status","Offen"); oc = tk.get("comment","")
        ns, nc = u["status"], u["comment"]; iv = ns in STATUS_KEYS
        if os_ != ns or oc != nc:
            c1,c2 = st.columns(2); c1.write(f"**[{tn}] {tk['title']}**")
            ch = []
            if os_ != ns: ch.append(f"Status: {status_to_display(os_)} \u2192 **{status_to_display(ns)}**" if iv else f"\u26a0\ufe0f {ns}")
            if oc != nc: ch.append(f"\U0001f4ac _{nc}_")
            c2.write(" | ".join(ch))
            if iv: valid.append({"n": tn, "s": ns, "c": nc})
    if valid:
        if st.button(f"\u2705 {t('apply_changes',count=len(valid))}", type="primary"):
            for u in valid: tk = mt[u["n"]-1]; tk["status"]=u["s"]; tk["comment"]=u["c"]
            save_tasks(spreadsheet, tasks); st.success(t("changes_applied")); st.rerun()

# --- Main -------------------------------------------------------------

def main():
    st.set_page_config(page_title="Team Task Manager", page_icon="\U0001f4cb", layout="wide")
    if "lang" not in st.session_state: st.session_state.lang = "de"

    spreadsheet = get_gsheet_connection()
    ensure_worksheets(spreadsheet)
    members = load_members(spreadsheet)
    tasks = load_tasks(spreadsheet)

    if not members: show_setup(spreadsheet); return
    if "user" not in st.session_state: show_login(members); return

    user = st.session_state.user
    drive_service = get_drive_service()

    render_logo(); render_lang_toggle()
    st.title(f"\U0001f4cb {t('app_title')}")
    render_sidebar(spreadsheet, members, tasks, user)

    t1,t2,t3,t4 = st.tabs([f"\u2795 {t('tab_new_task')}", f"\U0001f4cb {t('tab_my_tasks')}", f"\U0001f4ca {t('tab_team')}", f"\U0001f4e5 {t('tab_import')}"])
    with t1: tab_new(spreadsheet, members, tasks, user)
    with t2: tab_my(spreadsheet, tasks, user, drive_service, members)
    with t3: tab_team(spreadsheet, members, tasks, user, drive_service)
    with t4: tab_import(spreadsheet, members, tasks, user)


if __name__ == "__main__":
    main()
