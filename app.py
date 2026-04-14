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
        "cat_product_dev": "Produktentwicklung", "cat_project_mgmt": "Projektmanagement", "cat_operations": "Operations",
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
        "last_edited": "Zuletzt bearbeitet von {name} am {date}",
        "tab_projects": "Projekte",
        "projects": "Projekte",
        "new_project": "Neues Projekt erstellen",
        "project_title": "Projekttitel *",
        "project_desc": "Projektbeschreibung",
        "project_responsible": "Zuständige Mitarbeiter",
        "project_created": "Projekt '{title}' erstellt!",
        "project_saved": "Projekt gespeichert!",
        "project_no_projects": "Noch keine Projekte vorhanden.",
        "project_enter_title": "Bitte Projekttitel eingeben.",
        "project_all_visible": "Alle Mitarbeiter können Projekte einsehen. Nur Ersteller und Zuständige können bearbeiten.",
        "project_send": "Projekt an Zuständige senden",
        "project_readonly": "Nur Ersteller und Zuständige können dieses Projekt bearbeiten.",
        "project_created_by": "Erstellt von",
        "project_status_active": "Aktiv",
        "project_status_done": "Abgeschlossen",
        "project_status_paused": "Pausiert",
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
        "cat_product_dev": "Product Development", "cat_project_mgmt": "Project Management", "cat_operations": "Operations",
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
        "last_edited": "Last edited by {name} on {date}",
        "tab_projects": "Projects",
        "projects": "Projects",
        "new_project": "Create new project",
        "project_title": "Project title *",
        "project_desc": "Project description",
        "project_responsible": "Responsible members",
        "project_created": "Project '{title}' created!",
        "project_saved": "Project saved!",
        "project_no_projects": "No projects yet.",
        "project_enter_title": "Please enter a project title.",
        "project_all_visible": "All members can view projects. Only creator and responsible members can edit.",
        "project_send": "Send project to responsible members",
        "project_readonly": "Only creator and responsible members can edit this project.",
        "project_created_by": "Created by",
        "project_status_active": "Active",
        "project_status_done": "Completed",
        "project_status_paused": "Paused",
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
    return [t("cat_marketing"), t("cat_it"), t("cat_purchasing"), t("cat_sales"), t("cat_other"), t("cat_product_dev"), t("cat_project_mgmt"), t("cat_operations")]

PRIORITY_KEYS = ["Hoch", "Mittel", "Niedrig"]
STATUS_KEYS = ["Offen", "In Arbeit", "Erledigt"]
CATEGORY_KEYS = ["Marketing", "IT", "Einkauf", "Vertrieb", "Sonstiges", "Produktentwicklung", "Projektmanagement", "Operations"]
STATUS_COLORS_MAP = {"Offen": "gray", "In Arbeit": "blue", "Erledigt": "green"}
CATEGORY_COLORS_MAP = {"Marketing": "#E91E63", "IT": "#2196F3", "Einkauf": "#FF9800", "Vertrieb": "#4CAF50", "Sonstiges": "#9E9E9E", "Produktentwicklung": "#9C27B0", "Projektmanagement": "#00BCD4", "Operations": "#FF5722"}

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
    "last_edited_by", "last_edited_at",
]

PROJECT_HEADERS = [
    "title", "description", "created_by", "created_at",
    "responsible", "status", "due", "attachments",
    "last_edited_by", "last_edited_at",
]

LOGIN_LOG_HEADERS = ["timestamp", "name", "email", "login_type"]

def _migrate_worksheet(ws, expected_headers):
    """Migrate an existing worksheet object – no extra API call needed."""
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
    # Fetch worksheet list only once and reuse the objects – avoids redundant
    # fetch_sheet_metadata() calls that cause the APIError on rerun/edit.
    if st.session_state.get("_worksheets_ensured"):
        return
    existing_ws = {ws.title: ws for ws in spreadsheet.worksheets()}
    if "members" not in existing_ws:
        ws = spreadsheet.add_worksheet(title="members", rows=100, cols=len(MEMBER_HEADERS))
        ws.update("A1:E1", [MEMBER_HEADERS])
    else:
        _migrate_worksheet(existing_ws["members"], MEMBER_HEADERS)
    if "tasks" not in existing_ws:
        ws = spreadsheet.add_worksheet(title="tasks", rows=500, cols=len(TASK_HEADERS))
        ws.update(f"A1:{chr(64+len(TASK_HEADERS))}1", [TASK_HEADERS])
    else:
        _migrate_worksheet(existing_ws["tasks"], TASK_HEADERS)
    if "projects" not in existing_ws:
        ws = spreadsheet.add_worksheet(title="projects", rows=200, cols=len(PROJECT_HEADERS))
        ws.update(f"A1:{chr(64+len(PROJECT_HEADERS))}1", [PROJECT_HEADERS])
    else:
        _migrate_worksheet(existing_ws["projects"], PROJECT_HEADERS)
    if "login_log" not in existing_ws:
        ws = spreadsheet.add_worksheet(title="login_log", rows=2000, cols=len(LOGIN_LOG_HEADERS))
        ws.update("A1:D1", [LOGIN_LOG_HEADERS])
    else:
        _migrate_worksheet(existing_ws["login_log"], LOGIN_LOG_HEADERS)
    st.session_state["_worksheets_ensured"] = True

# --- Login logging ----------------------------------------------------

def log_login(spreadsheet, name, email, login_type="manual"):
    """Append one login entry to the login_log sheet."""
    try:
        ws = spreadsheet.worksheet("login_log")
        ts = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
        ws.append_row([ts, name, email, login_type], value_input_option="RAW")
    except Exception:
        pass  # Never block login because of logging failure

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
        task.setdefault("last_edited_by", "")
        task.setdefault("last_edited_at", "")
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
    n = len(TASK_HEADERS)
    ws.update(f"A1:{chr(64+n)}1", [TASK_HEADERS])
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
                task.get("last_edited_by", ""), task.get("last_edited_at", ""),
            ])
        ws.update(f"A2:{chr(64+n)}{len(rows)+1}", rows)

def load_projects(spreadsheet):
    ws = spreadsheet.worksheet("projects")
    records = ws.get_all_records()
    for p in records:
        p.setdefault("title", "")
        p.setdefault("description", "")
        p.setdefault("created_by", "")
        p.setdefault("created_at", "")
        p.setdefault("responsible", "")
        p.setdefault("status", "Offen")
        p.setdefault("due", "")
        p.setdefault("attachments", "")
        p.setdefault("last_edited_by", "")
        p.setdefault("last_edited_at", "")
        for k in p:
            p[k] = str(p[k]) if p[k] is not None else ""
    return records

def save_projects(spreadsheet, projects):
    ws = spreadsheet.worksheet("projects")
    ws.clear()
    n = len(PROJECT_HEADERS)
    ws.update(f"A1:{chr(64+n)}1", [PROJECT_HEADERS])
    if projects:
        rows = []
        for p in projects:
            att = p.get("attachments", "")
            if isinstance(att, list):
                att = json.dumps(att, ensure_ascii=False) if att else ""
            rows.append([
                p.get("title", ""), p.get("description", ""), p.get("created_by", ""),
                p.get("created_at", ""), p.get("responsible", ""), p.get("status", "Offen"),
                p.get("due", ""), att,
                p.get("last_edited_by", ""), p.get("last_edited_at", ""),
            ])
        ws.update(f"A2:{chr(64+n)}{len(rows)+1}", rows)

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
    first = _first_name(name)
    lines = [f"Hi {first},", "", "Please find your current tasks below:", "", "--- TASKS ---"]
    for i, tk in enumerate(tl, 1):
        prio = priority_to_display(tk.get("priority", "Mittel"))
        due = f" (Due: {tk['due']})" if tk.get("due") else ""
        lines.append(f"[{i}] {tk['title']} [{prio}]{due}")
        if tk.get("description"): lines.append(f"    Description: {tk['description']}")
        lines.append(f"    Status: {status_to_display(tk.get('status','Offen'))}")
        if tk.get("comment"): lines.append(f"    Comment: {tk['comment']}")
        lines.append("")
    lines.append("--- END ---")
    lines.append("")
    if app_url:
        lines += ["You can view and edit all tasks here:", app_url, ""]
    lines += ["Please update your progress directly via the link above or reply to this email.", ""]
    lines += ["If you don't have access yet, use your email address as username and the password: 12345", "You can change your password after your first login.", ""]
    lines += ["Best regards"]
    return "\n".join(lines)

def _first_name(full_name):
    """Extract first name from full name."""
    return full_name.strip().split()[0] if full_name.strip() else full_name

def build_email_body_short(name, tl, app_url=""):
    """Shorter version for mailto links (URL length limit ~2000 chars)."""
    first = _first_name(name)
    lines = [f"Hi {first},", "", "Please find your current tasks below:", ""]
    for i, tk in enumerate(tl, 1):
        lines.append(f"{i}. {tk['title']} - {status_to_display(tk.get('status','Offen'))}")
    lines.append("")
    if app_url:
        lines += ["You can view and edit all tasks here:", app_url, ""]
    lines += ["If you don't have access yet, use your email as username and password: 12345", "You can change your password after your first login.", ""]
    lines.append("Best regards")
    return "\n".join(lines)

def build_mailto(email, name, tl, app_url=""):
    subj = f"Your Tasks - {date.today().strftime('%d.%m.%Y')}"
    body = build_email_body(name, tl, app_url)
    full_url = f"mailto:{email}?" + urllib.parse.urlencode({"subject": subj, "body": body}, quote_via=urllib.parse.quote)
    # If URL too long, use short version
    if len(full_url) > 1800:
        body = build_email_body_short(name, tl, app_url)
        full_url = f"mailto:{email}?" + urllib.parse.urlencode({"subject": subj, "body": body}, quote_via=urllib.parse.quote)
    return full_url

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
    edited_by = _esc(tk.get("last_edited_by", ""))
    edited_at = _esc(tk.get("last_edited_at", ""))
    edit_html = f'<div class="task-info-row"><span class="task-info-label">Bearbeitet</span><span class="task-info-value" style="font-weight:400;font-size:0.9em;">{edited_by}, {edited_at}</span></div>' if edited_by else ""

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
        f'{due_html}{att_html}{edit_html}'
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
            st.session_state["_pending_login_log"] = {"name": m["name"], "email": m["email"], "type": "manual"}
            st.query_params["uid"] = m["email"]
            st.rerun()

# --- Sidebar ----------------------------------------------------------

def render_sidebar(spreadsheet, members, tasks, user):
    ue, ur = user["email"], user["role"]
    st.sidebar.write(f"\U0001f464 **{user['name']}**"); st.sidebar.caption(f"{ue} ({ur.capitalize()})")
    if st.sidebar.button(f"\U0001f6aa {t('logout')}"):
        del st.session_state.user
        st.query_params.clear()
        st.rerun()
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

    # --- Admin: Login History ---
    if ur == "admin":
        st.sidebar.divider()
        st.sidebar.header("📊 Login-Historie")
        if st.sidebar.button("🔄 Historie laden", key="load_login_log"):
            st.session_state["_show_login_log"] = True
        if st.session_state.get("_show_login_log", False):
            try:
                ws = spreadsheet.worksheet("login_log")
                rows = ws.get_all_records()
                if not rows:
                    st.sidebar.info("Noch keine Login-Einträge.")
                else:
                    # Summary stats
                    from collections import Counter
                    all_emails = [r["email"] for r in rows]
                    counts = Counter(all_emails)
                    st.sidebar.markdown("**Logins gesamt:** " + str(len(rows)))
                    st.sidebar.markdown("**Letzte 7 Tage:**")
                    today = datetime.now().date()
                    recent = []
                    for r in rows:
                        try:
                            d = datetime.strptime(r["timestamp"], "%d.%m.%Y %H:%M:%S").date()
                            if (today - d).days <= 7:
                                recent.append(r)
                        except: pass
                    user_counts_7d = Counter(r["email"] for r in recent)
                    member_names_map = {m["email"]: m["name"] for m in members}
                    for email, cnt in sorted(user_counts_7d.items(), key=lambda x: -x[1]):
                        name = member_names_map.get(email, email)
                        st.sidebar.markdown(f"- **{name}**: {cnt}x")
                    if not user_counts_7d:
                        st.sidebar.caption("Keine Logins in den letzten 7 Tagen.")
                    # Last 30 entries (newest first)
                    st.sidebar.markdown("**Letzte Logins:**")
                    for r in reversed(rows[-30:]):
                        name = member_names_map.get(r.get("email",""), r.get("name","?"))
                        ltype = "🔗" if r.get("login_type") == "auto" else "🔑"
                        st.sidebar.caption(f"{ltype} {r.get('timestamp','')} — {name}")
            except Exception as ex:
                st.sidebar.error(f"Fehler: {ex}")

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
                if asgn:
                    an = [f"{m['name']} ({m['email']})" for m in asgn]
                    ca = next((f"{m['name']} ({m['email']})" for m in asgn if m["email"] == tk["assigned_to"]), an[0] if an else "")
                    ea = e1.selectbox(t("assign_to"), an, index=an.index(ca) if ca in an else 0, key=f"{prefix}_ea_{tidx}")
                else:
                    # Member without own team — show current assignment as read-only
                    assigned_member = next((m for m in members if m["email"] == tk["assigned_to"]), None)
                    assigned_name = f"{assigned_member['name']} ({assigned_member['email']})" if assigned_member else tk["assigned_to"]
                    e1.text_input(t("assign_to"), value=assigned_name, disabled=True, key=f"{prefix}_ea_{tidx}")
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
                save_clicked = sb1.form_submit_button(f"💾 {t('save_changes')}")
                cancel_clicked = sb2.form_submit_button(f"✕ {t('cancel')}")
                if save_clicked:
                    tk["title"] = et; tk["description"] = ed
                    if asgn:
                        tk["assigned_to"] = asgn[an.index(ea)]["email"]
                    tk["priority"] = priority_to_key(ep); tk["status"] = status_to_key(es)
                    tk["category"] = category_to_key(ec)
                    tk["due"] = edue.strftime("%d.%m.%Y") if edue else ""
                    tk["last_edited_by"] = user["name"]
                    tk["last_edited_at"] = datetime.now().strftime("%d.%m.%Y %H:%M")
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

    today = date.today()

    for member in team:
        mt = search_tasks([tk for tk in vis if tk["assigned_to"] == member["email"] and tk.get("status", "Offen") in fsk and tk.get("category", "Sonstiges") in fck], sq)
        arch = get_archived(tasks, member["email"])
        if not mt and not arch: continue

        # Count high priority and overdue tasks
        high_count = sum(1 for tk in mt if tk.get("priority", "") == "Hoch" and tk.get("status", "Offen") != "Erledigt")
        overdue_count = 0
        for tk in mt:
            if tk.get("status", "Offen") == "Erledigt": continue
            due_str = tk.get("due", "")
            if due_str:
                try:
                    due_date = datetime.strptime(due_str, "%d.%m.%Y").date()
                    if due_date < today:
                        overdue_count += 1
                except: pass

        # Build expander label with badges
        label = f"**{member['name']}** \u2014 {len(mt)} {t('tasks_count')}"
        if high_count:
            label += f"  🔴 {high_count} Prio Hoch"
        if overdue_count:
            label += f"  ⚠️ {overdue_count} überfällig"

        if not mt and arch:
            with st.expander(f"**{member['name']}** \u2014 0 {t('tasks_count')}", expanded=False):
                _archive_team(member, tasks, spreadsheet)
            continue

        with st.expander(label, expanded=False):
            # Show overdue warning banner inside expander
            if overdue_count:
                st.markdown(f'<div style="background:#fff0f0;border-left:4px solid #e53935;padding:6px 12px;border-radius:4px;margin-bottom:8px;font-size:0.88em;color:#c62828;">⚠️ <strong>{overdue_count} Aufgabe(n) mit abgelaufener Deadline</strong></div>', unsafe_allow_html=True)
            render_task_list(mt, tasks, members, spreadsheet, user, drive_service, prefix=f"tm_{member['email'][:8]}")
            amt = [tk for tk in tasks if tk["assigned_to"]==member["email"] and tk.get("private")!="ja" and tk.get("archived")!="ja"]
            if amt:
                ml = build_mailto(member["email"], member["name"], amt, app_url)
                email_body_text = build_email_body(member["name"], amt, app_url)
                email_subj_text = f"Your Tasks - {date.today().strftime('%d.%m.%Y')}"
                mail_show_key = f"mail_show_{member['email']}"
                if st.button(f"📧 {t('send_email_to', name=member['name'])}", key=f"mail_{member['email']}"):
                    st.session_state[mail_show_key] = not st.session_state.get(mail_show_key, False)
                    st.rerun()
                if st.session_state.get(mail_show_key, False):
                    st.info(f"**An:** {member['email']}\n\n**Betreff:** {email_subj_text}")
                    st.code(email_body_text, language=None)
                    st.caption("Bitte kopiere den Text und füge ihn in eine neue Email in Outlook ein.")
            _archive_team(member, tasks, spreadsheet)

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

# --- Tab: Projects ----------------------------------------------------

PROJECT_STATUS_KEYS = ["Aktiv", "Pausiert", "Abgeschlossen"]

def get_project_status_options():
    return [t("project_status_active"), t("project_status_paused"), t("project_status_done")]

def project_status_to_key(display):
    opts = get_project_status_options()
    return PROJECT_STATUS_KEYS[opts.index(display)] if display in opts else PROJECT_STATUS_KEYS[0]

def project_status_to_display(key):
    return get_project_status_options()[PROJECT_STATUS_KEYS.index(key)] if key in PROJECT_STATUS_KEYS else get_project_status_options()[0]

PROJECT_STATUS_COLORS = {"Aktiv": "#4CAF50", "Pausiert": "#FF9800", "Abgeschlossen": "#9e9e9e"}


def tab_projects(spreadsheet, members, projects, user, drive_service):
    st.markdown(COMPACT_CSS, unsafe_allow_html=True)
    ue = user["email"]
    app_url = get_app_url()

    st.caption(t("project_all_visible"))

    # Create new project
    with st.expander(f"➕ {t('new_project')}", expanded=False):
        with st.form("new_project", clear_on_submit=True):
            pt = st.text_input(t("project_title"))
            pd = st.text_area(t("project_desc"), height=100)
            member_names = [m["name"] for m in members]
            resp = st.multiselect(t("project_responsible"), member_names)
            c1, c2 = st.columns(2)
            ps = c1.selectbox(t("status"), get_project_status_options(), key="new_proj_status")
            pdue = c2.date_input(t("due_date"), value=None, key="new_proj_due")
            if st.form_submit_button(f"✅ {t('new_project')}", type="primary"):
                if pt:
                    resp_emails = [m["email"] for m in members if m["name"] in resp]
                    projects.append({
                        "title": pt, "description": pd, "created_by": ue,
                        "created_at": datetime.now().strftime("%d.%m.%Y %H:%M"),
                        "responsible": ",".join(resp_emails),
                        "status": project_status_to_key(ps),
                        "due": pdue.strftime("%d.%m.%Y") if pdue else "",
                        "attachments": "",
                        "last_edited_by": user["name"],
                        "last_edited_at": datetime.now().strftime("%d.%m.%Y %H:%M"),
                    })
                    save_projects(spreadsheet, projects)
                    st.success(t("project_created", title=pt)); st.rerun()
                else:
                    st.warning(t("project_enter_title"))

    if not projects:
        st.info(t("project_no_projects"))
        return

    # Search
    sq = st.text_input("🔍", placeholder=t("search_placeholder"), key="proj_search", label_visibility="collapsed")

    for i, proj in enumerate(projects):
        # Search filter
        if sq:
            q = sq.lower()
            if q not in proj.get("title", "").lower() and q not in proj.get("description", "").lower():
                continue

        sk = proj.get("status", "Aktiv")
        sc = PROJECT_STATUS_COLORS.get(sk, "#9e9e9e")
        resp_emails = [e.strip() for e in proj.get("responsible", "").split(",") if e.strip()]
        resp_names = [m["name"] for m in members if m["email"] in resp_emails]
        can_edit = (ue == proj.get("created_by", "") or ue in resp_emails or user.get("role") == "admin")
        creator = next((m["name"] for m in members if m["email"] == proj.get("created_by", "")), proj.get("created_by", ""))

        # --- Collapsed summary card ---
        desc_full = proj.get("description", "")
        # First ~120 chars as preview (roughly 1-2 lines)
        desc_preview = (desc_full[:120] + "…") if len(desc_full) > 120 else desc_full
        due = proj.get("due", "")
        edited_by = _esc(proj.get("last_edited_by", ""))
        edited_at = _esc(proj.get("last_edited_at", ""))

        summary_html = (
            '<div class="task-card" style="padding:10px 16px;">'
            '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;">'
            '<div style="flex:1;min-width:0;">'
            f'<div class="task-heading" style="margin-bottom:3px;">{_esc(proj["title"])}</div>'
        )
        if desc_preview:
            summary_html += f'<div style="font-size:0.88em;color:#888;margin-bottom:4px;white-space:pre-wrap;">{_esc(desc_preview)}</div>'
        summary_html += (
            f'<div class="task-comment-text" style="font-size:0.82em;">'
            f'{t("project_created_by")}: <strong>{_esc(creator)}</strong> &nbsp;|&nbsp; '
            f'{t("project_responsible")}: <strong>{", ".join(resp_names) if resp_names else "-"}</strong>'
            f'</div>'
            '</div>'
            f'<div style="text-align:right;white-space:nowrap;">'
            f'<span class="badge" style="background:{sc};display:inline-block;">{project_status_to_display(sk)}</span>'
        )
        if due:
            summary_html += f'<div style="font-size:0.8em;color:#888;margin-top:4px;">📅 {_esc(due)}</div>'
        if edited_by:
            summary_html += f'<div style="font-size:0.75em;color:#aaa;margin-top:2px;">{edited_by}, {edited_at}</div>'
        summary_html += '</div></div></div>'

        proj_col, btn_col = st.columns([10, 1])
        with proj_col:
            st.markdown(summary_html, unsafe_allow_html=True)
        with btn_col:
            if can_edit:
                if st.button("✏", key=f"proj_e_{i}", help=t("hint_edit")):
                    st.session_state[f"proj_edit_{i}"] = not st.session_state.get(f"proj_edit_{i}", False)
                    st.rerun()

        # --- Expandable detail view (collapsed by default) ---
        if len(desc_full) > 120:
            with st.expander("📄 Vollständige Beschreibung anzeigen", expanded=False):
                st.markdown(desc_full)

        # Edit form
        if can_edit and st.session_state.get(f"proj_edit_{i}", False):
            with st.form(f"proj_ef_{i}"):
                ept = st.text_input(t("project_title"), value=proj["title"])
                epd = st.text_area(t("project_desc"), value=proj.get("description", ""), height=120)
                member_names_all = [m["name"] for m in members]
                current_resp = [m["name"] for m in members if m["email"] in resp_emails]
                eresp = st.multiselect(t("project_responsible"), member_names_all, default=current_resp, key=f"proj_resp_{i}")
                ec1, ec2 = st.columns(2)
                eps = ec1.selectbox(t("status"), get_project_status_options(), index=PROJECT_STATUS_KEYS.index(sk) if sk in PROJECT_STATUS_KEYS else 0, key=f"proj_st_{i}")
                try:
                    pdv = datetime.strptime(proj["due"], "%d.%m.%Y").date() if proj.get("due") else None
                except:
                    pdv = None
                epdue = ec2.date_input(t("due_date"), value=pdv, key=f"proj_due_{i}")
                sb1, sb2 = st.columns(2)
                if sb1.form_submit_button(f"💾 {t('save_changes')}"):
                    proj["title"] = ept; proj["description"] = epd
                    proj["responsible"] = ",".join([m["email"] for m in members if m["name"] in eresp])
                    proj["status"] = project_status_to_key(eps)
                    proj["due"] = epdue.strftime("%d.%m.%Y") if epdue else ""
                    proj["last_edited_by"] = user["name"]
                    proj["last_edited_at"] = datetime.now().strftime("%d.%m.%Y %H:%M")
                    st.session_state[f"proj_edit_{i}"] = False
                    save_projects(spreadsheet, projects); st.success(t("project_saved")); st.rerun()
                if sb2.form_submit_button(f"✕ {t('cancel')}"):
                    st.session_state[f"proj_edit_{i}"] = False; st.rerun()
            render_attachments(proj, i, spreadsheet, projects, drive_service, prefix=f"proj_att")
        elif not can_edit:
            pass  # Read-only, no edit button shown

        # Send email to responsible
        if can_edit and resp_emails:
            resp_names_str = ", ".join(resp_names)
            body_lines = [
                f"Hi,", "",
                f"Please find the details for the project \"{proj['title']}\" below:", "",
                f"Description: {proj.get('description', '-')}",
                f"Status: {project_status_to_display(sk)}",
                f"Due: {proj.get('due', '-')}",
                f"Responsible: {resp_names_str}", "",
            ]
            if app_url:
                body_lines += ["You can view and edit all projects here:", app_url, ""]
            body_lines += ["If you don't have access yet, use your email as username and password: 12345", "You can change your password after your first login.", ""]
            body_lines.append("Best regards")
            body = "\n".join(body_lines)
            subj = f"Project: {proj['title']}"
            if st.button(f"📧 {t('project_send')}", key=f"proj_mail_{i}"):
                st.session_state[f"proj_mail_show_{i}"] = True
            if st.session_state.get(f"proj_mail_show_{i}", False):
                st.info(f"**An:** {', '.join(resp_emails)}\n\n**Betreff:** {subj}")
                st.code(body, language=None)
                st.caption("Bitte kopiere den Text oben und füge ihn in eine neue Email in Outlook ein.")

        st.markdown("---")


# --- Main -------------------------------------------------------------

def main():
    st.set_page_config(page_title="Team Task Manager", page_icon="\U0001f4cb", layout="wide")
    if "lang" not in st.session_state: st.session_state.lang = "de"

    spreadsheet = get_gsheet_connection()
    ensure_worksheets(spreadsheet)
    members = load_members(spreadsheet)
    tasks = load_tasks(spreadsheet)
    projects = load_projects(spreadsheet)

    if not members: show_setup(spreadsheet); return

    # Auto-login from URL query parameter (survives app restarts)
    if "user" not in st.session_state:
        qp = st.query_params
        if "uid" in qp:
            uid = qp["uid"]
            member = next((m for m in members if m["email"].lower() == uid.lower()), None)
            if member:
                st.session_state.user = {"email": member["email"], "name": member["name"], "role": member.get("role", "member")}
                st.session_state["_pending_login_log"] = {"name": member["name"], "email": member["email"], "type": "auto"}
            else:
                st.query_params.clear()
        if "user" not in st.session_state:
            show_login(members); return

    user = st.session_state.user

    # Flush pending login log (written after rerun so spreadsheet is ready)
    if "_pending_login_log" in st.session_state:
        ll = st.session_state.pop("_pending_login_log")
        log_login(spreadsheet, ll["name"], ll["email"], ll["type"])
    drive_service = get_drive_service()

    render_logo(); render_lang_toggle()
    st.title(f"\U0001f4cb {t('app_title')}")
    render_sidebar(spreadsheet, members, tasks, user)

    t1,t2,t3,t4,t5 = st.tabs([
        f"➕ {t('tab_new_task')}",
        f"📋 {t('tab_my_tasks')}",
        f"📊 {t('tab_team')}",
        f"📁 {t('tab_projects')}",
        f"📥 {t('tab_import')}",
    ])
    with t1: tab_new(spreadsheet, members, tasks, user)
    with t2: tab_my(spreadsheet, tasks, user, drive_service, members)
    with t3: tab_team(spreadsheet, members, tasks, user, drive_service)
    with t4: tab_projects(spreadsheet, members, projects, user, drive_service)
    with t5: tab_import(spreadsheet, members, tasks, user)


if __name__ == "__main__":
    main()
