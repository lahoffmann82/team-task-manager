import streamlit as st
import json
import re
import urllib.parse
import base64
import io
from datetime import date, datetime

from google.oauth2.service_account import Credentials
import gspread
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ─── Übersetzungen ─────────────────────────────────────────────

TRANSLATIONS = {
    "de": {
        "app_title": "Team Task Manager",
        "login_title": "Team Task Manager — Login",
        "setup_title": "Team Task Manager — Ersteinrichtung",
        "logout": "Ausloggen",
        "save": "Speichern",
        "saved": "Gespeichert!",
        "delete": "Löschen",
        "edit": "Bearbeiten",
        "save_changes": "Änderungen speichern",
        "cancel": "Abbrechen",
        "email": "Email",
        "password": "Passwort",
        "login": "Einloggen",
        "enter_email_pw": "Bitte Email und Passwort eingeben.",
        "no_member_found": "Kein Mitarbeiter mit dieser Email gefunden.",
        "wrong_password": "Falsches Passwort.",
        "setup_welcome": "Willkommen! Bitte lege den ersten Administrator an.",
        "your_name": "Dein Name *",
        "your_email": "Deine Email *",
        "create_admin": "Admin-Account erstellen",
        "admin_created": "Admin-Account erstellt!",
        "fill_all_fields": "Bitte alle Felder ausfüllen.",
        "my_team": "Mein Team",
        "invite_member": "Mitarbeiter einladen",
        "name": "Name",
        "invite": "Einladen",
        "email_exists": "Email existiert bereits.",
        "enter_name_email": "Bitte Name und Email eingeben.",
        "invited": "eingeladen! (Passwort: {pw})",
        "current_members": "Aktuelle Mitarbeiter",
        "change_password": "Passwort ändern",
        "old_password": "Altes Passwort",
        "new_password": "Neues Passwort",
        "confirm_password": "Neues Passwort bestätigen",
        "change": "Ändern",
        "user_not_found": "User nicht gefunden.",
        "old_pw_wrong": "Altes Passwort ist falsch.",
        "enter_new_pw": "Bitte neues Passwort eingeben.",
        "pw_mismatch": "Neue Passwörter stimmen nicht überein.",
        "pw_changed": "Passwort geändert!",
        "tab_new_task": "Neue Aufgabe",
        "tab_my_tasks": "Meine Aufgaben",
        "tab_team": "Team & Emails",
        "tab_import": "Import",
        "task_type": "Aufgabentyp",
        "team_task": "Team-Aufgabe (für Mitarbeiter)",
        "private_task": "Private Aufgabe (nur für mich)",
        "create_task": "Neue Aufgabe erstellen",
        "title": "Titel *",
        "description": "Beschreibung",
        "assign_to": "Zuweisen an *",
        "category": "Kategorie",
        "priority": "Priorität",
        "due_date": "Fällig am",
        "create_btn": "Aufgabe erstellen",
        "task_created": "Aufgabe '{title}' erstellt!",
        "private_created": "Private Aufgabe '{title}' erstellt!",
        "enter_title": "Bitte Titel eingeben.",
        "invite_first": "Bitte lade zuerst Mitarbeiter in der Sidebar ein.",
        "assigned_to_private": "Zugewiesen an: **{name}** (privat)",
        "no_tasks": "Du hast aktuell keine Aufgaben.",
        "status_filter": "Status-Filter",
        "category_filter": "Kategorie-Filter",
        "boss_tasks": "Aufgaben von Vorgesetzten ({count})",
        "private_tasks": "Meine privaten Aufgaben ({count})",
        "no_filtered": "Keine Aufgaben für die gewählten Filter.",
        "status": "Status",
        "comment": "Kommentar",
        "no_members": "Noch keine Mitarbeiter vorhanden.",
        "no_tasks_yet": "Noch keine Aufgaben vorhanden.",
        "tasks_count": "Aufgabe(n)",
        "send_email_to": "Aufgaben an {name} senden",
        "send_all_emails": "Alle Emails auf einmal öffnen",
        "import_title": "Email-Antwort importieren",
        "import_desc": "Füge die Antwort-Email eines Mitarbeiters hier ein. Der Status und Kommentar werden automatisch erkannt.",
        "reply_from": "Antwort von Mitarbeiter",
        "paste_reply": "Email-Antwort hier einfügen",
        "no_updates": "Keine Aufgaben-Updates erkannt.",
        "detected_changes": "Erkannte Änderungen",
        "task_not_found": "Aufgabe [{num}] existiert nicht — übersprungen.",
        "no_change": "keine Änderung",
        "unknown_status": "Unbekannter Status: '{status}'",
        "apply_changes": "{count} Änderung(en) übernehmen",
        "changes_applied": "Änderungen übernommen!",
        "no_valid_changes": "Keine gültigen Änderungen erkannt.",
        "email_greeting": "Hallo {name},",
        "email_intro": "hier sind deine aktuellen Aufgaben:",
        "email_tasks_header": "--- AUFGABEN ---",
        "email_tasks_footer": "--- ENDE ---",
        "email_app_link": "Bitte aktualisiere deinen Fortschritt direkt in der App:",
        "email_reply_hint": "Oder antworte auf diese Email mit aktualisierten Status- und Kommentar-Feldern.",
        "email_closing": "Viele Gruesse",
        "email_subject": "Deine Aufgaben - {date}",
        "email_due": "Faellig: {date}",
        "email_desc": "Beschreibung: {desc}",
        "priority_high": "Hoch",
        "priority_medium": "Mittel",
        "priority_low": "Niedrig",
        "status_open": "Offen",
        "status_in_progress": "In Arbeit",
        "status_done": "Erledigt",
        "cat_marketing": "Marketing",
        "cat_it": "IT",
        "cat_purchasing": "Einkauf",
        "cat_sales": "Vertrieb",
        "cat_other": "Sonstiges",
        # New translations
        "archive": "Archivieren",
        "archived": "Archiviert",
        "show_archive": "Archiv anzeigen",
        "hide_archive": "Archiv ausblenden",
        "archive_empty": "Keine archivierten Aufgaben.",
        "task_archived": "Aufgabe archiviert!",
        "delete_member_title": "Mitarbeiter entfernen",
        "delete_member_confirm": "Passwort eingeben um {name} zu entfernen:",
        "delete_member_btn": "Entfernen bestätigen",
        "delete_member_wrong_pw": "Falsches Passwort. Mitarbeiter nicht entfernt.",
        "delete_member_success": "{name} wurde entfernt.",
        "attachments": "Anhänge",
        "upload_files": "Dateien hochladen",
        "no_attachments": "Keine Anhänge",
        "file_uploaded": "Datei '{name}' hochgeladen!",
        "open_file": "Öffnen",
    },
    "en": {
        "app_title": "Team Task Manager",
        "login_title": "Team Task Manager — Login",
        "setup_title": "Team Task Manager — Initial Setup",
        "logout": "Logout",
        "save": "Save",
        "saved": "Saved!",
        "delete": "Delete",
        "edit": "Edit",
        "save_changes": "Save changes",
        "cancel": "Cancel",
        "email": "Email",
        "password": "Password",
        "login": "Login",
        "enter_email_pw": "Please enter email and password.",
        "no_member_found": "No member found with this email.",
        "wrong_password": "Wrong password.",
        "setup_welcome": "Welcome! Please create the first administrator account.",
        "your_name": "Your name *",
        "your_email": "Your email *",
        "create_admin": "Create admin account",
        "admin_created": "Admin account created!",
        "fill_all_fields": "Please fill in all fields.",
        "my_team": "My Team",
        "invite_member": "Invite member",
        "name": "Name",
        "invite": "Invite",
        "email_exists": "Email already exists.",
        "enter_name_email": "Please enter name and email.",
        "invited": "invited! (Password: {pw})",
        "current_members": "Current members",
        "change_password": "Change password",
        "old_password": "Old password",
        "new_password": "New password",
        "confirm_password": "Confirm new password",
        "change": "Change",
        "user_not_found": "User not found.",
        "old_pw_wrong": "Old password is incorrect.",
        "enter_new_pw": "Please enter a new password.",
        "pw_mismatch": "New passwords don't match.",
        "pw_changed": "Password changed!",
        "tab_new_task": "New Task",
        "tab_my_tasks": "My Tasks",
        "tab_team": "Team & Emails",
        "tab_import": "Import",
        "task_type": "Task type",
        "team_task": "Team task (for members)",
        "private_task": "Private task (only for me)",
        "create_task": "Create new task",
        "title": "Title *",
        "description": "Description",
        "assign_to": "Assign to *",
        "category": "Category",
        "priority": "Priority",
        "due_date": "Due date",
        "create_btn": "Create task",
        "task_created": "Task '{title}' created!",
        "private_created": "Private task '{title}' created!",
        "enter_title": "Please enter a title.",
        "invite_first": "Please invite team members in the sidebar first.",
        "assigned_to_private": "Assigned to: **{name}** (private)",
        "no_tasks": "You currently have no tasks.",
        "status_filter": "Status filter",
        "category_filter": "Category filter",
        "boss_tasks": "Tasks from manager ({count})",
        "private_tasks": "My private tasks ({count})",
        "no_filtered": "No tasks match the selected filters.",
        "status": "Status",
        "comment": "Comment",
        "no_members": "No members yet.",
        "no_tasks_yet": "No tasks yet.",
        "tasks_count": "task(s)",
        "send_email_to": "Send tasks to {name}",
        "send_all_emails": "Open all emails at once",
        "import_title": "Import email reply",
        "import_desc": "Paste a team member's email reply here. Status and comments will be detected automatically.",
        "reply_from": "Reply from member",
        "paste_reply": "Paste email reply here",
        "no_updates": "No task updates detected.",
        "detected_changes": "Detected changes",
        "task_not_found": "Task [{num}] does not exist — skipped.",
        "no_change": "no change",
        "unknown_status": "Unknown status: '{status}'",
        "apply_changes": "Apply {count} change(s)",
        "changes_applied": "Changes applied!",
        "no_valid_changes": "No valid changes detected.",
        "email_greeting": "Hello {name},",
        "email_intro": "here are your current tasks:",
        "email_tasks_header": "--- TASKS ---",
        "email_tasks_footer": "--- END ---",
        "email_app_link": "Please update your progress directly in the app:",
        "email_reply_hint": "Or reply to this email with updated status and comment fields.",
        "email_closing": "Best regards",
        "email_subject": "Your tasks - {date}",
        "email_due": "Due: {date}",
        "email_desc": "Description: {desc}",
        "priority_high": "High",
        "priority_medium": "Medium",
        "priority_low": "Low",
        "status_open": "Open",
        "status_in_progress": "In Progress",
        "status_done": "Done",
        "cat_marketing": "Marketing",
        "cat_it": "IT",
        "cat_purchasing": "Purchasing",
        "cat_sales": "Sales",
        "cat_other": "Other",
        "archive": "Archive",
        "archived": "Archived",
        "show_archive": "Show archive",
        "hide_archive": "Hide archive",
        "archive_empty": "No archived tasks.",
        "task_archived": "Task archived!",
        "delete_member_title": "Remove member",
        "delete_member_confirm": "Enter password to remove {name}:",
        "delete_member_btn": "Confirm removal",
        "delete_member_wrong_pw": "Wrong password. Member not removed.",
        "delete_member_success": "{name} has been removed.",
        "attachments": "Attachments",
        "upload_files": "Upload files",
        "no_attachments": "No attachments",
        "file_uploaded": "File '{name}' uploaded!",
        "open_file": "Open",
    },
}


def get_lang():
    return st.session_state.get("lang", "de")


def t(key, **kwargs):
    lang = get_lang()
    text = TRANSLATIONS.get(lang, TRANSLATIONS["de"]).get(key, key)
    if kwargs:
        text = text.format(**kwargs)
    return text


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
CATEGORY_COLORS_MAP = {
    "Marketing": "#E91E63", "IT": "#2196F3", "Einkauf": "#FF9800",
    "Vertrieb": "#4CAF50", "Sonstiges": "#9E9E9E",
}


def priority_to_key(display):
    opts = get_priority_options()
    return PRIORITY_KEYS[opts.index(display)] if display in opts else PRIORITY_KEYS[0]

def priority_to_display(key):
    return get_priority_options()[PRIORITY_KEYS.index(key)] if key in PRIORITY_KEYS else get_priority_options()[0]

def status_to_key(display):
    opts = get_status_options()
    return STATUS_KEYS[opts.index(display)] if display in opts else STATUS_KEYS[0]

def status_to_display(key):
    return get_status_options()[STATUS_KEYS.index(key)] if key in STATUS_KEYS else get_status_options()[0]

def category_to_key(display):
    opts = get_category_options()
    return CATEGORY_KEYS[opts.index(display)] if display in opts else CATEGORY_KEYS[0]

def category_to_display(key):
    return get_category_options()[CATEGORY_KEYS.index(key)] if key in CATEGORY_KEYS else get_category_options()[0]


DEFAULT_PASSWORD = "12345"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ─── Logo ──────────────────────────────────────────────────────

LOGO_SVG = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 80" width="280">
  <style>
    .title { font-family: 'Arial Black', Arial, sans-serif; font-weight: 900; font-size: 32px; fill: #000; letter-spacing: 2px; }
    .subtitle { font-family: Arial, sans-serif; font-weight: 700; font-size: 14px; fill: #fff; letter-spacing: 3px; }
  </style>
  <text x="10" y="35" class="title">ALIMENTAST</text>
  <text x="313" y="35" class="title">I</text>
  <text x="323" y="35" class="title">C</text>
  <circle cx="318" cy="12" r="4" fill="#E31E24"/>
  <rect x="10" y="48" width="340" height="26" fill="#000" rx="2"/>
  <text x="35" y="67" class="subtitle">HOUSE OF BRANDS</text>
</svg>'''


def render_logo():
    st.markdown(f'<div style="margin-bottom:10px;">{LOGO_SVG}</div>', unsafe_allow_html=True)


def render_lang_toggle():
    lang = get_lang()
    c1, c2 = st.columns([8, 1])
    with c2:
        new_lang = st.selectbox("🌐", ["Deutsch", "English"], index=0 if lang == "de" else 1, key="lang_select", label_visibility="collapsed")
        new_code = "de" if new_lang == "Deutsch" else "en"
        if new_code != lang:
            st.session_state.lang = new_code
            st.rerun()
    return c1


# ─── Google Sheets + Drive ─────────────────────────────────────


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


def get_or_create_attachments_folder(drive_service):
    """Get or create 'TaskManager_Attachments' folder in Drive."""
    query = "name='TaskManager_Attachments' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = drive_service.files().list(q=query, fields="files(id)").execute()
    files = results.get("files", [])
    if files:
        return files[0]["id"]
    folder_meta = {
        "name": "TaskManager_Attachments",
        "mimeType": "application/vnd.google-apps.folder",
    }
    folder = drive_service.files().create(body=folder_meta, fields="id").execute()
    # Make it accessible via link
    drive_service.permissions().create(
        fileId=folder["id"],
        body={"type": "anyone", "role": "reader"},
    ).execute()
    return folder["id"]


def upload_to_drive(drive_service, file_bytes, filename, mime_type, folder_id):
    """Upload file to Google Drive and return file ID + web link."""
    file_meta = {"name": filename, "parents": [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
    uploaded = drive_service.files().create(body=file_meta, media_body=media, fields="id,webViewLink").execute()
    # Make readable by anyone with link
    drive_service.permissions().create(
        fileId=uploaded["id"],
        body={"type": "anyone", "role": "reader"},
    ).execute()
    return {"id": uploaded["id"], "name": filename, "link": uploaded.get("webViewLink", "")}


def get_app_url():
    return st.secrets.get("app", {}).get("url", "")


MEMBER_HEADERS = ["name", "email", "password", "role", "invited_by"]
TASK_HEADERS = [
    "title", "description", "assigned_to", "priority",
    "due", "created", "status", "comment", "category", "sort_order",
    "created_by", "private", "archived", "attachments",
]


def _migrate_worksheet(spreadsheet, ws_name, expected_headers):
    """Add missing columns to an existing worksheet without losing data."""
    ws = spreadsheet.worksheet(ws_name)
    current_headers = ws.row_values(1)
    missing = [h for h in expected_headers if h not in current_headers]
    if not missing:
        return
    # Append missing headers to the right
    start_col = len(current_headers) + 1
    for i, header in enumerate(missing):
        col = start_col + i
        ws.update_cell(1, col, header)


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
        rows = [
            [task.get("title", ""), task.get("description", ""), task.get("assigned_to", ""),
             task.get("priority", ""), task.get("due", ""), task.get("created", ""),
             task.get("status", "Offen"), task.get("comment", ""), task.get("category", "Sonstiges"),
             task.get("sort_order", 0), task.get("created_by", ""), task.get("private", ""),
             task.get("archived", ""), task.get("attachments", "")]
            for task in tasks
        ]
        ws.update(f"A2:N{len(rows)+1}", rows)


def get_task_attachments(task):
    """Parse attachments JSON from task."""
    raw = task.get("attachments", "")
    if not raw:
        return []
    try:
        return json.loads(raw)
    except (json.JSONDecodeError, TypeError):
        return []


def set_task_attachments(task, attachments):
    """Store attachments as JSON string."""
    task["attachments"] = json.dumps(attachments, ensure_ascii=False) if attachments else ""


# ─── Hilfsfunktionen ──────────────────────────────────────────


def get_my_team(members, user_email):
    return [m for m in members if m.get("invited_by") == user_email]


def get_assignable_members(members, user_email, user_role):
    if user_role == "admin":
        return [m for m in members]
    return get_my_team(members, user_email)


def get_visible_tasks(tasks, members, user_email, user_role):
    if user_role == "admin":
        return [task for task in tasks if (task.get("private") != "ja" or task.get("created_by") == user_email) and task.get("archived") != "ja"]
    team_emails = [m["email"] for m in get_my_team(members, user_email)]
    return [task for task in tasks if task["assigned_to"] in team_emails and task.get("private") != "ja" and task.get("archived") != "ja"]


def get_archived_tasks(tasks, member_email):
    return [task for task in tasks if task["assigned_to"] == member_email and task.get("archived") == "ja"]


# ─── Email ─────────────────────────────────────────────────────


def build_email_body(name, tasks_list, app_url=""):
    lines = [t("email_greeting", name=name), "", t("email_intro"), "", t("email_tasks_header")]
    for i, task in enumerate(tasks_list, 1):
        priority = f" [{priority_to_display(task['priority'])}]" if task.get("priority") else ""
        due = f" ({t('email_due', date=task['due'])})" if task.get("due") else ""
        cat = f" #{category_to_display(task.get('category', 'Sonstiges'))}" if task.get("category") else ""
        lines.append(f"[{i}] {task['title']}{priority}{due}{cat}")
        if task.get("description"):
            lines.append(f"    {t('email_desc', desc=task['description'])}")
        lines.append(f"    {t('status')}: {status_to_display(task.get('status', 'Offen'))}")
        lines.append(f"    {t('comment')}: {task.get('comment', '')}")
        lines.append("")
    lines.append(t("email_tasks_footer"))
    lines.append("")
    if app_url:
        lines += [t("email_app_link"), app_url, ""]
    lines += [t("email_reply_hint"), "", t("email_closing")]
    return "\n".join(lines)


def build_mailto_link(email, name, tasks_list, app_url=""):
    subject = t("email_subject", date=date.today().strftime("%d.%m.%Y"))
    body = build_email_body(name, tasks_list, app_url)
    params = urllib.parse.urlencode({"subject": subject, "body": body}, quote_via=urllib.parse.quote)
    return f"mailto:{email}?{params}"


def parse_reply(text):
    updates = []
    pattern = re.compile(
        r"\[(\d+)\]\s*(.+?)(?:\n|\r\n?)(?:.*?(?:\n|\r\n?))*?\s*Status:\s*(.+?)(?:\n|\r\n?)\s*Kommentar:\s*(.*?)(?:\n|\r\n?|$)",
        re.DOTALL,
    )
    for match in pattern.finditer(text):
        idx, title, status_text, comment_text = int(match.group(1)), match.group(2).strip(), match.group(3).strip(), match.group(4).strip()
        status_key = None
        for key in STATUS_KEYS:
            if key.lower() == status_text.lower():
                status_key = key
                break
        if not status_key:
            for i, opt in enumerate(get_status_options()):
                if opt.lower() == status_text.lower():
                    status_key = STATUS_KEYS[i]
                    break
        updates.append({"index": idx, "title": title, "status": status_key or status_text, "comment": comment_text})
    return updates


# ─── Login & Setup ─────────────────────────────────────────────


def show_setup(spreadsheet):
    render_logo()
    render_lang_toggle()
    st.title(f"📋 {t('setup_title')}")
    st.info(t("setup_welcome"))
    with st.form("setup_form"):
        name = st.text_input(t("your_name"))
        email = st.text_input(t("your_email"))
        password = st.text_input(t("password") + " *", type="password", value=DEFAULT_PASSWORD)
        if st.form_submit_button(t("create_admin")):
            if not name or not email or not password:
                st.warning(t("fill_all_fields"))
            else:
                save_members(spreadsheet, [{"name": name, "email": email, "password": password, "role": "admin", "invited_by": ""}])
                st.session_state.user = {"email": email, "name": name, "role": "admin"}
                st.success(t("admin_created"))
                st.rerun()


def show_login(members):
    render_logo()
    render_lang_toggle()
    st.title(f"📋 {t('login_title')}")
    with st.form("login_form"):
        email = st.text_input(t("email"))
        password = st.text_input(t("password"), type="password")
        if st.form_submit_button(t("login")):
            if not email or not password:
                st.warning(t("enter_email_pw"))
            else:
                member = next((m for m in members if m["email"].lower() == email.lower()), None)
                if not member:
                    st.error(t("no_member_found"))
                elif member.get("password", DEFAULT_PASSWORD) != password:
                    st.error(t("wrong_password"))
                else:
                    st.session_state.user = {"email": member["email"], "name": member["name"], "role": member.get("role", "member")}
                    st.rerun()


# ─── Sidebar ──────────────────────────────────────────────────


def render_sidebar(spreadsheet, members, tasks, user):
    ue, ur = user["email"], user["role"]

    st.sidebar.write(f"👤 **{user['name']}**")
    st.sidebar.caption(f"{ue} ({ur.capitalize()})")
    if st.sidebar.button(f"🚪 {t('logout')}"):
        del st.session_state.user
        st.rerun()

    st.sidebar.divider()
    st.sidebar.header(f"👥 {t('my_team')}")

    with st.sidebar.form("add_member", clear_on_submit=True):
        st.subheader(t("invite_member"))
        new_name = st.text_input(t("name"))
        new_email = st.text_input(t("email"))
        if st.form_submit_button(t("invite")):
            if new_name and new_email:
                if any(m["email"].lower() == new_email.lower() for m in members):
                    st.error(t("email_exists"))
                else:
                    members.append({"name": new_name, "email": new_email, "password": DEFAULT_PASSWORD, "role": "member", "invited_by": ue})
                    save_members(spreadsheet, members)
                    st.success(f"{new_name} {t('invited', pw=DEFAULT_PASSWORD)}")
                    st.rerun()
            else:
                st.warning(t("enter_name_email"))

    my_team = get_my_team(members, ue) if ur != "admin" else members
    if my_team:
        st.sidebar.subheader(t("current_members"))
        for i, m in enumerate(my_team):
            if m["email"] == ue:
                continue
            midx = members.index(m)
            badge = "👑" if m.get("role") == "admin" else ""
            st.sidebar.write(f"{badge} **{m['name']}** ({m['email']})")

            # Delete with password confirmation
            del_key = f"del_confirm_{midx}"
            if del_key not in st.session_state:
                st.session_state[del_key] = False

            if not st.session_state[del_key]:
                if st.sidebar.button(f"✕ {t('delete')}", key=f"del_btn_{midx}"):
                    st.session_state[del_key] = True
                    st.rerun()
            else:
                with st.sidebar.form(f"del_form_{midx}"):
                    st.warning(t("delete_member_confirm", name=m["name"]))
                    del_pw = st.text_input(t("password"), type="password", key=f"del_pw_{midx}")
                    c1, c2 = st.columns(2)
                    confirm = c1.form_submit_button(f"🗑 {t('delete_member_btn')}")
                    cancel = c2.form_submit_button(t("cancel"))
                    if confirm:
                        current_user = next((x for x in members if x["email"] == ue), None)
                        if current_user and del_pw == current_user.get("password", DEFAULT_PASSWORD):
                            tasks[:] = [task for task in tasks if task["assigned_to"] != m["email"]]
                            members.pop(midx)
                            save_members(spreadsheet, members)
                            save_tasks(spreadsheet, tasks)
                            st.session_state[del_key] = False
                            st.success(t("delete_member_success", name=m["name"]))
                            st.rerun()
                        else:
                            st.error(t("delete_member_wrong_pw"))
                    if cancel:
                        st.session_state[del_key] = False
                        st.rerun()

    st.sidebar.divider()
    st.sidebar.header(f"🔑 {t('change_password')}")
    with st.sidebar.form("change_pw"):
        old_pw = st.text_input(t("old_password"), type="password")
        new_pw = st.text_input(t("new_password"), type="password")
        new_pw2 = st.text_input(t("confirm_password"), type="password")
        if st.form_submit_button(t("change")):
            member = next((m for m in members if m["email"] == ue), None)
            if not member:
                st.error(t("user_not_found"))
            elif old_pw != member.get("password", DEFAULT_PASSWORD):
                st.error(t("old_pw_wrong"))
            elif not new_pw:
                st.warning(t("enter_new_pw"))
            elif new_pw != new_pw2:
                st.error(t("pw_mismatch"))
            else:
                member["password"] = new_pw
                save_members(spreadsheet, members)
                st.success(t("pw_changed"))


# ─── Attachments UI ───────────────────────────────────────────


def render_attachments(task, tidx, spreadsheet, tasks, drive_service, prefix="att"):
    """Render attachment section for a task."""
    attachments = get_task_attachments(task)

    # Show existing attachments
    if attachments:
        st.caption(f"📎 {t('attachments')} ({len(attachments)})")
        for ai, att in enumerate(attachments):
            link = att.get("link", "")
            name = att.get("name", "file")
            st.markdown(f"[📄 {name}]({link})")
    else:
        st.caption(f"📎 {t('no_attachments')}")

    # Upload new
    uploaded = st.file_uploader(
        t("upload_files"), accept_multiple_files=True,
        key=f"{prefix}_upload_{tidx}",
        label_visibility="collapsed",
    )
    if uploaded:
        folder_id = get_or_create_attachments_folder(drive_service)
        for f in uploaded:
            file_info = upload_to_drive(drive_service, f.read(), f.name, f.type or "application/octet-stream", folder_id)
            attachments.append(file_info)
            st.success(t("file_uploaded", name=f.name))
        set_task_attachments(task, attachments)
        save_tasks(spreadsheet, tasks)
        st.rerun()


# ─── Tab: Neue Aufgabe ────────────────────────────────────────


def tab_new_task(spreadsheet, members, tasks, user):
    assignable = get_assignable_members(members, user["email"], user["role"])
    task_type = st.radio(t("task_type"), [f"📤 {t('team_task')}", f"🔒 {t('private_task')}"], horizontal=True, key="task_type")
    is_private = "🔒" in task_type

    if not is_private and not assignable:
        st.info(t("invite_first"))
        return

    with st.form("add_task", clear_on_submit=True):
        st.subheader(t("create_task"))
        title = st.text_input(t("title"))
        description = st.text_area(t("description"))
        c1, c2 = st.columns(2)
        if is_private:
            c1.info(t("assigned_to_private", name=user["name"]))
            assigned_email = user["email"]
        else:
            mn = [f"{m['name']} ({m['email']})" for m in assignable]
            assigned = c1.selectbox(t("assign_to"), mn)
        cat_display = c2.selectbox(t("category"), get_category_options())
        c3, c4 = st.columns(2)
        prio_display = c3.selectbox(t("priority"), get_priority_options())
        due = c4.date_input(t("due_date"), value=None)

        if st.form_submit_button(f"✅ {t('create_btn')}", type="primary"):
            if title:
                if not is_private:
                    assigned_email = assignable[mn.index(assigned)]["email"]
                max_sort = max((task.get("sort_order", 0) for task in tasks), default=0)
                new_task = {
                    "title": title, "description": description, "assigned_to": assigned_email,
                    "priority": priority_to_key(prio_display),
                    "due": due.strftime("%d.%m.%Y") if due else "",
                    "created": date.today().strftime("%d.%m.%Y"),
                    "status": "Offen", "comment": "", "category": category_to_key(cat_display),
                    "sort_order": max_sort + 1, "created_by": user["email"],
                    "private": "ja" if is_private else "", "archived": "", "attachments": "",
                }
                tasks.append(new_task)
                save_tasks(spreadsheet, tasks)
                st.success(t("private_created" if is_private else "task_created", title=title))
                st.rerun()
            else:
                st.warning(t("enter_title"))


# ─── Tab: Meine Aufgaben ──────────────────────────────────────


def render_task_card(task, idx, spreadsheet, tasks, user, drive_service, prefix="my", allow_delete=False):
    sk = task.get("status", "Offen")
    ck = task.get("category", "Sonstiges")
    is_priv = task.get("private") == "ja"

    with st.container(border=True):
        c1, c2, c3 = st.columns([5, 2, 2])
        pfx = "🔒 " if is_priv else ""
        c1.write(f"### {pfx}{task['title']}")
        c2.markdown(f":{STATUS_COLORS_MAP.get(sk, 'gray')}[{status_to_display(sk)}]")
        if task.get("due"):
            c3.write(f"📅 {task['due']}")

        tc1, tc2 = st.columns([1, 5])
        tc1.markdown(f'<span style="background:{CATEGORY_COLORS_MAP.get(ck, "#9E9E9E")};color:white;padding:2px 8px;border-radius:12px;font-size:0.8em;">{category_to_display(ck)}</span>', unsafe_allow_html=True)
        if task.get("priority"):
            pc = {"Hoch": "red", "Mittel": "orange", "Niedrig": "green"}.get(task["priority"], "gray")
            tc2.markdown(f"{t('priority')}: :{pc}[{priority_to_display(task['priority'])}]")

        if task.get("description"):
            st.write(f"_{task['description']}_")

        # Attachments
        render_attachments(task, idx, spreadsheet, tasks, drive_service, prefix=prefix)

        # Status + Comment
        st.divider()
        cs, cc = st.columns([1, 2])
        cur_idx = STATUS_KEYS.index(sk) if sk in STATUS_KEYS else 0
        new_s = cs.selectbox(t("status"), get_status_options(), index=cur_idx, key=f"{prefix}_s_{idx}")
        new_c = cc.text_input(t("comment"), value=task.get("comment", ""), key=f"{prefix}_c_{idx}")
        new_sk = status_to_key(new_s)

        if new_sk != sk or new_c != task.get("comment", ""):
            if st.button(f"💾 {t('save')}", key=f"{prefix}_sv_{idx}"):
                task["status"] = new_sk
                task["comment"] = new_c
                save_tasks(spreadsheet, tasks)
                st.success(t("saved"))
                st.rerun()


def tab_my_tasks(spreadsheet, tasks, user, drive_service):
    boss_tasks = [task for task in tasks if task["assigned_to"] == user["email"] and task.get("private") != "ja" and task.get("archived") != "ja"]
    private_tasks = [task for task in tasks if task["assigned_to"] == user["email"] and task.get("private") == "ja" and task.get("created_by") == user["email"] and task.get("archived") != "ja"]

    if not boss_tasks and not private_tasks:
        st.info(t("no_tasks"))
        return

    cf1, cf2 = st.columns(2)
    fs = cf1.multiselect(t("status_filter"), get_status_options(), default=get_status_options(), key="my_sf")
    fc = cf2.multiselect(t("category_filter"), get_category_options(), default=get_category_options(), key="my_cf")
    fsk = [status_to_key(s) for s in fs]
    fck = [category_to_key(c) for c in fc]

    def filt(tl):
        return [task for task in tl if task.get("status", "Offen") in fsk and task.get("category", "Sonstiges") in fck]

    fb = filt(boss_tasks)
    if fb:
        st.subheader(f"📥 {t('boss_tasks', count=len(fb))}")
        for i, task in enumerate(fb):
            render_task_card(task, i, spreadsheet, tasks, user, drive_service, prefix="boss", allow_delete=False)

    fp = filt(private_tasks)
    if fp:
        st.subheader(f"🔒 {t('private_tasks', count=len(fp))}")
        for i, task in enumerate(fp):
            render_task_card(task, i, spreadsheet, tasks, user, drive_service, prefix="priv", allow_delete=False)

    if not fb and not fp:
        st.info(t("no_filtered"))


# ─── Tab: Team & Emails ───────────────────────────────────────


def tab_team_overview(spreadsheet, members, tasks, user, drive_service):
    app_url = get_app_url()
    ue, ur = user["email"], user["role"]
    visible = get_visible_tasks(tasks, members, ue, ur)
    team = get_assignable_members(members, ue, ur)

    if not team:
        st.info(t("no_members"))
        return
    if not visible:
        st.info(t("no_tasks_yet"))
        return

    cf1, cf2 = st.columns(2)
    fs = cf1.multiselect(t("status_filter"), get_status_options(), default=get_status_options(), key="tsf")
    fc = cf2.multiselect(t("category_filter"), get_category_options(), default=get_category_options(), key="tcf")
    fsk = [status_to_key(s) for s in fs]
    fck = [category_to_key(c) for c in fc]

    for member in team:
        mt = [task for task in visible if task["assigned_to"] == member["email"] and task.get("status", "Offen") in fsk and task.get("category", "Sonstiges") in fck]
        if not mt:
            # Still show member for archive button
            archived = get_archived_tasks(tasks, member["email"])
            if archived:
                with st.expander(f"**{member['name']}** — 0 {t('tasks_count')}", expanded=False):
                    _render_archive_section(member, tasks, spreadsheet)
            continue

        mt.sort(key=lambda x: x.get("sort_order", 0))

        with st.expander(f"**{member['name']}** — {len(mt)} {t('tasks_count')}", expanded=True):
            for j, task in enumerate(mt):
                tidx = tasks.index(task)
                c1, c2, c3, c4, c5 = st.columns([4, 1, 1, 1, 1])

                c1.write(f"**{task['title']}**")
                sk = task.get("status", "Offen")
                c2.markdown(f":{STATUS_COLORS_MAP.get(sk, 'gray')}[{status_to_display(sk)}]")
                ck = task.get("category", "Sonstiges")
                c3.markdown(f'<span style="background:{CATEGORY_COLORS_MAP.get(ck, "#9E9E9E")};color:white;padding:2px 6px;border-radius:10px;font-size:0.75em;">{category_to_display(ck)}</span>', unsafe_allow_html=True)

                # Sort buttons
                if c4.button("⬆", key=f"up_{tidx}"):
                    if j > 0:
                        prev = mt[j-1]
                        task["sort_order"], prev["sort_order"] = prev["sort_order"], task["sort_order"]
                        save_tasks(spreadsheet, tasks)
                        st.rerun()
                if c5.button("⬇", key=f"down_{tidx}"):
                    if j < len(mt)-1:
                        nxt = mt[j+1]
                        task["sort_order"], nxt["sort_order"] = nxt["sort_order"], task["sort_order"]
                        save_tasks(spreadsheet, tasks)
                        st.rerun()

                # Details
                details = []
                if task.get("priority"):
                    pc = {"Hoch": "red", "Mittel": "orange", "Niedrig": "green"}.get(task["priority"], "gray")
                    details.append(f"{t('priority')}: :{pc}[{priority_to_display(task['priority'])}]")
                if task.get("due"):
                    details.append(f"📅 {task['due']}")
                if task.get("description"):
                    details.append(f"_{task['description']}_")
                if task.get("comment"):
                    details.append(f"💬 {task['comment']}")
                # Show attachment count
                att_count = len(get_task_attachments(task))
                if att_count > 0:
                    details.append(f"📎 {att_count}")
                if details:
                    st.caption(" | ".join(details))

                # Archive button for completed tasks
                if sk == "Erledigt":
                    if st.button(f"📦 {t('archive')}", key=f"arch_{tidx}"):
                        task["archived"] = "ja"
                        save_tasks(spreadsheet, tasks)
                        st.success(t("task_archived"))
                        st.rerun()

                # Edit expander
                with st.expander(f"✏️ {t('edit')}", expanded=False):
                    with st.form(f"edit_{tidx}"):
                        et = st.text_input(t("title"), value=task["title"])
                        ed = st.text_area(t("description"), value=task.get("description", ""))
                        e1, e2 = st.columns(2)
                        assignable = get_assignable_members(members, ue, ur)
                        an = [f"{m['name']} ({m['email']})" for m in assignable]
                        ca = next((f"{m['name']} ({m['email']})" for m in assignable if m["email"] == task["assigned_to"]), an[0] if an else "")
                        ai = an.index(ca) if ca in an else 0
                        ea = e1.selectbox(t("assign_to"), an, index=ai, key=f"ea_{tidx}")
                        ecat = e2.selectbox(t("category"), get_category_options(), index=CATEGORY_KEYS.index(task.get("category", "Sonstiges")) if task.get("category", "Sonstiges") in CATEGORY_KEYS else 0, key=f"ec_{tidx}")
                        e3, e4 = st.columns(2)
                        ep = e3.selectbox(t("priority"), get_priority_options(), index=PRIORITY_KEYS.index(task.get("priority", "Mittel")) if task.get("priority", "Mittel") in PRIORITY_KEYS else 0, key=f"ep_{tidx}")
                        es = e4.selectbox(t("status"), get_status_options(), index=STATUS_KEYS.index(sk) if sk in STATUS_KEYS else 0, key=f"es_{tidx}")
                        try:
                            dv = datetime.strptime(task["due"], "%d.%m.%Y").date() if task.get("due") else None
                        except ValueError:
                            dv = None
                        edue = st.date_input(t("due_date"), value=dv, key=f"ed_{tidx}")
                        if st.form_submit_button(f"💾 {t('save_changes')}"):
                            task["title"] = et
                            task["description"] = ed
                            task["assigned_to"] = assignable[an.index(ea)]["email"]
                            task["priority"] = priority_to_key(ep)
                            task["status"] = status_to_key(es)
                            task["category"] = category_to_key(ecat)
                            task["due"] = edue.strftime("%d.%m.%Y") if edue else ""
                            save_tasks(spreadsheet, tasks)
                            st.success(t("saved"))
                            st.rerun()

                    # Attachments in edit section
                    render_attachments(task, tidx, spreadsheet, tasks, drive_service, prefix=f"team_att")

            # Email button
            amt = [task for task in tasks if task["assigned_to"] == member["email"] and task.get("private") != "ja" and task.get("archived") != "ja"]
            if amt:
                ml = build_mailto_link(member["email"], member["name"], amt, app_url)
                st.markdown(f'<a href="{ml}" target="_blank" style="display:inline-block;padding:8px 16px;background:#4CAF50;color:white;border-radius:4px;text-decoration:none;margin-top:8px;">📧 {t("send_email_to", name=member["name"])}</a>', unsafe_allow_html=True)

            # Archive section
            _render_archive_section(member, tasks, spreadsheet)

    st.divider()
    if st.button(f"📧 {t('send_all_emails')}"):
        links = []
        for member in team:
            mt = [task for task in visible if task["assigned_to"] == member["email"] and task.get("private") != "ja"]
            if mt:
                links.append(build_mailto_link(member["email"], member["name"], mt, app_url))
        js = "".join(f'window.open("{link}");' for link in links)
        st.components.v1.html(f"<script>{js}</script>", height=0)


def _render_archive_section(member, tasks, spreadsheet):
    """Render archive toggle for a member."""
    archived = get_archived_tasks(tasks, member["email"])
    if not archived:
        return

    arch_key = f"show_arch_{member['email']}"
    if arch_key not in st.session_state:
        st.session_state[arch_key] = False

    if st.button(
        f"{'📂' if st.session_state[arch_key] else '📦'} {t('hide_archive') if st.session_state[arch_key] else t('show_archive')} ({len(archived)})",
        key=f"arch_toggle_{member['email']}",
    ):
        st.session_state[arch_key] = not st.session_state[arch_key]
        st.rerun()

    if st.session_state[arch_key]:
        for task in archived:
            sk = task.get("status", "Erledigt")
            st.markdown(
                f"~~{task['title']}~~ | {status_to_display(sk)} | "
                f"{category_to_display(task.get('category', 'Sonstiges'))} | "
                f"📅 {task.get('due', '-')}"
            )
            att = get_task_attachments(task)
            if att:
                for a in att:
                    st.markdown(f"  📎 [{a['name']}]({a['link']})")


# ─── Tab: Import ──────────────────────────────────────────────


def tab_import_replies(spreadsheet, members, tasks, user):
    st.subheader(f"📥 {t('import_title')}")
    st.write(t("import_desc"))
    team = get_assignable_members(members, user["email"], user["role"])
    if not team:
        st.info(t("no_members"))
        return

    mn = [f"{m['name']} ({m['email']})" for m in team]
    sm = st.selectbox(t("reply_from"), mn, key="import_member")
    me = team[mn.index(sm)]["email"]
    rt = st.text_area(t("paste_reply"), height=300, key="reply_text")

    if rt:
        updates = parse_reply(rt)
        mtasks = [task for task in tasks if task["assigned_to"] == me and task.get("archived") != "ja"]
        if not updates:
            st.warning(t("no_updates"))
        else:
            st.subheader(t("detected_changes"))
            valid = []
            for u in updates:
                tn = u["index"]
                if tn < 1 or tn > len(mtasks):
                    st.warning(t("task_not_found", num=tn))
                    continue
                task = mtasks[tn-1]
                os, oc = task.get("status", "Offen"), task.get("comment", "")
                ns, nc = u["status"], u["comment"]
                iv = ns in STATUS_KEYS
                if os != ns or oc != nc:
                    c1, c2 = st.columns(2)
                    c1.write(f"**[{tn}] {task['title']}**")
                    ch = []
                    if os != ns:
                        ch.append(f"Status: {status_to_display(os)} → **{status_to_display(ns)}**" if iv else f"⚠️ {t('unknown_status', status=ns)}")
                    if oc != nc:
                        ch.append(f"{t('comment')}: _{nc}_")
                    c2.write(" | ".join(ch))
                    if iv:
                        valid.append({"task_num": tn, "status": ns, "comment": nc})
                else:
                    st.write(f"[{tn}] {task['title']} — {t('no_change')}")

            if valid:
                if st.button(f"✅ {t('apply_changes', count=len(valid))}", type="primary"):
                    for u in valid:
                        task = mtasks[u["task_num"]-1]
                        task["status"] = u["status"]
                        task["comment"] = u["comment"]
                    save_tasks(spreadsheet, tasks)
                    st.success(t("changes_applied"))
                    st.rerun()
            elif updates:
                st.info(t("no_valid_changes"))


# ─── Main ───────────────────────────────────────────────────────


def main():
    st.set_page_config(page_title="Team Task Manager", page_icon="📋", layout="wide")
    if "lang" not in st.session_state:
        st.session_state.lang = "de"

    spreadsheet = get_gsheet_connection()
    ensure_worksheets(spreadsheet)
    members = load_members(spreadsheet)
    tasks = load_tasks(spreadsheet)

    if not members:
        show_setup(spreadsheet)
        return
    if "user" not in st.session_state:
        show_login(members)
        return

    user = st.session_state.user
    drive_service = get_drive_service()

    render_logo()
    render_lang_toggle()
    st.title(f"📋 {t('app_title')}")
    render_sidebar(spreadsheet, members, tasks, user)

    tab1, tab2, tab3, tab4 = st.tabs([
        f"➕ {t('tab_new_task')}", f"📋 {t('tab_my_tasks')}",
        f"📊 {t('tab_team')}", f"📥 {t('tab_import')}",
    ])
    with tab1:
        tab_new_task(spreadsheet, members, tasks, user)
    with tab2:
        tab_my_tasks(spreadsheet, tasks, user, drive_service)
    with tab3:
        tab_team_overview(spreadsheet, members, tasks, user, drive_service)
    with tab4:
        tab_import_replies(spreadsheet, members, tasks, user)


if __name__ == "__main__":
    main()
