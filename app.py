import streamlit as st
import json
import re
import urllib.parse
from datetime import date, datetime

from google.oauth2.service_account import Credentials
import gspread

# ─── Konstanten ────────────────────────────────────────────────

PRIORITY_OPTIONS = ["Hoch", "Mittel", "Niedrig"]
STATUS_OPTIONS = ["Offen", "In Arbeit", "Erledigt"]
STATUS_COLORS = {"Offen": "gray", "In Arbeit": "blue", "Erledigt": "green"}
CATEGORY_OPTIONS = ["Marketing", "IT", "Einkauf", "Vertrieb", "Sonstiges"]
CATEGORY_COLORS = {
    "Marketing": "#E91E63",
    "IT": "#2196F3",
    "Einkauf": "#FF9800",
    "Vertrieb": "#4CAF50",
    "Sonstiges": "#9E9E9E",
}
DEFAULT_PASSWORD = "12345"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ─── Google Sheets ─────────────────────────────────────────────


@st.cache_resource
def get_gsheet_connection():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(st.secrets["spreadsheet"]["key"])
    return spreadsheet


def get_app_url():
    return st.secrets.get("app", {}).get("url", "")


MEMBER_HEADERS = ["name", "email", "password", "role", "invited_by"]
TASK_HEADERS = [
    "title", "description", "assigned_to", "priority",
    "due", "created", "status", "comment", "category", "sort_order",
]


def ensure_worksheets(spreadsheet):
    existing = [ws.title for ws in spreadsheet.worksheets()]
    if "members" not in existing:
        ws = spreadsheet.add_worksheet(title="members", rows=100, cols=len(MEMBER_HEADERS))
        ws.update("A1:E1", [MEMBER_HEADERS])
    if "tasks" not in existing:
        ws = spreadsheet.add_worksheet(title="tasks", rows=500, cols=len(TASK_HEADERS))
        ws.update("A1:J1", [TASK_HEADERS])


def load_members(spreadsheet):
    ws = spreadsheet.worksheet("members")
    records = ws.get_all_records()
    for m in records:
        m.setdefault("password", DEFAULT_PASSWORD)
        m.setdefault("role", "member")
        m.setdefault("invited_by", "")
        # Ensure all values are strings
        for k in m:
            m[k] = str(m[k]) if m[k] is not None else ""
    return records


def load_tasks(spreadsheet):
    ws = spreadsheet.worksheet("tasks")
    records = ws.get_all_records()
    for t in records:
        t.setdefault("status", "Offen")
        t.setdefault("comment", "")
        t.setdefault("category", "Sonstiges")
        t.setdefault("sort_order", 0)
        # Ensure all values are strings except sort_order
        for k in t:
            if k == "sort_order":
                try:
                    t[k] = int(t[k]) if t[k] else 0
                except (ValueError, TypeError):
                    t[k] = 0
            else:
                t[k] = str(t[k]) if t[k] is not None else ""
    return records


def save_members(spreadsheet, members):
    ws = spreadsheet.worksheet("members")
    ws.clear()
    ws.update("A1:E1", [MEMBER_HEADERS])
    if members:
        rows = [
            [m.get("name", ""), m.get("email", ""), m.get("password", DEFAULT_PASSWORD),
             m.get("role", "member"), m.get("invited_by", "")]
            for m in members
        ]
        ws.update(f"A2:E{len(rows)+1}", rows)


def save_tasks(spreadsheet, tasks):
    ws = spreadsheet.worksheet("tasks")
    ws.clear()
    ws.update("A1:J1", [TASK_HEADERS])
    if tasks:
        rows = [
            [
                t.get("title", ""), t.get("description", ""),
                t.get("assigned_to", ""), t.get("priority", ""),
                t.get("due", ""), t.get("created", ""),
                t.get("status", "Offen"), t.get("comment", ""),
                t.get("category", "Sonstiges"), t.get("sort_order", 0),
            ]
            for t in tasks
        ]
        ws.update(f"A2:J{len(rows)+1}", rows)


# ─── Hilfsfunktionen ──────────────────────────────────────────


def get_my_team(members, user_email):
    """Mitarbeiter die vom User eingeladen wurden."""
    return [m for m in members if m.get("invited_by") == user_email]


def get_assignable_members(members, user_email, user_role):
    """Mitarbeiter an die der User Aufgaben zuweisen kann."""
    if user_role == "admin":
        return [m for m in members]
    else:
        return get_my_team(members, user_email)


def get_visible_tasks(tasks, members, user_email, user_role):
    """Aufgaben die der User in der Team-Übersicht sehen kann."""
    if user_role == "admin":
        return tasks
    team_emails = [m["email"] for m in get_my_team(members, user_email)]
    return [t for t in tasks if t["assigned_to"] in team_emails]


# ─── Email-Funktionen ─────────────────────────────────────────


def build_email_body(name, tasks, app_url=""):
    lines = [
        f"Hallo {name},",
        "",
        "hier sind deine aktuellen Aufgaben:",
        "",
        "--- AUFGABEN ---",
    ]
    for i, t in enumerate(tasks, 1):
        priority = f" [{t['priority']}]" if t.get("priority") else ""
        due = f" (Faellig: {t['due']})" if t.get("due") else ""
        cat = f" #{t['category']}" if t.get("category") else ""
        lines.append(f"[{i}] {t['title']}{priority}{due}{cat}")
        if t.get("description"):
            lines.append(f"    Beschreibung: {t['description']}")
        lines.append(f"    Status: {t.get('status', 'Offen')}")
        lines.append(f"    Kommentar: {t.get('comment', '')}")
        lines.append("")
    lines.append("--- ENDE ---")
    lines.append("")
    if app_url:
        lines.append(f"Bitte aktualisiere deinen Fortschritt direkt in der App:")
        lines.append(app_url)
        lines.append("")
    lines.append(
        "Oder antworte auf diese Email mit aktualisierten Status- und Kommentar-Feldern."
    )
    lines.append("")
    lines.append("Viele Gruesse")
    return "\n".join(lines)


def build_mailto_link(email, name, tasks, app_url=""):
    subject = f"Deine Aufgaben - {date.today().strftime('%d.%m.%Y')}"
    body = build_email_body(name, tasks, app_url)
    params = urllib.parse.urlencode(
        {"subject": subject, "body": body}, quote_via=urllib.parse.quote
    )
    return f"mailto:{email}?{params}"


def parse_reply(text):
    updates = []
    pattern = re.compile(
        r"\[(\d+)\]\s*(.+?)(?:\n|\r\n?)"
        r"(?:.*?(?:\n|\r\n?))*?"
        r"\s*Status:\s*(.+?)(?:\n|\r\n?)"
        r"\s*Kommentar:\s*(.*?)(?:\n|\r\n?|$)",
        re.DOTALL,
    )
    for match in pattern.finditer(text):
        idx = int(match.group(1))
        title = match.group(2).strip()
        status = match.group(3).strip()
        comment = match.group(4).strip()
        status_normalized = None
        for opt in STATUS_OPTIONS:
            if opt.lower() == status.lower():
                status_normalized = opt
                break
        if status_normalized is None:
            status_normalized = status
        updates.append(
            {"index": idx, "title": title, "status": status_normalized, "comment": comment}
        )
    return updates


# ─── Login & Setup ─────────────────────────────────────────────


def show_setup(spreadsheet):
    """Ersteinrichtung: Ersten Admin anlegen."""
    st.title("📋 Team Task Manager — Ersteinrichtung")
    st.info("Willkommen! Bitte lege den ersten Administrator an.")

    with st.form("setup_form"):
        name = st.text_input("Dein Name *")
        email = st.text_input("Deine Email *")
        password = st.text_input("Passwort *", type="password", value=DEFAULT_PASSWORD)
        submitted = st.form_submit_button("Admin-Account erstellen")

    if submitted:
        if not name or not email or not password:
            st.warning("Bitte alle Felder ausfüllen.")
            return
        members = [
            {"name": name, "email": email, "password": password, "role": "admin", "invited_by": ""}
        ]
        save_members(spreadsheet, members)
        st.session_state.user = {"email": email, "name": name, "role": "admin"}
        st.success("Admin-Account erstellt!")
        st.rerun()


def show_login(members):
    st.title("📋 Team Task Manager — Login")
    with st.form("login_form"):
        email = st.text_input("Email")
        password = st.text_input("Passwort", type="password")
        submitted = st.form_submit_button("Einloggen")

    if submitted:
        if not email or not password:
            st.warning("Bitte Email und Passwort eingeben.")
            return

        member = next((m for m in members if m["email"] == email), None)
        if member is None:
            st.error("Kein Mitarbeiter mit dieser Email gefunden.")
            return

        if member.get("password", DEFAULT_PASSWORD) != password:
            st.error("Falsches Passwort.")
            return

        st.session_state.user = {
            "email": member["email"],
            "name": member["name"],
            "role": member.get("role", "member"),
        }
        st.rerun()


# ─── Sidebar ──────────────────────────────────────────────────


def render_sidebar(spreadsheet, members, tasks, user):
    user_email = user["email"]
    user_role = user["role"]

    st.sidebar.write(f"👤 **{user['name']}**")
    st.sidebar.caption(f"{user['email']} ({user['role'].capitalize()})")

    if st.sidebar.button("🚪 Ausloggen"):
        del st.session_state.user
        st.rerun()

    st.sidebar.divider()

    # Mitarbeiter verwalten
    st.sidebar.header("👥 Mein Team")

    with st.sidebar.form("add_member", clear_on_submit=True):
        st.subheader("Mitarbeiter einladen")
        new_name = st.text_input("Name")
        new_email = st.text_input("Email")
        if st.form_submit_button("Einladen"):
            if new_name and new_email:
                if any(m["email"] == new_email for m in members):
                    st.error("Email existiert bereits.")
                else:
                    members.append({
                        "name": new_name,
                        "email": new_email,
                        "password": DEFAULT_PASSWORD,
                        "role": "member",
                        "invited_by": user_email,
                    })
                    save_members(spreadsheet, members)
                    st.success(f"{new_name} eingeladen! (Passwort: {DEFAULT_PASSWORD})")
                    st.rerun()
            else:
                st.warning("Bitte Name und Email eingeben.")

    # Team anzeigen
    my_team = get_my_team(members, user_email) if user_role != "admin" else members
    if my_team:
        st.sidebar.subheader("Aktuelle Mitarbeiter")
        for i, m in enumerate(my_team):
            if m["email"] == user_email:
                continue
            col1, col2 = st.sidebar.columns([3, 1])
            role_badge = "👑" if m.get("role") == "admin" else ""
            col1.write(f"{role_badge} **{m['name']}**")
            col1.caption(m["email"])
            member_idx = members.index(m)
            if col2.button("✕", key=f"del_member_{member_idx}"):
                tasks[:] = [t for t in tasks if t["assigned_to"] != m["email"]]
                members.pop(member_idx)
                save_members(spreadsheet, members)
                save_tasks(spreadsheet, tasks)
                st.rerun()

    # Passwort ändern
    st.sidebar.divider()
    st.sidebar.header("🔑 Passwort ändern")
    with st.sidebar.form("change_pw"):
        old_pw = st.text_input("Altes Passwort", type="password")
        new_pw = st.text_input("Neues Passwort", type="password")
        new_pw2 = st.text_input("Neues Passwort bestätigen", type="password")
        if st.form_submit_button("Ändern"):
            member = next((m for m in members if m["email"] == user_email), None)
            if member is None:
                st.error("User nicht gefunden.")
            elif old_pw != member.get("password", DEFAULT_PASSWORD):
                st.error("Altes Passwort ist falsch.")
            elif not new_pw:
                st.warning("Bitte neues Passwort eingeben.")
            elif new_pw != new_pw2:
                st.error("Neue Passwörter stimmen nicht überein.")
            else:
                member["password"] = new_pw
                save_members(spreadsheet, members)
                st.success("Passwort geändert!")


# ─── Tab 1: Neue Aufgabe ──────────────────────────────────────


def tab_new_task(spreadsheet, members, tasks, user):
    assignable = get_assignable_members(members, user["email"], user["role"])

    if not assignable:
        st.info("Bitte lade zuerst Mitarbeiter in der Sidebar ein.")
        return

    with st.form("add_task", clear_on_submit=True):
        st.subheader("Neue Aufgabe erstellen")
        title = st.text_input("Titel *")
        description = st.text_area("Beschreibung")

        col1, col2 = st.columns(2)
        member_names = [f"{m['name']} ({m['email']})" for m in assignable]
        assigned = col1.selectbox("Zuweisen an *", member_names)
        category = col2.selectbox("Kategorie", CATEGORY_OPTIONS)

        col3, col4 = st.columns(2)
        priority = col3.selectbox("Priorität", PRIORITY_OPTIONS)
        due = col4.date_input("Fällig am", value=None)

        if st.form_submit_button("✅ Aufgabe erstellen", type="primary"):
            if title and assigned:
                member_idx = member_names.index(assigned)
                max_sort = max((t.get("sort_order", 0) for t in tasks), default=0)
                task = {
                    "title": title,
                    "description": description,
                    "assigned_to": assignable[member_idx]["email"],
                    "priority": priority,
                    "due": due.strftime("%d.%m.%Y") if due else "",
                    "created": date.today().strftime("%d.%m.%Y"),
                    "status": "Offen",
                    "comment": "",
                    "category": category,
                    "sort_order": max_sort + 1,
                }
                tasks.append(task)
                save_tasks(spreadsheet, tasks)
                st.success(f"Aufgabe '{title}' erstellt!")
                st.rerun()
            else:
                st.warning("Bitte Titel eingeben und Mitarbeiter auswählen.")


# ─── Tab 2: Meine Aufgaben ────────────────────────────────────


def tab_my_tasks(spreadsheet, tasks, user):
    my_tasks = [t for t in tasks if t["assigned_to"] == user["email"]]

    if not my_tasks:
        st.info("Du hast aktuell keine Aufgaben.")
        return

    # Filter
    col_f1, col_f2 = st.columns(2)
    filter_status = col_f1.multiselect(
        "Status-Filter", STATUS_OPTIONS, default=STATUS_OPTIONS, key="my_status_filter"
    )
    filter_cat = col_f2.multiselect(
        "Kategorie-Filter", CATEGORY_OPTIONS, default=CATEGORY_OPTIONS, key="my_cat_filter"
    )

    filtered = [
        t for t in my_tasks
        if t.get("status", "Offen") in filter_status
        and t.get("category", "Sonstiges") in filter_cat
    ]

    for i, task in enumerate(filtered):
        status = task.get("status", "Offen")
        s_color = STATUS_COLORS.get(status, "gray")
        cat = task.get("category", "Sonstiges")
        cat_color = CATEGORY_COLORS.get(cat, "#9E9E9E")

        with st.container(border=True):
            col1, col2, col3 = st.columns([5, 2, 2])
            col1.write(f"### {task['title']}")
            col2.markdown(f":{s_color}[{status}]")
            if task.get("due"):
                col3.write(f"📅 {task['due']}")

            tag_col1, tag_col2 = st.columns([1, 5])
            tag_col1.markdown(
                f'<span style="background:{cat_color};color:white;padding:2px 8px;'
                f'border-radius:12px;font-size:0.8em;">{cat}</span>',
                unsafe_allow_html=True,
            )
            if task.get("priority"):
                p_color = {"Hoch": "red", "Mittel": "orange", "Niedrig": "green"}.get(
                    task["priority"], "gray"
                )
                tag_col2.markdown(f"Priorität: :{p_color}[{task['priority']}]")

            if task.get("description"):
                st.write(f"_{task['description']}_")

            # Status + Kommentar bearbeiten
            st.divider()
            col_s, col_c = st.columns([1, 2])
            current_idx = STATUS_OPTIONS.index(status) if status in STATUS_OPTIONS else 0
            new_status = col_s.selectbox(
                "Status", STATUS_OPTIONS, index=current_idx, key=f"my_status_{i}"
            )
            new_comment = col_c.text_input(
                "Kommentar", value=task.get("comment", ""), key=f"my_comment_{i}"
            )

            if new_status != status or new_comment != task.get("comment", ""):
                if st.button("💾 Speichern", key=f"my_save_{i}"):
                    task["status"] = new_status
                    task["comment"] = new_comment
                    save_tasks(spreadsheet, tasks)
                    st.success("Gespeichert!")
                    st.rerun()


# ─── Tab 3: Team-Übersicht & Emails ───────────────────────────


def tab_team_overview(spreadsheet, members, tasks, user):
    app_url = get_app_url()
    user_email = user["email"]
    user_role = user["role"]

    visible_tasks = get_visible_tasks(tasks, members, user_email, user_role)
    team = get_assignable_members(members, user_email, user_role)

    if not team:
        st.info("Noch keine Mitarbeiter vorhanden.")
        return
    if not visible_tasks:
        st.info("Noch keine Aufgaben vorhanden.")
        return

    # Filter
    col_f1, col_f2 = st.columns(2)
    filter_status = col_f1.multiselect(
        "Status-Filter", STATUS_OPTIONS, default=STATUS_OPTIONS, key="team_status_filter"
    )
    filter_cat = col_f2.multiselect(
        "Kategorie-Filter", CATEGORY_OPTIONS, default=CATEGORY_OPTIONS, key="team_cat_filter"
    )

    for member in team:
        member_tasks = [
            t for t in visible_tasks
            if t["assigned_to"] == member["email"]
            and t.get("status", "Offen") in filter_status
            and t.get("category", "Sonstiges") in filter_cat
        ]
        if not member_tasks:
            continue

        # Sortieren
        member_tasks.sort(key=lambda t: t.get("sort_order", 0))

        with st.expander(
            f"**{member['name']}** — {len(member_tasks)} Aufgabe(n)", expanded=True
        ):
            for j, task in enumerate(member_tasks):
                task_idx = tasks.index(task)
                col1, col2, col3, col4, col5, col6 = st.columns([4, 1, 1, 1, 1, 1])

                # Titel
                col1.write(f"**{task['title']}**")

                # Status
                status = task.get("status", "Offen")
                color = STATUS_COLORS.get(status, "gray")
                col2.markdown(f":{color}[{status}]")

                # Kategorie
                cat = task.get("category", "Sonstiges")
                cat_color = CATEGORY_COLORS.get(cat, "#9E9E9E")
                col3.markdown(
                    f'<span style="background:{cat_color};color:white;padding:2px 6px;'
                    f'border-radius:10px;font-size:0.75em;">{cat}</span>',
                    unsafe_allow_html=True,
                )

                # Sortierung
                if col4.button("⬆", key=f"up_{task_idx}"):
                    if j > 0:
                        prev_task = member_tasks[j - 1]
                        task["sort_order"], prev_task["sort_order"] = (
                            prev_task["sort_order"], task["sort_order"]
                        )
                        save_tasks(spreadsheet, tasks)
                        st.rerun()

                if col5.button("⬇", key=f"down_{task_idx}"):
                    if j < len(member_tasks) - 1:
                        next_task = member_tasks[j + 1]
                        task["sort_order"], next_task["sort_order"] = (
                            next_task["sort_order"], task["sort_order"]
                        )
                        save_tasks(spreadsheet, tasks)
                        st.rerun()

                # Löschen
                if col6.button("🗑", key=f"del_task_{task_idx}"):
                    tasks.pop(task_idx)
                    save_tasks(spreadsheet, tasks)
                    st.rerun()

                # Details
                details = []
                if task.get("priority"):
                    p_color = {"Hoch": "red", "Mittel": "orange", "Niedrig": "green"}.get(
                        task["priority"], "gray"
                    )
                    details.append(f"Priorität: :{p_color}[{task['priority']}]")
                if task.get("due"):
                    details.append(f"📅 {task['due']}")
                if task.get("description"):
                    details.append(f"_{task['description']}_")
                if task.get("comment"):
                    details.append(f"💬 {task['comment']}")
                if details:
                    st.caption(" | ".join(details))

                # Bearbeiten
                with st.expander(f"✏️ Bearbeiten", expanded=False):
                    with st.form(f"edit_task_{task_idx}"):
                        edit_title = st.text_input("Titel", value=task["title"])
                        edit_desc = st.text_area("Beschreibung", value=task.get("description", ""))

                        ec1, ec2 = st.columns(2)
                        assignable = get_assignable_members(members, user_email, user_role)
                        assign_names = [f"{m['name']} ({m['email']})" for m in assignable]
                        current_assign = next(
                            (f"{m['name']} ({m['email']})" for m in assignable
                             if m["email"] == task["assigned_to"]),
                            assign_names[0] if assign_names else "",
                        )
                        assign_idx = assign_names.index(current_assign) if current_assign in assign_names else 0
                        edit_assign = ec1.selectbox(
                            "Zuweisen an", assign_names, index=assign_idx, key=f"edit_assign_{task_idx}"
                        )
                        edit_cat = ec2.selectbox(
                            "Kategorie", CATEGORY_OPTIONS,
                            index=CATEGORY_OPTIONS.index(task.get("category", "Sonstiges"))
                            if task.get("category", "Sonstiges") in CATEGORY_OPTIONS else 0,
                            key=f"edit_cat_{task_idx}",
                        )

                        ec3, ec4 = st.columns(2)
                        edit_priority = ec3.selectbox(
                            "Priorität", PRIORITY_OPTIONS,
                            index=PRIORITY_OPTIONS.index(task.get("priority", "Mittel"))
                            if task.get("priority", "Mittel") in PRIORITY_OPTIONS else 0,
                            key=f"edit_prio_{task_idx}",
                        )
                        edit_status = ec4.selectbox(
                            "Status", STATUS_OPTIONS,
                            index=STATUS_OPTIONS.index(task.get("status", "Offen"))
                            if task.get("status", "Offen") in STATUS_OPTIONS else 0,
                            key=f"edit_status_{task_idx}",
                        )

                        try:
                            due_val = datetime.strptime(task["due"], "%d.%m.%Y").date() if task.get("due") else None
                        except ValueError:
                            due_val = None
                        edit_due = st.date_input("Fällig am", value=due_val, key=f"edit_due_{task_idx}")

                        if st.form_submit_button("💾 Änderungen speichern"):
                            a_idx = assign_names.index(edit_assign)
                            task["title"] = edit_title
                            task["description"] = edit_desc
                            task["assigned_to"] = assignable[a_idx]["email"]
                            task["priority"] = edit_priority
                            task["status"] = edit_status
                            task["category"] = edit_cat
                            task["due"] = edit_due.strftime("%d.%m.%Y") if edit_due else ""
                            save_tasks(spreadsheet, tasks)
                            st.success("Gespeichert!")
                            st.rerun()

            # Email-Button
            all_member_tasks = [t for t in tasks if t["assigned_to"] == member["email"]]
            if all_member_tasks:
                mailto = build_mailto_link(member["email"], member["name"], all_member_tasks, app_url)
                st.markdown(
                    f'<a href="{mailto}" target="_blank" style="display:inline-block;'
                    f"padding:8px 16px;background:#4CAF50;color:white;"
                    f'border-radius:4px;text-decoration:none;margin-top:8px;">'
                    f"📧 Aufgaben an {member['name']} senden</a>",
                    unsafe_allow_html=True,
                )

    st.divider()
    if st.button("📧 Alle Emails auf einmal öffnen"):
        links = []
        for member in team:
            member_tasks = [t for t in visible_tasks if t["assigned_to"] == member["email"]]
            if member_tasks:
                links.append(build_mailto_link(member["email"], member["name"], member_tasks, app_url))
        js = "".join(f'window.open("{link}");' for link in links)
        st.components.v1.html(f"<script>{js}</script>", height=0)


# ─── Tab 4: Antworten importieren ─────────────────────────────


def tab_import_replies(spreadsheet, members, tasks, user):
    st.subheader("📥 Email-Antwort importieren")
    st.write(
        "Füge die Antwort-Email eines Mitarbeiters hier ein. "
        "Der Status und Kommentar werden automatisch erkannt."
    )

    team = get_assignable_members(members, user["email"], user["role"])
    if not team:
        st.info("Noch keine Mitarbeiter vorhanden.")
        return

    member_names = [f"{m['name']} ({m['email']})" for m in team]
    selected_member = st.selectbox(
        "Antwort von Mitarbeiter", member_names, key="import_member"
    )
    member_idx = member_names.index(selected_member)
    member_email = team[member_idx]["email"]

    reply_text = st.text_area(
        "Email-Antwort hier einfügen",
        height=300,
        key="reply_text",
        placeholder="--- AUFGABEN ---\n[1] Website Redesign [Hoch]\n    Status: In Arbeit\n    Kommentar: Bin dran\n--- ENDE ---",
    )

    if reply_text:
        updates = parse_reply(reply_text)
        member_tasks = [t for t in tasks if t["assigned_to"] == member_email]

        if not updates:
            st.warning("Keine Aufgaben-Updates erkannt.")
        else:
            st.subheader("Erkannte Änderungen")
            valid_updates = []
            for u in updates:
                task_num = u["index"]
                if task_num < 1 or task_num > len(member_tasks):
                    st.warning(f"Aufgabe [{task_num}] existiert nicht — übersprungen.")
                    continue

                task = member_tasks[task_num - 1]
                old_status = task.get("status", "Offen")
                old_comment = task.get("comment", "")
                new_status = u["status"]
                new_comment = u["comment"]

                is_valid = new_status in STATUS_OPTIONS
                s_changed = old_status != new_status
                c_changed = old_comment != new_comment

                if s_changed or c_changed:
                    col1, col2 = st.columns([1, 1])
                    col1.write(f"**[{task_num}] {task['title']}**")
                    changes = []
                    if s_changed:
                        if is_valid:
                            changes.append(f"Status: {old_status} → **{new_status}**")
                        else:
                            changes.append(f"⚠️ Unbekannter Status: '{new_status}'")
                    if c_changed:
                        changes.append(f"Kommentar: _{new_comment}_")
                    col2.write(" | ".join(changes))

                    if is_valid:
                        valid_updates.append(
                            {"task_num": task_num, "status": new_status, "comment": new_comment}
                        )
                else:
                    st.write(f"[{task_num}] {task['title']} — keine Änderung")

            if valid_updates:
                if st.button(
                    f"✅ {len(valid_updates)} Änderung(en) übernehmen", type="primary"
                ):
                    for u in valid_updates:
                        task = member_tasks[u["task_num"] - 1]
                        task["status"] = u["status"]
                        task["comment"] = u["comment"]
                    save_tasks(spreadsheet, tasks)
                    st.success("Änderungen übernommen!")
                    st.rerun()
            elif updates:
                st.info("Keine gültigen Änderungen erkannt.")


# ─── Main ───────────────────────────────────────────────────────


def main():
    st.set_page_config(page_title="Team Task Manager", page_icon="📋", layout="wide")

    spreadsheet = get_gsheet_connection()
    ensure_worksheets(spreadsheet)

    members = load_members(spreadsheet)
    tasks = load_tasks(spreadsheet)

    # Ersteinrichtung
    if not members:
        show_setup(spreadsheet)
        return

    # Login
    if "user" not in st.session_state:
        show_login(members)
        return

    user = st.session_state.user

    # Header
    st.title("📋 Team Task Manager")

    # Sidebar
    render_sidebar(spreadsheet, members, tasks, user)

    # Tabs
    tab1, tab2, tab3, tab4 = st.tabs(
        ["➕ Neue Aufgabe", "📋 Meine Aufgaben", "📊 Team & Emails", "📥 Import"]
    )

    with tab1:
        tab_new_task(spreadsheet, members, tasks, user)
    with tab2:
        tab_my_tasks(spreadsheet, tasks, user)
    with tab3:
        tab_team_overview(spreadsheet, members, tasks, user)
    with tab4:
        tab_import_replies(spreadsheet, members, tasks, user)


if __name__ == "__main__":
    main()
