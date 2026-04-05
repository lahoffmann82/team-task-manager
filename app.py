import streamlit as st
import json
import re
import urllib.parse
from datetime import date

from google.oauth2.service_account import Credentials
import gspread

PRIORITY_OPTIONS = ["Hoch", "Mittel", "Niedrig"]
STATUS_OPTIONS = ["Offen", "In Arbeit", "Erledigt"]
STATUS_COLORS = {"Offen": "gray", "In Arbeit": "blue", "Erledigt": "green"}
LOGIN_CODE = "12345"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ─── Google Sheets Verbindung ──────────────────────────────────


@st.cache_resource
def get_gsheet_connection():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(st.secrets["spreadsheet"]["key"])
    return spreadsheet


def ensure_worksheets(spreadsheet):
    existing = [ws.title for ws in spreadsheet.worksheets()]
    if "members" not in existing:
        ws = spreadsheet.add_worksheet(title="members", rows=100, cols=2)
        ws.update("A1:B1", [["name", "email"]])
    if "tasks" not in existing:
        ws = spreadsheet.add_worksheet(title="tasks", rows=500, cols=8)
        ws.update("A1:H1", [["title", "description", "assigned_to", "priority", "due", "created", "status", "comment"]])


def load_members(spreadsheet):
    ws = spreadsheet.worksheet("members")
    records = ws.get_all_records()
    return records


def load_tasks(spreadsheet):
    ws = spreadsheet.worksheet("tasks")
    records = ws.get_all_records()
    for t in records:
        if not t.get("status"):
            t["status"] = "Offen"
        if not t.get("comment"):
            t["comment"] = ""
    return records


def save_members(spreadsheet, members):
    ws = spreadsheet.worksheet("members")
    ws.clear()
    ws.update("A1:B1", [["name", "email"]])
    if members:
        rows = [[m["name"], m["email"]] for m in members]
        ws.update(f"A2:B{len(rows)+1}", rows)


def save_tasks(spreadsheet, tasks):
    ws = spreadsheet.worksheet("tasks")
    ws.clear()
    ws.update("A1:H1", [["title", "description", "assigned_to", "priority", "due", "created", "status", "comment"]])
    if tasks:
        rows = [
            [
                t.get("title", ""),
                t.get("description", ""),
                t.get("assigned_to", ""),
                t.get("priority", ""),
                t.get("due", ""),
                t.get("created", ""),
                t.get("status", "Offen"),
                t.get("comment", ""),
            ]
            for t in tasks
        ]
        ws.update(f"A2:H{len(rows)+1}", rows)


# ─── Email-Funktionen ─────────────────────────────────────────


def build_email_body(name, tasks):
    lines = [
        f"Hallo {name},",
        "",
        "hier sind deine aktuellen Aufgaben.",
        "Bitte aktualisiere den Status und Kommentar und sende diese Email als Antwort zurueck.",
        "",
        "--- AUFGABEN ---",
    ]
    for i, t in enumerate(tasks, 1):
        priority = f" [{t['priority']}]" if t.get("priority") else ""
        due = f" (Faellig: {t['due']})" if t.get("due") else ""
        lines.append(f"[{i}] {t['title']}{priority}{due}")
        if t.get("description"):
            lines.append(f"    Beschreibung: {t['description']}")
        lines.append(f"    Status: {t.get('status', 'Offen')}")
        lines.append(f"    Kommentar: {t.get('comment', '')}")
        lines.append("")
    lines.append("--- ENDE ---")
    lines.append("")
    lines.append("Viele Gruesse")
    return "\n".join(lines)


def build_mailto_link(email, name, tasks):
    subject = f"Deine Aufgaben - {date.today().strftime('%d.%m.%Y')}"
    body = build_email_body(name, tasks)
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
            {
                "index": idx,
                "title": title,
                "status": status_normalized,
                "comment": comment,
            }
        )
    return updates


# ─── Login ─────────────────────────────────────────────────────


def show_login(members):
    st.title("📋 Team Task Manager — Login")
    with st.form("login_form"):
        email = st.text_input("Email")
        code = st.text_input("Code", type="password")
        submitted = st.form_submit_button("Einloggen")

    if submitted:
        if not email or not code:
            st.warning("Bitte Email und Code eingeben.")
            return

        if code != LOGIN_CODE:
            st.error("Falscher Code.")
            return

        member = next((m for m in members if m["email"] == email), None)
        if member is None:
            st.error("Kein Mitarbeiter mit dieser Email gefunden.")
            return

        st.session_state.user = {
            "email": member["email"],
            "name": member["name"],
        }
        st.rerun()


# ─── Hauptansicht ──────────────────────────────────────────────


def app_view(spreadsheet, members, tasks):
    # Sidebar: Mitarbeiter verwalten
    st.sidebar.header("👥 Mitarbeiter")

    with st.sidebar.form("add_member", clear_on_submit=True):
        st.subheader("Mitarbeiter hinzufügen")
        new_name = st.text_input("Name")
        new_email = st.text_input("Email")
        if st.form_submit_button("Hinzufügen"):
            if new_name and new_email:
                if any(m["email"] == new_email for m in members):
                    st.error("Email existiert bereits.")
                else:
                    members.append({"name": new_name, "email": new_email})
                    save_members(spreadsheet, members)
                    st.rerun()
            else:
                st.warning("Bitte Name und Email eingeben.")

    if members:
        st.sidebar.subheader("Aktuelle Mitarbeiter")
        for i, m in enumerate(members):
            col1, col2 = st.sidebar.columns([3, 1])
            col1.write(f"**{m['name']}** ({m['email']})")
            if col2.button("✕", key=f"del_member_{i}"):
                tasks[:] = [t for t in tasks if t["assigned_to"] != m["email"]]
                members.pop(i)
                save_members(spreadsheet, members)
                save_tasks(spreadsheet, tasks)
                st.rerun()

    # Tabs
    tab1, tab2, tab3 = st.tabs(
        ["➕ Neue Aufgabe", "📊 Übersicht & Emails", "📥 Antworten importieren"]
    )

    # Tab 1: Aufgabe erstellen
    with tab1:
        if not members:
            st.info("Bitte zuerst Mitarbeiter in der Sidebar anlegen.")
        else:
            with st.form("add_task", clear_on_submit=True):
                st.subheader("Neue Aufgabe erstellen")
                title = st.text_input("Titel *")
                description = st.text_area("Beschreibung")
                col1, col2, col3 = st.columns(3)
                member_names = [f"{m['name']} ({m['email']})" for m in members]
                assigned = col1.selectbox("Zuweisen an *", member_names)
                priority = col2.selectbox("Priorität", PRIORITY_OPTIONS)
                due = col3.date_input("Fällig am", value=None)

                if st.form_submit_button("Aufgabe erstellen"):
                    if title and assigned:
                        member_idx = member_names.index(assigned)
                        task = {
                            "title": title,
                            "description": description,
                            "assigned_to": members[member_idx]["email"],
                            "priority": priority,
                            "due": due.strftime("%d.%m.%Y") if due else "",
                            "created": date.today().strftime("%d.%m.%Y"),
                            "status": "Offen",
                            "comment": "",
                        }
                        tasks.append(task)
                        save_tasks(spreadsheet, tasks)
                        st.success(f"Aufgabe '{title}' erstellt!")
                        st.rerun()
                    else:
                        st.warning("Bitte Titel eingeben und Mitarbeiter auswählen.")

    # Tab 2: Übersicht & Emails
    with tab2:
        if not members:
            st.info("Noch keine Mitarbeiter vorhanden.")
        elif not tasks:
            st.info("Noch keine Aufgaben vorhanden.")
        else:
            filter_status = st.multiselect(
                "Nach Status filtern", STATUS_OPTIONS, default=STATUS_OPTIONS
            )

            for member in members:
                member_tasks = [
                    t
                    for t in tasks
                    if t["assigned_to"] == member["email"]
                    and t.get("status", "Offen") in filter_status
                ]
                if not member_tasks:
                    continue

                with st.expander(
                    f"**{member['name']}** — {len(member_tasks)} Aufgabe(n)",
                    expanded=True,
                ):
                    for j, task in enumerate(member_tasks):
                        col1, col2, col3, col4 = st.columns([4, 2, 2, 1])
                        col1.write(f"**{task['title']}**")

                        status = task.get("status", "Offen")
                        color = STATUS_COLORS.get(status, "gray")
                        col2.markdown(f":{color}[{status}]")

                        if task.get("due"):
                            col3.write(f"📅 {task['due']}")

                        task_idx = tasks.index(task)
                        if col4.button("🗑", key=f"del_task_{task_idx}"):
                            tasks.pop(task_idx)
                            save_tasks(spreadsheet, tasks)
                            st.rerun()

                        details = []
                        if task.get("priority"):
                            p_color = {
                                "Hoch": "red",
                                "Mittel": "orange",
                                "Niedrig": "green",
                            }.get(task["priority"], "gray")
                            details.append(
                                f"Priorität: :{p_color}[{task['priority']}]"
                            )
                        if task.get("description"):
                            details.append(f"_{task['description']}_")
                        if task.get("comment"):
                            details.append(f"💬 **Kommentar:** {task['comment']}")
                        if details:
                            st.caption(" | ".join(details))

                    all_member_tasks = [
                        t for t in tasks if t["assigned_to"] == member["email"]
                    ]
                    mailto = build_mailto_link(
                        member["email"], member["name"], all_member_tasks
                    )
                    st.markdown(
                        f'<a href="{mailto}" target="_blank" style="display:inline-block;'
                        f"padding:8px 16px;background:#4CAF50;color:white;"
                        f'border-radius:4px;text-decoration:none;">'
                        f"📧 Email an {member['name']} senden</a>",
                        unsafe_allow_html=True,
                    )

            st.divider()
            if st.button("📧 Alle Emails auf einmal öffnen"):
                links = []
                for member in members:
                    member_tasks = [
                        t for t in tasks if t["assigned_to"] == member["email"]
                    ]
                    if member_tasks:
                        links.append(
                            build_mailto_link(
                                member["email"], member["name"], member_tasks
                            )
                        )
                js = "".join(f'window.open("{link}");' for link in links)
                st.components.v1.html(f"<script>{js}</script>", height=0)

    # Tab 3: Antworten importieren
    with tab3:
        st.subheader("📥 Email-Antwort importieren")
        st.write(
            "Füge die Antwort-Email eines Mitarbeiters hier ein. "
            "Der Status und Kommentar werden automatisch erkannt."
        )

        if not members:
            st.info("Noch keine Mitarbeiter vorhanden.")
        else:
            member_names = [f"{m['name']} ({m['email']})" for m in members]
            selected_member = st.selectbox(
                "Antwort von Mitarbeiter", member_names, key="import_member"
            )
            member_idx = member_names.index(selected_member)
            member_email = members[member_idx]["email"]

            reply_text = st.text_area(
                "Email-Antwort hier einfügen",
                height=300,
                key="reply_text",
                placeholder="--- AUFGABEN ---\n[1] Website Redesign [Hoch]\n    Status: In Arbeit\n    Kommentar: Bin dran, wird bis Freitag fertig\n--- ENDE ---",
            )

            if reply_text:
                updates = parse_reply(reply_text)
                member_tasks = [
                    t for t in tasks if t["assigned_to"] == member_email
                ]

                if not updates:
                    st.warning(
                        "Keine Aufgaben-Updates erkannt. "
                        "Bitte stelle sicher, dass das Format korrekt ist."
                    )
                else:
                    st.subheader("Erkannte Änderungen")
                    valid_updates = []
                    for u in updates:
                        task_num = u["index"]
                        if task_num < 1 or task_num > len(member_tasks):
                            st.warning(
                                f"Aufgabe [{task_num}] existiert nicht — wird übersprungen."
                            )
                            continue

                        task = member_tasks[task_num - 1]
                        old_status = task.get("status", "Offen")
                        old_comment = task.get("comment", "")
                        new_status = u["status"]
                        new_comment = u["comment"]

                        is_valid_status = new_status in STATUS_OPTIONS
                        status_changed = old_status != new_status
                        comment_changed = old_comment != new_comment

                        if status_changed or comment_changed:
                            col1, col2 = st.columns([1, 1])
                            with col1:
                                st.write(f"**[{task_num}] {task['title']}**")
                            with col2:
                                changes = []
                                if status_changed:
                                    if is_valid_status:
                                        changes.append(
                                            f"Status: {old_status} → **{new_status}**"
                                        )
                                    else:
                                        changes.append(
                                            f"⚠️ Unbekannter Status: '{new_status}'"
                                        )
                                if comment_changed:
                                    changes.append(f"Kommentar: _{new_comment}_")
                                st.write(" | ".join(changes))

                            if is_valid_status:
                                valid_updates.append(
                                    {
                                        "task_num": task_num,
                                        "status": new_status,
                                        "comment": new_comment,
                                    }
                                )
                        else:
                            st.write(
                                f"[{task_num}] {task['title']} — keine Änderung"
                            )

                    if valid_updates:
                        if st.button(
                            f"✅ {len(valid_updates)} Änderung(en) übernehmen",
                            type="primary",
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

    # Not logged in
    if "user" not in st.session_state:
        show_login(members)
        return

    user = st.session_state.user

    # Header + logout
    st.title("📋 Team Task Manager")
    st.sidebar.write(f"Eingeloggt als **{user['name']}**")
    if st.sidebar.button("🚪 Ausloggen"):
        del st.session_state.user
        st.rerun()

    app_view(spreadsheet, members, tasks)


if __name__ == "__main__":
    main()
