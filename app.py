from __future__ import annotations

import os
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path
import re
import unicodedata
from urllib.parse import urlencode
from zoneinfo import ZoneInfo

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "oncologia_cuidado.db"
DATE_FMT = "%Y-%m-%d"
GOOGLE_SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
DEFAULT_SPREADSHEET_ID = "1W1FKPD-F5Fmq4it8vT9_2vFwZX_SgtkD6m0-bmBABaM"
AUTO_SYNC_MINUTES = 5
PRIMARY_WORKBOOK_NAME = "PLANILHA DE PRESCRIÇÕES - MÉDICOS JULIANA - Copiar.xlsx"
UPLOADED_WORKBOOK_NAME = "cloud_primary_workbook.xlsx"
ONEDRIVE_CLOUDSTORAGE_DIR = Path.home() / "Library" / "CloudStorage"
APP_TIMEZONE = ZoneInfo("America/Sao_Paulo")
MONTH_LABELS_PT = {
    1: "janeiro",
    2: "fevereiro",
    3: "marco",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro",
}


def format_month_label_pt(value: str) -> str:
    month_ref = datetime.strptime(value, "%Y-%m")
    return f"{MONTH_LABELS_PT[month_ref.month]}/{month_ref.year}"

STATUS_LABELS = {
    "scheduled": "Programado",
    "done": "Realizado",
    "pending": "Pendente",
    "attention": "Atencao",
}

PRESCRIPTION_LABELS = {
    "not_requested": "Não solicitada",
    "requested": "Solicitada ao médico",
    "prescribed": "Prescrita",
    "sent_to_insurance": "Enviada ao convênio",
}

AUTHORIZATION_LABELS = {
    "not_sent": "Não enviada",
    "pending": "Em análise",
    "authorized": "Autorizada",
    "denied": "Negada",
}

SCHEDULING_LABELS = {
    "not_booked": "Sem agenda",
    "awaiting_slot": "Aguardando vaga",
    "scheduled": "Agendado",
    "confirmed": "Confirmado",
}

PATIENT_COLUMNS = {
    "insurance_name": "TEXT",
    "prescription_status": "TEXT NOT NULL DEFAULT 'not_requested'",
    "prescription_requested_date": "TEXT",
    "authorization_status": "TEXT NOT NULL DEFAULT 'not_sent'",
    "authorization_submission_date": "TEXT",
    "authorization_valid_until": "TEXT",
    "scheduling_status": "TEXT NOT NULL DEFAULT 'not_booked'",
    "scheduled_cycle_date": "TEXT",
    "next_cycle_alert_days": "INTEGER NOT NULL DEFAULT 7",
    "protocol_next_cycle_date": "TEXT",
    "source_sheet_name": "TEXT",
    "source_row_number": "INTEGER",
}

CHEMO_SESSION_COLUMNS = {
    "prescription_status": "TEXT NOT NULL DEFAULT 'not_requested'",
    "authorization_status": "TEXT NOT NULL DEFAULT 'not_sent'",
    "scheduling_status": "TEXT NOT NULL DEFAULT 'not_booked'",
}


st.set_page_config(
    page_title="Navegação Oncológica",
    page_icon="stethoscope",
    layout="wide",
)


APP_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Manrope', sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(3, 105, 161, 0.12), transparent 28%),
        radial-gradient(circle at top right, rgba(14, 116, 144, 0.10), transparent 24%),
        linear-gradient(180deg, #f7fbfd 0%, #edf6f8 100%);
}

.hero {
    background: linear-gradient(135deg, #0f3d4c 0%, #16697a 58%, #2e8fa3 100%);
    border-radius: 26px;
    color: #ffffff !important;
    padding: 28px 30px;
    box-shadow: 0 22px 48px rgba(15, 61, 76, 0.18);
    margin-bottom: 18px;
}

.hero, .hero * {
    color: #ffffff !important;
}

.hero h1 {
    margin: 0;
    font-size: 2.6rem;
    font-weight: 800;
    color: #ffffff !important;
    text-shadow: 0 1px 2px rgba(0, 0, 0, 0.08);
}

.hero p {
    margin: 12px 0 0 0;
    max-width: 940px;
    color: #f3fbff !important;
    line-height: 1.55;
}

.login-shell {
    max-width: 460px;
    margin: 38px auto 0 auto;
}

.login-card {
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid rgba(15, 61, 76, 0.08);
    border-radius: 24px;
    padding: 24px 24px 10px 24px;
    box-shadow: 0 16px 38px rgba(15, 61, 76, 0.10);
}

.login-title {
    font-size: 1.2rem;
    font-weight: 800;
    color: #123847;
    margin-bottom: 6px;
}

.login-copy {
    color: #56707a;
    line-height: 1.5;
    margin-bottom: 10px;
}

.login-card .stButton button,
.login-card .stForm button,
.login-card .stForm [data-testid="stFormSubmitButton"] button {
    background: linear-gradient(135deg, #ffffff 0%, #f4fbff 100%) !important;
    color: #0f3d4c !important;
    border: 1px solid rgba(15, 61, 76, 0.18) !important;
    box-shadow: 0 10px 20px rgba(15, 61, 76, 0.10) !important;
    font-weight: 800 !important;
}

.login-card .stButton button p,
.login-card .stButton button span,
.login-card .stForm button p,
.login-card .stForm button span,
.login-card .stForm [data-testid="stFormSubmitButton"] button p,
.login-card .stForm [data-testid="stFormSubmitButton"] button span {
    color: #0f3d4c !important;
}

.panel {
    background: rgba(255, 255, 255, 0.84);
    border: 1px solid rgba(15, 61, 76, 0.08);
    border-radius: 22px;
    padding: 18px 20px;
    box-shadow: 0 14px 36px rgba(26, 55, 77, 0.08);
}

.metric-card {
    border-radius: 22px;
    padding: 18px 20px;
    color: white;
    min-height: 150px;
    margin-bottom: 14px;
    box-shadow: 0 14px 30px rgba(16, 33, 54, 0.10);
}

.metric-a { background: linear-gradient(135deg, #0f3d4c 0%, #16697a 100%); }
.metric-b { background: linear-gradient(135deg, #6b3f1d 0%, #d97706 100%); }
.metric-c { background: linear-gradient(135deg, #7f1d1d 0%, #dc2626 100%); }
.metric-d { background: linear-gradient(135deg, #1d4d4f 0%, #2a9d8f 100%); }
.metric-protocol { background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 55%, #60a5fa 100%); }

.metric-label {
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-size: 0.76rem;
    opacity: 0.92;
    font-weight: 700;
    color: #ffffff !important;
}

.metric-value {
    font-size: 2rem;
    font-weight: 800;
    margin-top: 12px;
    color: #ffffff !important;
}

.metric-copy {
    margin-top: 12px;
    line-height: 1.45;
    font-size: 0.92rem;
    color: rgba(255,255,255,0.92) !important;
}

.section-title {
    font-size: 1.1rem;
    font-weight: 800;
    color: #123847;
    margin-bottom: 8px;
}

.subtle {
    color: #56707a;
    line-height: 1.5;
}

.flag {
    display: inline-block;
    padding: 6px 10px;
    border-radius: 999px;
    font-size: 0.82rem;
    font-weight: 700;
}

.flag-red { background: #fee2e2; color: #991b1b; }
.flag-amber { background: #fef3c7; color: #92400e; }
.flag-green { background: #dcfce7; color: #166534; }
.flag-blue { background: #dbeafe; color: #1d4ed8; }
.flag-protocol-strong { background: #dbeafe; color: #1d4ed8; }
.flag-protocol-soft { background: #e0e7ff; color: #3730a3; }

.section-chip {
    display: inline-block;
    padding: 5px 10px;
    border-radius: 999px;
    font-size: 0.78rem;
    font-weight: 800;
    margin-bottom: 8px;
}

.section-chip-operational {
    background: #fee2e2;
    color: #991b1b;
}

.section-chip-protocol {
    background: #dbeafe;
    color: #1d4ed8;
}

.protocol-card {
    background: linear-gradient(180deg, rgba(239, 246, 255, 0.94) 0%, rgba(219, 234, 254, 0.88) 100%);
    border: 1px solid rgba(37, 99, 235, 0.14);
    border-left: 4px solid #2563eb;
    border-radius: 16px;
    padding: 12px 14px;
    margin-bottom: 10px;
}

.protocol-card-title {
    font-weight: 800;
    color: #123847;
}

.protocol-card-copy {
    color: #31556a;
    font-size: 0.92rem;
    margin-top: 4px;
}

.calendar-grid {
    display: grid;
    grid-template-columns: repeat(7, minmax(0, 1fr));
    gap: 10px;
}

.calendar-head {
    font-size: 0.82rem;
    font-weight: 800;
    color: #35515b;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    padding: 4px 2px;
}

.calendar-day {
    min-height: 138px;
    background: rgba(255, 255, 255, 0.92);
    border: 1px solid rgba(15, 61, 76, 0.08);
    border-radius: 18px;
    padding: 10px;
    box-shadow: 0 10px 24px rgba(26, 55, 77, 0.05);
}

.calendar-day.muted {
    background: rgba(245, 248, 250, 0.92);
    color: #8aa0a8;
}

.calendar-date {
    font-size: 0.92rem;
    font-weight: 800;
    color: #123847;
    margin-bottom: 8px;
}

.calendar-count {
    display: inline-block;
    font-size: 0.76rem;
    font-weight: 800;
    color: #9a3412;
    background: #ffedd5;
    border-radius: 999px;
    padding: 4px 8px;
    margin-bottom: 8px;
}

.calendar-item {
    font-size: 0.8rem;
    line-height: 1.35;
    background: #e0f2f7;
    color: #164e63;
    border-radius: 12px;
    padding: 6px 8px;
    margin-bottom: 6px;
}

.calendar-item.more {
    background: #f3f4f6;
    color: #4b5563;
}

.calendar-link-card {
    background: rgba(255, 255, 255, 0.94);
    color: #164e63;
    border: 1px solid rgba(15, 61, 76, 0.08);
    border-radius: 14px;
    padding: 7px 8px 6px 8px;
    margin-bottom: 5px;
    line-height: 1.15;
    box-shadow: 0 6px 14px rgba(15, 61, 76, 0.06);
}

.calendar-link-meta {
    color: #6b7f88;
    font-size: 0.68rem;
    margin-top: 1px;
    margin-bottom: 0;
}

.calendar-doctor-group {
    font-size: 0.7rem;
    font-weight: 700;
    color: #5b7681;
    letter-spacing: 0.01em;
    margin: 6px 0 3px 2px;
}

.stApp, .stApp p, .stApp span, .stApp label {
    color: #123847;
}

[data-testid="stForm"] {
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid rgba(15, 61, 76, 0.08);
    border-radius: 18px;
    padding: 14px 16px 6px 16px;
}

[data-testid="stForm"] label,
[data-testid="stWidgetLabel"] p,
[data-testid="stMarkdownContainer"] p,
.stCheckbox label p {
    color: #123847 !important;
    font-weight: 700 !important;
}

.stSelectbox > div > div,
.stDateInput > div > div,
.stTextInput > div > div,
.stTextArea textarea {
    background: #ffffff !important;
    color: #123847 !important;
    border: 1px solid rgba(15, 61, 76, 0.16) !important;
}

.stTextArea textarea::placeholder,
input::placeholder {
    color: #6b7f88 !important;
}

.stCheckbox [data-baseweb="checkbox"] {
    background: #ffffff !important;
}

.stCheckbox svg {
    fill: #0f4c5c !important;
}

.stButton button, .stForm button[kind="secondaryFormSubmit"] {
    background: #0f3d4c !important;
    color: #ffffff !important;
    border: 1px solid #0f3d4c !important;
}

.stButton button p, .stButton button span {
    color: #ffffff !important;
}

[data-testid="stButton"] button[kind="tertiary"] {
    background: transparent !important;
    border: none !important;
    color: #0f4c5c !important;
    text-decoration: underline !important;
    font-weight: 800 !important;
    font-size: 0.72rem !important;
    padding: 0 !important;
    min-height: auto !important;
    box-shadow: none !important;
    justify-content: flex-start !important;
    line-height: 1.1 !important;
}

[data-testid="stButton"] button[kind="tertiary"] p,
[data-testid="stButton"] button[kind="tertiary"] span {
    color: #0f4c5c !important;
}
</style>
"""


def get_connection() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def load_access_credentials() -> dict[str, str]:
    secret_username = None
    secret_password = None
    try:
        auth_secrets = st.secrets.get("auth", {})
        secret_username = auth_secrets.get("username")
        secret_password = auth_secrets.get("password")
    except Exception:
        auth_secrets = {}

    username = secret_username or os.getenv("ONCO_APP_USERNAME") or "juliana"
    password = secret_password or os.getenv("ONCO_APP_PASSWORD") or "Navegacao2026!"
    return {"username": str(username), "password": str(password)}


def authenticate_access(username: str, password: str) -> bool:
    credentials = load_access_credentials()
    return username.strip().lower() == credentials["username"].strip().lower() and password == credentials["password"]


def ensure_auth_session_state() -> None:
    if "auth_user" not in st.session_state:
        st.session_state["auth_user"] = None


def render_login_gate() -> None:
    st.markdown(
        """
        <style>
            .stForm [data-testid="stFormSubmitButton"] button,
            .stForm [data-testid="stFormSubmitButton"] button:hover,
            .stForm [data-testid="stFormSubmitButton"] button:focus {
                background: linear-gradient(135deg, #ffffff 0%, #eef8fc 100%) !important;
                color: #0f3d4c !important;
                border: 1px solid rgba(15, 61, 76, 0.22) !important;
                box-shadow: 0 10px 20px rgba(15, 61, 76, 0.10) !important;
                font-weight: 800 !important;
            }

            .stForm [data-testid="stFormSubmitButton"] button p,
            .stForm [data-testid="stFormSubmitButton"] button span,
            .stForm [data-testid="stFormSubmitButton"] button:hover p,
            .stForm [data-testid="stFormSubmitButton"] button:hover span {
                color: #0f3d4c !important;
            }
        </style>
        <div class="hero">
            <h1 style="color:#ffffff !important;">Navegação Oncológica</h1>
            <p style="color:#f3fbff !important;">
                Acesso web protegido para acompanhar agenda, ciclos e autorizações dos pacientes.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Acesso ao painel</div>', unsafe_allow_html=True)
    st.caption("Entre com o usuário e a senha configurados para este aplicativo.")

    with st.form("login_gate_form"):
        username = st.text_input("Usuário")
        password = st.text_input("Senha", type="password")
        submitted = st.form_submit_button("Entrar", use_container_width=True)
        if submitted:
            if authenticate_access(username, password):
                st.session_state["auth_user"] = username.strip()
                st.success("Acesso liberado.")
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")

    st.markdown("</div>", unsafe_allow_html=True)


def init_db() -> None:
    conn = get_connection()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS doctors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            specialty TEXT,
            active INTEGER NOT NULL DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS patients (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            doctor_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            diagnosis TEXT,
            regimen TEXT,
            cycle_interval_days INTEGER NOT NULL DEFAULT 21,
            last_chemo_date TEXT,
            next_chemo_date TEXT,
            support_plan TEXT,
            notes TEXT,
            active INTEGER NOT NULL DEFAULT 1,
            insurance_name TEXT,
            prescription_status TEXT NOT NULL DEFAULT 'not_requested',
            prescription_requested_date TEXT,
            authorization_status TEXT NOT NULL DEFAULT 'not_sent',
            authorization_submission_date TEXT,
            authorization_valid_until TEXT,
            scheduling_status TEXT NOT NULL DEFAULT 'not_booked',
            scheduled_cycle_date TEXT,
            next_cycle_alert_days INTEGER NOT NULL DEFAULT 7,
            protocol_next_cycle_date TEXT,
            FOREIGN KEY (doctor_id) REFERENCES doctors(id)
        );

        CREATE TABLE IF NOT EXISTS support_medications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            patient_id INTEGER NOT NULL,
            medication_name TEXT NOT NULL,
            purpose TEXT,
            frequency_label TEXT,
            next_due_date TEXT,
            status TEXT NOT NULL DEFAULT 'pending',
            notes TEXT,
            FOREIGN KEY (patient_id) REFERENCES patients(id)
        );

        CREATE TABLE IF NOT EXISTS chemo_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            patient_id INTEGER NOT NULL,
            scheduled_date TEXT NOT NULL,
            cycle_label TEXT,
            status TEXT NOT NULL DEFAULT 'scheduled',
            notes TEXT,
            prescription_status TEXT NOT NULL DEFAULT 'not_requested',
            authorization_status TEXT NOT NULL DEFAULT 'not_sent',
            scheduling_status TEXT NOT NULL DEFAULT 'not_booked',
            FOREIGN KEY (patient_id) REFERENCES patients(id)
        );

        CREATE TABLE IF NOT EXISTS app_state (
            key TEXT PRIMARY KEY,
            value TEXT
        );
        """
    )
    ensure_patient_columns(conn)
    ensure_chemo_session_columns(conn)
    conn.commit()
    conn.close()


def ensure_patient_columns(conn: sqlite3.Connection) -> None:
    current = {
        row["name"]
        for row in conn.execute("PRAGMA table_info(patients)").fetchall()
    }
    for column_name, column_type in PATIENT_COLUMNS.items():
        if column_name not in current:
            conn.execute(f"ALTER TABLE patients ADD COLUMN {column_name} {column_type}")


def ensure_chemo_session_columns(conn: sqlite3.Connection) -> None:
    current = {
        row["name"]
        for row in conn.execute("PRAGMA table_info(chemo_sessions)").fetchall()
    }
    for column_name, column_type in CHEMO_SESSION_COLUMNS.items():
        if column_name not in current:
            conn.execute(f"ALTER TABLE chemo_sessions ADD COLUMN {column_name} {column_type}")
    conn.execute(
        """
        UPDATE chemo_sessions
        SET
            prescription_status = COALESCE(NULLIF(prescription_status, ''), 'not_requested'),
            authorization_status = COALESCE(NULLIF(authorization_status, ''), 'not_sent'),
            scheduling_status = COALESCE(NULLIF(scheduling_status, ''), 'not_booked')
        """
    )


def set_app_state(key: str, value: str) -> None:
    conn = get_connection()
    conn.execute(
        """
        INSERT INTO app_state (key, value)
        VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
        """,
        (key, value),
    )
    conn.commit()
    conn.close()


def get_app_state(key: str) -> str | None:
    conn = get_connection()
    row = conn.execute(
        "SELECT value FROM app_state WHERE key = ? LIMIT 1",
        (key,),
    ).fetchone()
    conn.close()
    return row["value"] if row else None


def parse_date(value: str | None) -> date | None:
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if not isinstance(value, str):
        return None
    value = value.strip()
    if not value:
        return None
    for fmt in (DATE_FMT, "%d/%m/%Y", "%d/%m/%y", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def format_date(value: str | None) -> str:
    parsed = parse_date(value)
    return parsed.strftime("%d/%m/%Y") if parsed else "-"


def format_sync_timestamp(value: str | None) -> str:
    if not value:
        return "ainda não sincronizado"
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError:
        return value
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone(APP_TIMEZONE)
    return parsed.strftime("%d/%m/%Y às %H:%M")


def days_until(value: str | None) -> int | None:
    parsed = parse_date(value)
    if not parsed:
        return None
    return (parsed - date.today()).days


def build_status_flag(delta_days: int | None) -> tuple[str, str]:
    if delta_days is None:
        return "Sem data", "flag-blue"
    if delta_days < 0:
        return f"Atrasado ha {abs(delta_days)} dia(s)", "flag-red"
    if delta_days <= 2:
        return f"Em {delta_days} dia(s)", "flag-amber"
    return f"Em {delta_days} dia(s)", "flag-green"


def format_status(value: str, mapping: dict[str, str]) -> str:
    return mapping.get(value, value)


def read_query_param(name: str) -> str | None:
    value = st.query_params.get(name)
    if value is None:
        return None
    if isinstance(value, list):
        return value[0] if value else None
    return str(value)


def sync_navigation_state_from_query_params() -> None:
    view = read_query_param("view")
    patient_id = read_query_param("patient_id")
    cycle_date = read_query_param("cycle_date")
    if view == "patient_detail" and patient_id:
        st.session_state["current_view"] = "patient_detail"
        st.session_state["selected_calendar_patient_id"] = int(patient_id)
        st.session_state["selected_calendar_cycle_date"] = cycle_date
    else:
        st.session_state["current_view"] = "main"
        st.session_state.pop("selected_calendar_cycle_date", None)


def open_patient_detail(patient_id: int, cycle_date: str) -> None:
    st.query_params.clear()
    st.query_params["view"] = "patient_detail"
    st.query_params["patient_id"] = str(patient_id)
    st.query_params["cycle_date"] = cycle_date
    st.session_state["current_view"] = "patient_detail"
    st.session_state["selected_calendar_patient_id"] = patient_id
    st.session_state["selected_calendar_cycle_date"] = cycle_date


def close_patient_detail() -> None:
    st.query_params.clear()
    st.session_state["current_view"] = "main"
    st.session_state.pop("selected_calendar_patient_id", None)
    st.session_state.pop("selected_calendar_cycle_date", None)


def normalize_header(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    normalized = "".join(char for char in normalized if not unicodedata.combining(char))
    normalized = normalized.lower().strip()
    normalized = re.sub(r"[^a-z0-9]+", "_", normalized)
    return normalized.strip("_")


def canonical_sheet_field(header: str) -> str:
    mapping = {
        "nome_do_paciente": "patient_name",
        "nome": "patient_name",
        "nome_": "patient_name",
        "paciente": "patient_name",
        "diagnostico": "diagnosis",
        "sitio": "diagnosis",
        "tto": "regimen",
        "tratamento": "regimen",
        "protocolo": "regimen",
        "proxima_infusao": "next_infusion",
        "data_da_infusao": "next_infusion",
        "data_de_infusao": "next_infusion",
        "data_do_tratamento": "next_infusion",
        "infusao": "next_infusion",
        "observacao": "notes",
        "observacoes": "notes",
        "solicitar_pm": "prescription_prompt",
        "medico": "doctor_name",
        "prontuario": "medical_record",
        "data_de_nascimento": "birth_date",
        "data_de_atencao": "attention_date",
        "proxima_consulta": "next_consultation_date",
        "ultima_consulta": "last_consultation_date",
        "navegacao": "navigation_notes",
    }
    return mapping.get(header, header)


def detect_sheet_header_and_rows(values: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    for index, row in enumerate(values):
        normalized_row = [normalize_header(cell) for cell in row]
        canonical: list[str] = []
        counts: dict[str, int] = {}
        for cell in normalized_row:
            base = canonical_sheet_field(cell)
            if not base:
                canonical.append("")
                continue
            counts[base] = counts.get(base, 0) + 1
            suffix = counts[base]
            canonical.append(base if suffix == 1 else f"{base}_{suffix}")
        if "patient_name" in canonical and ("regimen" in canonical or "diagnosis" in canonical):
            return canonical, values[index + 1 :]
    return [], []


def find_primary_workbook_file() -> Path | None:
    uploaded_override = DATA_DIR / UPLOADED_WORKBOOK_NAME
    if uploaded_override.exists():
        return uploaded_override

    if ONEDRIVE_CLOUDSTORAGE_DIR.exists():
        onedrive_matches = sorted(
            ONEDRIVE_CLOUDSTORAGE_DIR.glob(f"OneDrive*/*{PRIMARY_WORKBOOK_NAME}"),
            key=lambda path: path.stat().st_mtime,
            reverse=True,
        )
        if onedrive_matches:
            return onedrive_matches[0]

    direct = BASE_DIR / PRIMARY_WORKBOOK_NAME
    if direct.exists():
        return direct
    matches = sorted(BASE_DIR.glob("*.xlsx"))
    return matches[0] if matches else None


def save_uploaded_primary_workbook(uploaded_file) -> Path:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    target = DATA_DIR / UPLOADED_WORKBOOK_NAME
    target.write_bytes(uploaded_file.getbuffer())
    return target


def column_index_to_letter(index: int) -> str:
    result = ""
    current = index + 1
    while current:
        current, remainder = divmod(current - 1, 26)
        result = chr(65 + remainder) + result
    return result


def find_google_service_account_file() -> Path | None:
    candidates = sorted(BASE_DIR.glob("*.json"))
    return candidates[0] if candidates else None


def build_sheets_service():
    credential_file = find_google_service_account_file()
    if credential_file is None:
        raise FileNotFoundError("Arquivo JSON da conta de servico nao encontrado na pasta do projeto.")
    creds = Credentials.from_service_account_file(
        str(credential_file),
        scopes=GOOGLE_SHEETS_SCOPES,
    )
    return build("sheets", "v4", credentials=creds)


def get_spreadsheet_id() -> str:
    try:
        return st.secrets.get("GOOGLE_SHEETS_SPREADSHEET_ID", DEFAULT_SPREADSHEET_ID)
    except Exception:
        return DEFAULT_SPREADSHEET_ID


def get_google_sheet_titles(service, spreadsheet_id: str) -> list[str]:
    metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    titles = [sheet["properties"]["title"] for sheet in metadata.get("sheets", [])]
    return [title for title in titles if title.startswith(("Dr.", "Dra."))]


def get_workbook_doctor_sheet_titles(workbook_path: Path) -> list[str]:
    excel_file = pd.ExcelFile(workbook_path)
    return [title for title in excel_file.sheet_names if title.startswith(("Dr.", "Dra."))]


def extract_dates_from_text(value: object) -> list[date]:
    parsed = parse_date(value)
    if parsed:
        return [parsed]

    text = normalize_uploaded_text(value)
    if not text:
        return []

    matches = re.findall(r"\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b", text)
    dates: list[date] = []
    for match in matches:
        if match.count("/") == 1:
            match = f"{match}/{date.today().year}"
        parsed_match = parse_date(match)
        if parsed_match:
            dates.append(parsed_match)
    return dates


def extract_cycle_dates(row_data: dict[str, str]) -> list[date]:
    all_dates: list[date] = []
    for key, value in row_data.items():
        if key.startswith("next_infusion"):
            all_dates.extend(extract_dates_from_text(value))
    return sorted(set(all_dates))


def cycle_date_to_string(value: object) -> str | None:
    parsed = parse_date(value if isinstance(value, str) else value)
    return parsed.strftime(DATE_FMT) if parsed else None


def choose_next_relevant_date(dates: list[date]) -> date | None:
    if not dates:
        return None
    threshold = date.today() - timedelta(days=7)
    future_or_recent = [item for item in dates if item >= threshold]
    if future_or_recent:
        return future_or_recent[0]
    return dates[-1]


def choose_following_protocol_date(dates: list[date]) -> date | None:
    if not dates:
        return None
    reference = choose_next_relevant_date(dates)
    if reference is None:
        return None
    future_dates = [item for item in dates if item > reference]
    if future_dates:
        return future_dates[0]
    return reference


def evaluate_patient_alerts(row: pd.Series) -> tuple[str, str, int]:
    delta = days_until(row.get("next_chemo_date"))
    authorization_status = row.get("authorization_status")
    scheduling_status = row.get("scheduling_status")

    reasons: list[str] = []
    severity = 0

    if delta is not None and delta < 0:
        reasons.append("ciclo atrasado")
        severity = max(severity, 3)
    if authorization_status in {"not_sent", "pending", "denied"}:
        reasons.append("convênio ainda não liberado")
        severity = max(severity, 3 if authorization_status in {"not_sent", "denied"} else 2)
    if scheduling_status in {"not_booked", "awaiting_slot"}:
        reasons.append("risco de ficar fora da agenda")
        severity = max(severity, 3 if scheduling_status == "not_booked" else 2)
    if not reasons and delta is not None and delta <= 7:
        reasons.append("acompanhar confirmacao final")
        severity = max(severity, 1)
    if not reasons:
        return "Fluxo em dia", "flag-green", 0

    if severity >= 3:
        css_class = "flag-red"
    elif severity == 2:
        css_class = "flag-amber"
    else:
        css_class = "flag-blue"
    return " | ".join(reasons), css_class, severity


def evaluate_protocol_alert(row: pd.Series) -> tuple[str, str, int]:
    reference_date = row.get("protocol_next_cycle_date") or row.get("next_chemo_date")
    delta = days_until(reference_date)
    alert_days = int(row.get("next_cycle_alert_days") or 21)
    prescription_status = row.get("prescription_status")

    if delta is None:
        return "Sem próximo ciclo definido", "flag-blue", 0
    if prescription_status in {"prescribed", "sent_to_insurance"}:
        return "Próximo ciclo já encaminhado", "flag-green", 0
    if delta < 0:
        return f"Janela de {alert_days} dias já venceu para solicitar o próximo ciclo", "flag-protocol-strong", 3
    if delta <= alert_days:
        if delta == 0:
            message = f"Janela de {alert_days} dias vence hoje"
        else:
            message = f"Faltam {delta} dia(s) para o próximo ciclo. Solicitar nova prescrição."
        severity = 3 if prescription_status == "not_requested" else 2
        css_class = "flag-protocol-strong" if severity == 3 else "flag-protocol-soft"
        return message, css_class, severity
    return "Fora da janela de protocolo", "flag-green", 0


@st.cache_data(show_spinner=False)
def load_doctors() -> pd.DataFrame:
    conn = get_connection()
    df = pd.read_sql_query(
        """
        SELECT id, name, specialty, active
        FROM doctors
        ORDER BY name
        """,
        conn,
    )
    conn.close()
    return df


@st.cache_data(show_spinner=False)
def load_patients() -> pd.DataFrame:
    conn = get_connection()
    df = pd.read_sql_query(
        """
        SELECT
            p.id,
            p.name,
            d.name AS doctor_name,
            d.id AS doctor_id,
            p.diagnosis,
            p.regimen,
            p.cycle_interval_days,
            p.last_chemo_date,
            p.next_chemo_date,
            p.support_plan,
            p.notes,
            p.active,
            p.insurance_name,
            p.prescription_status,
            p.prescription_requested_date,
            p.authorization_status,
            p.authorization_submission_date,
            p.authorization_valid_until,
            p.scheduling_status,
            p.scheduled_cycle_date,
            p.next_cycle_alert_days,
            p.protocol_next_cycle_date,
            p.source_sheet_name,
            p.source_row_number
        FROM patients p
        JOIN doctors d ON d.id = p.doctor_id
        ORDER BY p.next_chemo_date IS NULL, p.next_chemo_date, p.name
        """,
        conn,
    )
    conn.close()
    return df


@st.cache_data(show_spinner=False)
def load_support_medications() -> pd.DataFrame:
    conn = get_connection()
    df = pd.read_sql_query(
        """
        SELECT
            s.id,
            p.name AS patient_name,
            d.name AS doctor_name,
            s.medication_name,
            s.purpose,
            s.frequency_label,
            s.next_due_date,
            s.status,
            s.notes
        FROM support_medications s
        JOIN patients p ON p.id = s.patient_id
        JOIN doctors d ON d.id = p.doctor_id
        ORDER BY s.next_due_date IS NULL, s.next_due_date, p.name
        """,
        conn,
    )
    conn.close()
    return df


@st.cache_data(show_spinner=False)
def load_chemo_sessions() -> pd.DataFrame:
    conn = get_connection()
    df = pd.read_sql_query(
        """
        SELECT
            c.id,
            p.id AS patient_id,
            p.name AS patient_name,
            d.name AS doctor_name,
            c.scheduled_date,
            c.cycle_label,
            c.status,
            c.notes
            ,
            c.prescription_status,
            c.authorization_status,
            c.scheduling_status
        FROM chemo_sessions c
        JOIN patients p ON p.id = c.patient_id
        JOIN doctors d ON d.id = p.doctor_id
        ORDER BY c.scheduled_date IS NULL, c.scheduled_date, p.name
        """,
        conn,
    )
    conn.close()
    return df


def refresh_data() -> None:
    load_doctors.clear()
    load_patients.clear()
    load_support_medications.clear()
    load_chemo_sessions.clear()


def insert_doctor(name: str, specialty: str) -> None:
    conn = get_connection()
    conn.execute(
        "INSERT INTO doctors (name, specialty) VALUES (?, ?)",
        (name.strip(), specialty.strip()),
    )
    conn.commit()
    conn.close()
    refresh_data()


def get_or_create_doctor_id(name: str, specialty: str = "") -> int:
    conn = get_connection()
    doctor_id = get_or_create_doctor_id_in_conn(conn, name, specialty)
    conn.commit()
    conn.close()
    return doctor_id


def get_or_create_doctor_id_in_conn(conn: sqlite3.Connection, name: str, specialty: str = "") -> int:
    existing = conn.execute(
        "SELECT id FROM doctors WHERE lower(name) = lower(?) LIMIT 1",
        (name.strip(),),
    ).fetchone()
    if existing:
        doctor_id = int(existing["id"])
    else:
        cursor = conn.execute(
            "INSERT INTO doctors (name, specialty) VALUES (?, ?)",
            (name.strip(), specialty.strip()),
        )
        doctor_id = int(cursor.lastrowid)
    return doctor_id


def insert_patient(payload: dict[str, object]) -> None:
    conn = get_connection()
    conn.execute(
        """
        INSERT INTO patients (
            doctor_id, name, diagnosis, regimen, cycle_interval_days,
            last_chemo_date, next_chemo_date, support_plan, notes,
            insurance_name, prescription_status, prescription_requested_date,
            authorization_status, authorization_submission_date,
            authorization_valid_until, scheduling_status, scheduled_cycle_date,
            next_cycle_alert_days
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            payload["doctor_id"],
            str(payload["name"]).strip(),
            str(payload["diagnosis"]).strip(),
            str(payload["regimen"]).strip(),
            payload["cycle_interval_days"],
            payload["last_chemo_date"],
            payload["next_chemo_date"],
            str(payload["support_plan"]).strip(),
            str(payload["notes"]).strip(),
            str(payload["insurance_name"]).strip(),
            payload["prescription_status"],
            payload["prescription_requested_date"],
            payload["authorization_status"],
            payload["authorization_submission_date"],
            payload["authorization_valid_until"],
            payload["scheduling_status"],
            payload["scheduled_cycle_date"],
            payload["next_cycle_alert_days"],
        ),
    )
    conn.commit()
    conn.close()
    refresh_data()


def insert_chemo_session_raw(conn: sqlite3.Connection, payload: dict[str, object]) -> None:
    conn.execute(
        """
        INSERT INTO chemo_sessions (
            patient_id, scheduled_date, cycle_label, status, notes,
            prescription_status, authorization_status, scheduling_status
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            payload["patient_id"],
            payload["scheduled_date"],
            str(payload["cycle_label"]).strip(),
            payload["status"],
            str(payload["notes"]).strip(),
            payload.get("prescription_status", "not_requested"),
            payload.get("authorization_status", "not_sent"),
            payload.get("scheduling_status", "not_booked"),
        ),
    )


def insert_support_med(payload: dict[str, object]) -> None:
    conn = get_connection()
    conn.execute(
        """
        INSERT INTO support_medications (
            patient_id, medication_name, purpose, frequency_label,
            next_due_date, status, notes
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            payload["patient_id"],
            str(payload["medication_name"]).strip(),
            str(payload["purpose"]).strip(),
            str(payload["frequency_label"]).strip(),
            payload["next_due_date"],
            payload["status"],
            str(payload["notes"]).strip(),
        ),
    )
    conn.commit()
    conn.close()
    refresh_data()


def insert_chemo_session(payload: dict[str, object]) -> None:
    conn = get_connection()
    insert_chemo_session_raw(conn, payload)
    conn.commit()
    conn.close()
    refresh_data()


def normalize_uploaded_date(value: object) -> str | None:
    parsed = parse_date(value)
    return parsed.strftime(DATE_FMT) if parsed else None


def normalize_uploaded_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def import_patients_dataframe(df: pd.DataFrame) -> tuple[int, int]:
    imported_patients = 0
    imported_sessions = 0
    conn = get_connection()

    for _, row in df.iterrows():
        doctor_name = normalize_uploaded_text(row.get("medico"))
        patient_name = normalize_uploaded_text(row.get("paciente"))
        if not doctor_name or not patient_name:
            continue

        doctor_id = get_or_create_doctor_id_in_conn(
            conn,
            doctor_name,
            normalize_uploaded_text(row.get("especialidade")),
        )

        patient_payload = {
            "doctor_id": doctor_id,
            "name": patient_name,
            "diagnosis": normalize_uploaded_text(row.get("diagnostico")),
            "regimen": normalize_uploaded_text(row.get("protocolo")),
            "cycle_interval_days": int(row.get("intervalo_dias") or 21),
            "last_chemo_date": normalize_uploaded_date(row.get("ultima_quimio")),
            "next_chemo_date": normalize_uploaded_date(row.get("proxima_quimio")),
            "support_plan": normalize_uploaded_text(row.get("suporte")),
            "notes": normalize_uploaded_text(row.get("observacoes")),
            "insurance_name": normalize_uploaded_text(row.get("convenio")),
            "prescription_status": normalize_uploaded_text(row.get("status_prescricao")) or "not_requested",
            "prescription_requested_date": normalize_uploaded_date(row.get("data_solicitacao_prescricao")),
            "authorization_status": normalize_uploaded_text(row.get("status_autorizacao")) or "not_sent",
            "authorization_submission_date": normalize_uploaded_date(row.get("data_envio_convenio")),
            "authorization_valid_until": normalize_uploaded_date(row.get("autorizacao_valida_ate")),
            "scheduling_status": normalize_uploaded_text(row.get("status_agendamento")) or "not_booked",
            "scheduled_cycle_date": normalize_uploaded_date(row.get("data_agendada")),
            "next_cycle_alert_days": int(row.get("alerta_novo_ciclo_dias") or 7),
        }

        cursor = conn.execute(
            """
            INSERT INTO patients (
                doctor_id, name, diagnosis, regimen, cycle_interval_days,
                last_chemo_date, next_chemo_date, support_plan, notes,
                insurance_name, prescription_status, prescription_requested_date,
                authorization_status, authorization_submission_date,
                authorization_valid_until, scheduling_status, scheduled_cycle_date,
                next_cycle_alert_days
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                patient_payload["doctor_id"],
                patient_payload["name"],
                patient_payload["diagnosis"],
                patient_payload["regimen"],
                patient_payload["cycle_interval_days"],
                patient_payload["last_chemo_date"],
                patient_payload["next_chemo_date"],
                patient_payload["support_plan"],
                patient_payload["notes"],
                patient_payload["insurance_name"],
                patient_payload["prescription_status"],
                patient_payload["prescription_requested_date"],
                patient_payload["authorization_status"],
                patient_payload["authorization_submission_date"],
                patient_payload["authorization_valid_until"],
                patient_payload["scheduling_status"],
                patient_payload["scheduled_cycle_date"],
                patient_payload["next_cycle_alert_days"],
            ),
        )
        imported_patients += 1
        patient_id = int(cursor.lastrowid)

        session_date = patient_payload["scheduled_cycle_date"] or patient_payload["next_chemo_date"]
        if session_date:
            insert_chemo_session_raw(
                conn,
                {
                    "patient_id": patient_id,
                    "scheduled_date": session_date,
                    "cycle_label": normalize_uploaded_text(row.get("ciclo")) or "Ciclo importado",
                    "status": "scheduled" if patient_payload["scheduling_status"] in {"scheduled", "confirmed"} else "attention",
                    "notes": normalize_uploaded_text(row.get("observacoes_agenda")) or "Sessao criada via importacao de planilha.",
                    "prescription_status": patient_payload["prescription_status"],
                    "authorization_status": patient_payload["authorization_status"],
                    "scheduling_status": patient_payload["scheduling_status"],
                },
            )
            imported_sessions += 1
            sync_patient_focus_fields(conn, patient_id)

    conn.commit()
    conn.close()
    refresh_data()
    return imported_patients, imported_sessions


def load_uploaded_dataframe(uploaded_file) -> pd.DataFrame:
    file_name = uploaded_file.name.lower()
    if file_name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


def map_sheet_row_to_patient_payload(
    row_data: dict[str, str],
    doctor_name: str,
    source_sheet_name: str,
    source_row_number: int,
) -> dict[str, object]:
    infusion_dates = extract_cycle_dates(row_data)
    next_infusion_date = choose_next_relevant_date(infusion_dates)
    protocol_cycle_date = choose_following_protocol_date(infusion_dates)
    next_infusion = next_infusion_date.strftime(DATE_FMT) if next_infusion_date else None
    protocol_next_cycle = protocol_cycle_date.strftime(DATE_FMT) if protocol_cycle_date else None

    notes_parts = [
        normalize_uploaded_text(row_data.get("notes")),
        normalize_uploaded_text(row_data.get("navigation_notes")),
    ]
    notes = " | ".join(part for part in notes_parts if part)
    solicitate_pm = normalize_uploaded_text(row_data.get("prescription_prompt"))
    prescription_status = "requested" if solicitate_pm else "not_requested"
    scheduling_status = "scheduled" if next_infusion else "not_booked"

    return {
        "name": normalize_uploaded_text(row_data.get("patient_name")),
        "diagnosis": normalize_uploaded_text(row_data.get("diagnosis")),
        "regimen": normalize_uploaded_text(row_data.get("regimen")),
        "cycle_interval_days": 30,
        "last_chemo_date": None,
        "next_chemo_date": next_infusion,
        "support_plan": "",
        "notes": notes,
        "insurance_name": "",
        "prescription_status": prescription_status,
        "prescription_requested_date": None,
        "authorization_status": "not_sent",
        "authorization_submission_date": None,
        "authorization_valid_until": None,
        "scheduling_status": scheduling_status,
        "scheduled_cycle_date": next_infusion,
        "next_cycle_alert_days": 21,
        "protocol_next_cycle_date": protocol_next_cycle,
        "doctor_name": doctor_name,
        "source_sheet_name": source_sheet_name,
        "source_row_number": source_row_number,
        "medical_record": normalize_uploaded_text(row_data.get("medical_record")),
        "birth_date": normalize_uploaded_date(row_data.get("birth_date")),
        "next_consultation_date": normalize_uploaded_date(row_data.get("next_consultation_date")),
        "attention_date": normalize_uploaded_date(row_data.get("attention_date")),
    }


def sync_google_sheets_to_db() -> tuple[int, int]:
    workbook_path = find_primary_workbook_file()
    if workbook_path is None:
        raise FileNotFoundError("Planilha principal .xlsx não encontrada na pasta do projeto.")

    doctor_sheets = get_workbook_doctor_sheet_titles(workbook_path)
    conn = get_connection()
    existing_patient_state = {
        (row["source_sheet_name"], int(row["source_row_number"])): {
            "insurance_name": row["insurance_name"],
            "notes": row["notes"],
            "prescription_status": row["prescription_status"],
            "authorization_status": row["authorization_status"],
            "scheduling_status": row["scheduling_status"],
            "next_cycle_alert_days": row["next_cycle_alert_days"],
            "next_chemo_date": row["next_chemo_date"],
            "scheduled_cycle_date": row["scheduled_cycle_date"],
            "protocol_next_cycle_date": row["protocol_next_cycle_date"],
        }
        for row in conn.execute(
            """
            SELECT
                source_sheet_name,
                source_row_number,
                insurance_name,
                notes,
                prescription_status,
                authorization_status,
                scheduling_status,
                next_cycle_alert_days,
                next_chemo_date,
                scheduled_cycle_date,
                protocol_next_cycle_date
            FROM patients
            WHERE source_sheet_name IS NOT NULL AND source_row_number IS NOT NULL
            """
        ).fetchall()
    }
    existing_session_state = {
        (row["source_sheet_name"], int(row["source_row_number"]), row["scheduled_date"], row["cycle_label"] or ""): {
            "status": row["status"],
            "notes": row["notes"],
            "prescription_status": row["prescription_status"] or row["patient_prescription_status"] or "not_requested",
            "authorization_status": row["authorization_status"] or row["patient_authorization_status"] or "not_sent",
            "scheduling_status": row["scheduling_status"] or row["patient_scheduling_status"] or "not_booked",
        }
        for row in conn.execute(
            """
            SELECT
                p.source_sheet_name,
                p.source_row_number,
                c.scheduled_date,
                c.cycle_label,
                c.status,
                c.notes,
                c.prescription_status,
                c.authorization_status,
                c.scheduling_status,
                p.prescription_status AS patient_prescription_status,
                p.authorization_status AS patient_authorization_status,
                p.scheduling_status AS patient_scheduling_status
            FROM chemo_sessions c
            JOIN patients p ON p.id = c.patient_id
            WHERE p.source_sheet_name IS NOT NULL AND p.source_row_number IS NOT NULL
            """
        ).fetchall()
    }

    conn.execute("DELETE FROM support_medications")
    conn.execute("DELETE FROM chemo_sessions")
    conn.execute("DELETE FROM patients")
    conn.execute("DELETE FROM doctors")

    imported = 0
    updated = 0

    for sheet_name in doctor_sheets:
        dataframe = pd.read_excel(workbook_path, sheet_name=sheet_name, header=None, dtype=object)
        values = dataframe.fillna("").values.tolist()
        headers, data_rows = detect_sheet_header_and_rows(values)
        if not headers or "patient_name" not in headers:
            continue

        doctor_id = get_or_create_doctor_id_in_conn(conn, sheet_name)
        header_row_number = next(
            (
                idx
                for idx, row in enumerate(values, start=1)
                if detect_sheet_header_and_rows([row])[0] == headers
            ),
            1,
        )

        for row_offset, row in enumerate(data_rows, start=1):
            row_number = header_row_number + row_offset
            padded = row + [""] * (len(headers) - len(row))
            row_data = dict(zip(headers, padded))
            payload = map_sheet_row_to_patient_payload(row_data, sheet_name, sheet_name, row_number)
            existing_state = existing_patient_state.get((sheet_name, row_number))
            if existing_state:
                for field_name in [
                    "insurance_name",
                    "notes",
                    "prescription_status",
                    "authorization_status",
                    "scheduling_status",
                    "next_chemo_date",
                    "scheduled_cycle_date",
                    "protocol_next_cycle_date",
                ]:
                    if existing_state.get(field_name) not in {None, ""}:
                        payload[field_name] = existing_state[field_name]
                existing_alert_days = existing_state.get("next_cycle_alert_days")
                if existing_alert_days not in {None, "", 0, 7}:
                    payload["next_cycle_alert_days"] = existing_alert_days
            if not payload["name"]:
                continue

            cursor = conn.execute(
                """
                INSERT INTO patients (
                    doctor_id, name, diagnosis, regimen, cycle_interval_days,
                    last_chemo_date, next_chemo_date, support_plan, notes,
                    insurance_name, prescription_status, prescription_requested_date,
                    authorization_status, authorization_submission_date,
                    authorization_valid_until, scheduling_status, scheduled_cycle_date,
                    next_cycle_alert_days, protocol_next_cycle_date, source_sheet_name, source_row_number
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    doctor_id,
                    payload["name"],
                    payload["diagnosis"],
                    payload["regimen"],
                    payload["cycle_interval_days"],
                    payload["last_chemo_date"],
                    payload["next_chemo_date"],
                    payload["support_plan"],
                    payload["notes"],
                    payload["insurance_name"],
                    payload["prescription_status"],
                    payload["prescription_requested_date"],
                    payload["authorization_status"],
                    payload["authorization_submission_date"],
                    payload["authorization_valid_until"],
                    payload["scheduling_status"],
                    payload["scheduled_cycle_date"],
                    payload["next_cycle_alert_days"],
                    payload["protocol_next_cycle_date"],
                    payload["source_sheet_name"],
                    payload["source_row_number"],
                ),
            )
            imported += 1
            patient_id = int(cursor.lastrowid)

            cycle_dates = extract_cycle_dates(row_data)
            if not cycle_dates and payload["scheduled_cycle_date"]:
                fallback_date = parse_date(payload["scheduled_cycle_date"])
                cycle_dates = [fallback_date] if fallback_date else []
            for cycle_date in cycle_dates:
                cycle_date_str = cycle_date.strftime(DATE_FMT)
                session_key = (sheet_name, row_number, cycle_date_str, payload["regimen"] or "Infusao")
                session_state = existing_session_state.get(session_key, {})
                insert_chemo_session_raw(
                    conn,
                    {
                        "patient_id": patient_id,
                        "scheduled_date": cycle_date_str,
                        "cycle_label": payload["regimen"] or "Infusao",
                        "status": session_state.get("status", "scheduled"),
                        "notes": session_state.get("notes", f"Sessao sincronizada de {workbook_path.name}."),
                        "prescription_status": session_state.get("prescription_status", payload["prescription_status"]),
                        "authorization_status": session_state.get("authorization_status", payload["authorization_status"]),
                        "scheduling_status": session_state.get("scheduling_status", payload["scheduling_status"]),
                    },
                )
            sync_patient_focus_fields(conn, patient_id)

    conn.commit()
    conn.close()
    set_app_state("last_google_sync_at", datetime.now(APP_TIMEZONE).isoformat(timespec="seconds"))
    refresh_data()
    return imported, updated


def update_google_sheet_patient_row(patient_row: pd.Series, updates: dict[str, object]) -> None:
    source_sheet_name = patient_row.get("source_sheet_name")
    source_row_number = patient_row.get("source_row_number")
    if not source_sheet_name or pd.isna(source_row_number):
        return

    workbook_path = find_primary_workbook_file()
    if workbook_path is None:
        raise FileNotFoundError("Planilha principal .xlsx não encontrada na pasta do projeto.")

    row_number = int(source_row_number)
    from openpyxl import load_workbook

    workbook = load_workbook(workbook_path)
    worksheet = workbook[source_sheet_name]
    field_map = {
        "diagnosis": 5,
        "regimen": 7,
        "next_chemo_date": 8,
        "notes": 13,
    }
    for field_name, column_index in field_map.items():
        if field_name not in updates:
            continue
        value = updates[field_name]
        if field_name == "next_chemo_date" and value:
            parsed = parse_date(value)
            value = parsed.strftime("%d/%m/%Y") if parsed else ""
        worksheet[f"{column_index_to_letter(column_index - 1)}{row_number}"] = value or ""
    workbook.save(workbook_path)


def render_metric(label: str, value: str, copy: str, css_class: str) -> None:
    st.markdown(
        f"""
        <div class="metric-card {css_class}">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-copy">{copy}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_calendar_events(filtered_patients: pd.DataFrame, filtered_sessions: pd.DataFrame) -> pd.DataFrame:
    if filtered_patients.empty or filtered_sessions.empty:
        return pd.DataFrame(
            columns=[
                "patient_id",
                "patient_name",
                "doctor_name",
                "cycle_label",
                "scheduled_date",
                "insurance_name",
                "prescription_status",
                "authorization_status",
                "scheduling_status",
            ]
        )

    events = filtered_sessions.copy()
    events["scheduled_date"] = pd.to_datetime(events["scheduled_date"], errors="coerce").dt.date
    events = events.dropna(subset=["scheduled_date"])
    if events.empty:
        return pd.DataFrame()

    patient_meta = filtered_patients[
        [
            "id",
            "insurance_name",
            "diagnosis",
            "regimen",
        ]
    ].rename(columns={"id": "patient_id"})
    merged = events.merge(patient_meta, on="patient_id", how="left")
    return merged.sort_values(["scheduled_date", "doctor_name", "patient_name"])


def default_cycle_date_for_patient(patient_sessions: pd.DataFrame, patient_row: pd.Series) -> str | None:
    if not patient_sessions.empty:
        future_sessions = patient_sessions[patient_sessions["session_date"] >= date.today()]
        if not future_sessions.empty:
            return future_sessions.iloc[0]["session_date"].strftime(DATE_FMT)
        return patient_sessions.iloc[-1]["session_date"].strftime(DATE_FMT)
    fallback_date = parse_date(patient_row.get("scheduled_cycle_date")) or parse_date(patient_row.get("next_chemo_date"))
    return fallback_date.strftime(DATE_FMT) if fallback_date else None


def abbreviate_patient_name(name: str) -> str:
    parts = [part for part in (name or "").split() if part]
    if len(parts) <= 2:
        return name
    middle = [f"{part[0]}." for part in parts[1:-1]]
    return " ".join([parts[0], *middle, parts[-1]])


def render_calendar_patient_detail_page(filtered_patients: pd.DataFrame, filtered_sessions: pd.DataFrame) -> None:
    selected_patient_id = st.session_state.get("selected_calendar_patient_id")
    if not selected_patient_id:
        st.info("Nenhum paciente foi selecionado.")
        return

    patient_matches = filtered_patients[filtered_patients["id"] == selected_patient_id]
    if patient_matches.empty:
        top_left, _ = st.columns([1, 5])
        with top_left:
            if st.button("Voltar para o calendário", use_container_width=True):
                close_patient_detail()
                st.rerun()
        st.warning("Não encontrei o paciente selecionado neste momento. Tente voltar ao calendário e abrir novamente.")
        return
    patient_row = patient_matches.iloc[0]

    patient_sessions = filtered_sessions[filtered_sessions["patient_id"] == selected_patient_id].copy()
    patient_sessions["session_date"] = pd.to_datetime(patient_sessions["scheduled_date"], errors="coerce").dt.date
    patient_sessions = patient_sessions.dropna(subset=["session_date"]).sort_values("session_date")

    selected_cycle_value = st.session_state.get("selected_calendar_cycle_date") or default_cycle_date_for_patient(patient_sessions, patient_row)
    date_options = [session_date.strftime(DATE_FMT) for session_date in patient_sessions["session_date"].tolist()]
    if selected_cycle_value and selected_cycle_value not in date_options:
        date_options.append(selected_cycle_value)
    date_options = sorted(set(date_options))

    top_left, top_right = st.columns([1, 5])
    with top_left:
        if st.button("Voltar para o calendário", use_container_width=True):
            close_patient_detail()
            st.rerun()
    with top_right:
        st.markdown("")

    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown(f'<div class="section-title">Resumo do paciente: {patient_row["name"]}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="subtle">{patient_row["doctor_name"]} | {patient_row["regimen"] or "Sem protocolo informado"} | {patient_row["diagnosis"] or "Sem diagnóstico informado"}</div>',
        unsafe_allow_html=True,
    )

    if date_options:
        default_index = date_options.index(selected_cycle_value) if selected_cycle_value in date_options else 0
        selected_cycle_date = st.selectbox(
            "Dia do ciclo / execução",
            date_options,
            index=default_index,
            format_func=lambda value: datetime.strptime(value, DATE_FMT).strftime("%d/%m/%Y"),
            key=f"calendar_patient_cycle_view_{selected_patient_id}",
        )
        st.session_state["selected_calendar_cycle_date"] = selected_cycle_date
        selected_cycle_value = selected_cycle_date
    else:
        selected_cycle_date = None

    selected_session = None
    if selected_cycle_value:
        selected_session_matches = patient_sessions[
            patient_sessions["session_date"] == parse_date(selected_cycle_value)
        ]
        if not selected_session_matches.empty:
            selected_session = selected_session_matches.iloc[0]
    if selected_session is None and not patient_sessions.empty:
        selected_session = patient_sessions.iloc[0]
        selected_cycle_value = selected_session["session_date"].strftime(DATE_FMT)

    info1, info2, info3, info4 = st.columns(4)
    with info1:
        st.write(f"**Convênio:** {patient_row['insurance_name'] or 'Não informado'}")
    with info2:
        st.write(f"**Ciclo em foco:** {format_date(selected_cycle_value)}")
    with info3:
        session_scheduling = selected_session["scheduling_status"] if selected_session is not None else "not_booked"
        st.write(f"**Agenda atual:** {format_status(session_scheduling, SCHEDULING_LABELS)}")
    with info4:
        st.write(f"**Próximo protocolo:** {format_date(patient_row.get('protocol_next_cycle_date'))}")

    if not patient_sessions.empty:
        session_display = patient_sessions[["session_date", "cycle_label", "status"]].copy()
        session_display["Data prevista"] = session_display["session_date"].apply(lambda value: value.strftime("%d/%m/%Y"))
        session_display["Status"] = session_display["status"].apply(lambda value: format_status(value, STATUS_LABELS))
        st.dataframe(
            session_display[["Data prevista", "cycle_label", "Status"]].rename(columns={"cycle_label": "Ciclo / protocolo"}),
            use_container_width=True,
            hide_index=True,
        )

    prescribed_flag = selected_session is not None and selected_session["prescription_status"] in {"prescribed", "sent_to_insurance"}
    authorization_flag = selected_session is not None and selected_session["authorization_status"] in {"pending", "authorized"}
    scheduling_flag = selected_session is not None and selected_session["scheduling_status"] in {"scheduled", "confirmed"}
    session_notes = selected_session["notes"] if selected_session is not None else ""

    with st.form(f"calendar_patient_detail_form_{selected_patient_id}"):
        notes = st.text_area("Observacoes do ciclo", value=session_notes or "")
        prescribed_checked = st.checkbox("Prescricao gerada", value=prescribed_flag)
        authorization_checked = st.checkbox("Encaminhado para autorização", value=authorization_flag)
        scheduled_checked = st.checkbox("Agendamento realizado", value=scheduling_flag)
        submitted = st.form_submit_button("Salvar acompanhamento", use_container_width=True)
        if submitted:
            if scheduled_checked:
                authorization_checked = True
                prescribed_checked = True
            elif authorization_checked:
                prescribed_checked = True

            session_matches = patient_sessions[patient_sessions["session_date"] == parse_date(selected_cycle_date)]
            if session_matches.empty:
                st.error("Não encontrei o ciclo selecionado para salvar as flags.")
            else:
                target_session = session_matches.iloc[0]
                update_chemo_session_record(
                    int(target_session["id"]),
                    {
                        "notes": notes.strip(),
                        "prescription_status": "prescribed" if prescribed_checked else "not_requested",
                        "authorization_status": "pending" if authorization_checked else "not_sent",
                        "scheduling_status": "scheduled" if scheduled_checked else "not_booked",
                    },
                )
                st.session_state["selected_calendar_cycle_date"] = selected_cycle_date
                st.success("Resumo do ciclo atualizado no app.")
                st.rerun()

    if date_options:
        st.session_state["selected_calendar_cycle_date"] = selected_cycle_value

    st.markdown("---")
    st.markdown("**Planejamento do próximo ciclo do protocolo**")
    protocol_default_date = parse_date(patient_row.get("protocol_next_cycle_date"))
    if protocol_default_date is None:
        base_date = parse_date(selected_cycle_value) or parse_date(patient_row.get("next_chemo_date")) or date.today()
        protocol_default_date = base_date + timedelta(days=int(patient_row.get("cycle_interval_days") or 30))

    with st.form(f"calendar_patient_protocol_form_{selected_patient_id}"):
        protocol_next_cycle_date = st.date_input(
            "Próxima data futura do protocolo",
            value=protocol_default_date,
        )
        protocol_alert_days = st.number_input(
            "Alertar com quantos dias de antecedência",
            min_value=1,
            max_value=90,
            value=int(patient_row.get("next_cycle_alert_days") or 21),
        )
        create_future_session = st.checkbox(
            "Criar também essa data na agenda futura do paciente",
            value=False,
        )
        protocol_submitted = st.form_submit_button("Salvar regra do protocolo", use_container_width=True)
        if protocol_submitted:
            update_patient_record(
                selected_patient_id,
                {
                    "protocol_next_cycle_date": protocol_next_cycle_date.strftime(DATE_FMT) if protocol_next_cycle_date else None,
                    "next_cycle_alert_days": int(protocol_alert_days),
                },
            )

            if create_future_session and protocol_next_cycle_date:
                existing_future_match = patient_sessions[
                    patient_sessions["session_date"] == protocol_next_cycle_date
                ]
                if existing_future_match.empty:
                    insert_chemo_session(
                        {
                            "patient_id": selected_patient_id,
                            "scheduled_date": protocol_next_cycle_date.strftime(DATE_FMT),
                            "cycle_label": patient_row["regimen"] or "Próximo ciclo",
                            "status": "scheduled",
                            "notes": "Sessão futura criada pelo planejamento do protocolo.",
                            "prescription_status": "not_requested",
                            "authorization_status": "not_sent",
                            "scheduling_status": "not_booked",
                        }
                    )
            st.success("Regra do protocolo atualizada.")
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def render_patient_link_card(patient_id: int, patient_name: str, doctor_name: str, cycle_date: date, *, show_doctor_name: bool = True) -> None:
    st.markdown('<div class="calendar-link-card">', unsafe_allow_html=True)
    if st.button(
        abbreviate_patient_name(patient_name),
        key=f"calendar_link_{patient_id}_{cycle_date.strftime(DATE_FMT)}",
        type="tertiary",
        use_container_width=False,
    ):
        open_patient_detail(patient_id, cycle_date.strftime(DATE_FMT))
        st.rerun()
    if show_doctor_name:
        st.markdown(f'<div class="calendar-link-meta">{doctor_name}</div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


def render_calendar_panel(
    filtered_patients: pd.DataFrame,
    filtered_sessions: pd.DataFrame,
    *,
    show_panel_wrapper: bool = True,
    show_intro: bool = True,
    show_metrics: bool = True,
    show_month_list: bool = True,
    month_select_key: str = "calendar_month",
) -> None:
    if show_panel_wrapper:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
    if show_intro:
        st.markdown('<div class="section-title">Calendario de infusoes</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="subtle">Visao inspirada em agenda mensal, projetando os proximos ciclos ao longo do ano a partir da programacao atual de cada paciente.</div>',
            unsafe_allow_html=True,
        )
    if filtered_patients.empty:
        st.info("Nenhum paciente encontrado para montar o calendário.")
        if show_panel_wrapper:
            st.markdown("</div>", unsafe_allow_html=True)
        return

    calendar_events = build_calendar_events(filtered_patients, filtered_sessions)
    if calendar_events.empty:
        st.info("Os pacientes filtrados ainda não possuem agendamentos para montar o calendário.")
        if show_panel_wrapper:
            st.markdown("</div>", unsafe_allow_html=True)
        return

    month_options = sorted(
        {
            scheduled.strftime("%Y-%m")
            for scheduled in calendar_events["scheduled_date"].tolist()
        }
    )
    default_month = date.today().strftime("%Y-%m")
    default_index = month_options.index(default_month) if default_month in month_options else 0
    selected_month = st.selectbox(
        "Mês do calendário",
        month_options,
        index=default_index,
        format_func=format_month_label_pt,
        key=month_select_key,
    )

    month_start = datetime.strptime(f"{selected_month}-01", DATE_FMT).date()
    next_month = (month_start.replace(day=28) + timedelta(days=4)).replace(day=1)
    month_end = next_month - timedelta(days=1)

    month_projection = calendar_events[
        (calendar_events["scheduled_date"] >= month_start) & (calendar_events["scheduled_date"] <= month_end)
    ].copy()

    if show_metrics:
        col1, col2, col3 = st.columns(3)
        with col1:
            render_metric("Infusoes no mes", str(len(month_projection)), "Total agendado no mes filtrado.", "metric-a")
        with col2:
            busy_days = int(month_projection["scheduled_date"].nunique()) if not month_projection.empty else 0
            render_metric("Dias com agenda", str(busy_days), "Dias do mes com pelo menos uma infusao.", "metric-b")
        with col3:
            busiest = 0 if month_projection.empty else int(month_projection.groupby("scheduled_date").size().max())
            render_metric("Pico no mesmo dia", str(busiest), "Maior concentracao de infusoes em um unico dia.", "metric-d")

    start_offset = month_start.weekday()
    total_slots = ((start_offset + month_end.day + 6) // 7) * 7
    day_map: dict[date, pd.DataFrame] = {
        projected_date: group
        for projected_date, group in month_projection.groupby("scheduled_date")
    }

    header_cols = st.columns(7)
    for idx, day_name in enumerate(["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]):
        with header_cols[idx]:
            st.markdown(
                f"""
                <div style="
                    color:#1f2937;
                    font-size:0.85rem;
                    font-weight:800;
                    text-transform:uppercase;
                    letter-spacing:0.05em;
                    padding:4px 2px 8px 2px;
                ">
                    {day_name}
                </div>
                """,
                unsafe_allow_html=True,
            )

    for week_start in range(0, total_slots, 7):
        week_cols = st.columns(7)
        for day_offset in range(7):
            slot = week_start + day_offset
            cell_date = month_start - timedelta(days=start_offset) + timedelta(days=slot)
            in_month = month_start <= cell_date <= month_end
            group = day_map.get(cell_date)
            count = 0 if group is None else len(group)
            bg_color = "rgba(255, 255, 255, 0.92)" if in_month else "rgba(245, 248, 250, 0.92)"
            date_color = "#123847" if in_month else "#8aa0a8"
            with week_cols[day_offset]:
                st.markdown(
                    f"""
                    <div style="
                        min-height: 168px;
                        background: {bg_color};
                        border: 1px solid rgba(15, 61, 76, 0.08);
                        border-radius: 18px;
                        padding: 10px;
                        box-shadow: 0 10px 24px rgba(26, 55, 77, 0.05);
                        margin-bottom: 10px;
                    ">
                        <div style="font-size:0.92rem; font-weight:800; color:{date_color}; margin-bottom:8px;">
                            {cell_date.day}
                        </div>
                    """,
                    unsafe_allow_html=True,
                )
                if count:
                    st.markdown(
                        f"""
                        <div style="
                            display:inline-block;
                            font-size:0.76rem;
                            font-weight:800;
                            color:#9a3412;
                            background:#ffedd5;
                            border-radius:999px;
                            padding:4px 8px;
                            margin-bottom:8px;
                        ">
                            {count} infusao(oes)
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    for doctor_name, doctor_group in group.sort_values(["doctor_name", "patient_name"]).groupby("doctor_name"):
                        st.markdown(f'<div class="calendar-doctor-group">{doctor_name}</div>', unsafe_allow_html=True)
                        for _, event in doctor_group.iterrows():
                            render_patient_link_card(
                                int(event["patient_id"]),
                                str(event["patient_name"]),
                                str(event["doctor_name"]),
                                event["scheduled_date"],
                                show_doctor_name=False,
                            )
                st.markdown("</div>", unsafe_allow_html=True)

    if show_month_list:
        st.markdown("")
        st.markdown("**Lista do mes**")
        if month_projection.empty:
            st.info("Nenhuma infusao projetada para este mes.")
        else:
            list_df = month_projection.copy()
            list_df["Data"] = list_df["scheduled_date"].apply(lambda value: value.strftime("%d/%m/%Y"))
            list_df["Prescricao"] = list_df["prescription_status"].apply(lambda value: format_status(value, PRESCRIPTION_LABELS))
            list_df["Autorizacao"] = list_df["authorization_status"].apply(lambda value: format_status(value, AUTHORIZATION_LABELS))
            list_df["Agenda"] = list_df["scheduling_status"].apply(lambda value: format_status(value, SCHEDULING_LABELS))
            st.dataframe(
                list_df[
                    [
                        "Data",
                        "patient_name",
                        "doctor_name",
                        "cycle_label",
                        "insurance_name",
                        "Prescricao",
                        "Autorizacao",
                        "Agenda",
                    ]
                ].rename(
                    columns={
                        "patient_name": "Paciente",
                        "doctor_name": "Medico",
                        "cycle_label": "Ciclo / protocolo",
                        "insurance_name": "Convenio",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    if show_panel_wrapper:
        st.markdown("</div>", unsafe_allow_html=True)


def render_simple_dashboard(
    filtered_patients: pd.DataFrame,
    filtered_support: pd.DataFrame,
    filtered_sessions: pd.DataFrame,
) -> None:
    today = date.today()
    next_week = today + timedelta(days=7)

    patients_df = filtered_patients.copy()
    support_df = filtered_support.copy()
    sessions_df = filtered_sessions.copy()

    if not support_df.empty:
        support_df["next_due_dt"] = pd.to_datetime(support_df["next_due_date"], errors="coerce")
    if not sessions_df.empty:
        sessions_df["scheduled_dt"] = pd.to_datetime(sessions_df["scheduled_date"], errors="coerce")

    chemo_today = (
        sessions_df["scheduled_dt"].dt.date.eq(today).sum()
        if not sessions_df.empty
        else 0
    )
    chemo_week = (
        sessions_df["scheduled_dt"].dt.date.between(today, next_week).sum()
        if not sessions_df.empty
        else 0
    )
    overdue_chemo = (
        sessions_df["scheduled_dt"].dt.date.lt(today).sum()
        if not sessions_df.empty
        else 0
    )
    overdue_support = (
        support_df["next_due_dt"].dt.date.lt(today).sum()
        if not support_df.empty
        else 0
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_metric("Pacientes ativos", str(len(patients_df)), "Carteira visivel no filtro atual.", "metric-a")
    with col2:
        render_metric("Quimio hoje", str(int(chemo_today)), "Pacientes com infusao prevista para hoje.", "metric-d")
    with col3:
        render_metric("Proximos 7 dias", str(int(chemo_week)), "Janela curta para contato e confirmacao.", "metric-b")
    with col4:
        render_metric("Pendencias", str(int(overdue_chemo + overdue_support)), "Quimios e suportes em atraso.", "metric-c")
    st.markdown("")
    render_calendar_panel(
        filtered_patients,
        filtered_sessions,
        show_panel_wrapper=True,
        show_intro=False,
        show_metrics=False,
        show_month_list=False,
        month_select_key="simple_dashboard_month",
    )


def build_operational_table(filtered_patients: pd.DataFrame) -> pd.DataFrame:
    display = filtered_patients.copy()
    if display.empty:
        return display
    display["Dias para ciclo"] = display["next_chemo_date"].apply(days_until)
    display["Próxima quimio"] = display["next_chemo_date"].apply(format_date)
    display["Próximo protocolo"] = display["protocol_next_cycle_date"].apply(format_date)
    display["Data agenda"] = display["scheduled_cycle_date"].apply(format_date)
    display["Status prescrição"] = display["prescription_status"].apply(lambda value: format_status(value, PRESCRIPTION_LABELS))
    display["Status autorização"] = display["authorization_status"].apply(lambda value: format_status(value, AUTHORIZATION_LABELS))
    display["Status agenda"] = display["scheduling_status"].apply(lambda value: format_status(value, SCHEDULING_LABELS))
    display["Alerta operacional"] = display.apply(lambda row: evaluate_patient_alerts(row)[0], axis=1)
    display["Severidade operacional"] = display.apply(lambda row: evaluate_patient_alerts(row)[2], axis=1)
    display["Alerta de protocolo"] = display.apply(lambda row: evaluate_protocol_alert(row)[0], axis=1)
    display["Severidade protocolo"] = display.apply(lambda row: evaluate_protocol_alert(row)[2], axis=1)
    return display


def update_patient_record(patient_id: int, updates: dict[str, object]) -> None:
    allowed_fields = [
        "diagnosis",
        "regimen",
        "next_chemo_date",
        "protocol_next_cycle_date",
        "notes",
        "insurance_name",
        "prescription_status",
        "authorization_status",
        "scheduling_status",
        "scheduled_cycle_date",
        "next_cycle_alert_days",
    ]
    assignments = []
    values: list[object] = []
    for field_name in allowed_fields:
        if field_name in updates:
            assignments.append(f"{field_name} = ?")
            values.append(updates[field_name])
    if not assignments:
        return

    conn = get_connection()
    conn.execute(
        f"UPDATE patients SET {', '.join(assignments)} WHERE id = ?",
        (*values, patient_id),
    )
    conn.commit()
    conn.close()
    refresh_data()


def get_focus_session_row(conn: sqlite3.Connection, patient_id: int) -> sqlite3.Row | None:
    today_value = date.today().strftime(DATE_FMT)
    row = conn.execute(
        """
        SELECT *
        FROM chemo_sessions
        WHERE patient_id = ? AND scheduled_date >= ?
        ORDER BY scheduled_date ASC, id ASC
        LIMIT 1
        """,
        (patient_id, today_value),
    ).fetchone()
    if row:
        return row
    return conn.execute(
        """
        SELECT *
        FROM chemo_sessions
        WHERE patient_id = ?
        ORDER BY scheduled_date DESC, id DESC
        LIMIT 1
        """,
        (patient_id,),
    ).fetchone()


def sync_patient_focus_fields(conn: sqlite3.Connection, patient_id: int) -> None:
    focus_session = get_focus_session_row(conn, patient_id)
    if not focus_session:
        return
    conn.execute(
        """
        UPDATE patients
        SET
            next_chemo_date = ?,
            scheduled_cycle_date = ?,
            prescription_status = ?,
            authorization_status = ?,
            scheduling_status = ?
        WHERE id = ?
        """,
        (
            focus_session["scheduled_date"],
            focus_session["scheduled_date"],
            focus_session["prescription_status"],
            focus_session["authorization_status"],
            focus_session["scheduling_status"],
            patient_id,
        ),
    )


def update_chemo_session_record(session_id: int, updates: dict[str, object]) -> None:
    allowed_fields = [
        "scheduled_date",
        "cycle_label",
        "status",
        "notes",
        "prescription_status",
        "authorization_status",
        "scheduling_status",
    ]
    assignments = []
    values: list[object] = []
    for field_name in allowed_fields:
        if field_name in updates:
            assignments.append(f"{field_name} = ?")
            values.append(updates[field_name])
    if not assignments:
        return

    conn = get_connection()
    conn.execute(
        f"UPDATE chemo_sessions SET {', '.join(assignments)} WHERE id = ?",
        (*values, session_id),
    )
    patient_row = conn.execute(
        "SELECT patient_id FROM chemo_sessions WHERE id = ? LIMIT 1",
        (session_id,),
    ).fetchone()
    if patient_row:
        sync_patient_focus_fields(conn, int(patient_row["patient_id"]))
    conn.commit()
    conn.close()
    refresh_data()


def render_dashboard(
    filtered_patients: pd.DataFrame,
    filtered_support: pd.DataFrame,
    filtered_sessions: pd.DataFrame,
) -> None:
    today = date.today()
    next_week = today + timedelta(days=7)

    patients_df = build_operational_table(filtered_patients)
    support_df = filtered_support.copy()
    sessions_df = filtered_sessions.copy()

    if not support_df.empty:
        support_df["next_due_dt"] = pd.to_datetime(support_df["next_due_date"], errors="coerce")
    if not sessions_df.empty:
        sessions_df["scheduled_dt"] = pd.to_datetime(sessions_df["scheduled_date"], errors="coerce")

    chemo_week = (
        patients_df["Dias para ciclo"].between(0, 7).sum()
        if not patients_df.empty
        else 0
    )
    protocol_window = (
        (patients_df["Severidade protocolo"] > 0).sum()
        if not patients_df.empty
        else 0
    )
    need_authorization = (
        patients_df["authorization_status"].isin(["not_sent", "pending", "denied"]).sum()
        if not patients_df.empty
        else 0
    )
    agenda_risk = (
        patients_df["scheduling_status"].isin(["not_booked", "awaiting_slot"]).sum()
        if not patients_df.empty
        else 0
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_metric("Ciclos nos próximos 7 dias", str(int(chemo_week)), "Janela curta para validar todo o fluxo.", "metric-a")
    with col2:
        render_metric("Janela de protocolo", str(int(protocol_window)), "Pacientes que atingiram a janela de 21 dias para solicitar o próximo ciclo.", "metric-protocol")
    with col3:
        render_metric("Pendentes no convênio", str(int(need_authorization)), "Inclui não enviados, em análise ou negados.", "metric-c")
    with col4:
        render_metric("Risco de ficar fora da agenda", str(int(agenda_risk)), "Pacientes sem vaga confirmada para infusão.", "metric-d")

    left, right = st.columns([1.25, 1])
    with left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<span class="section-chip section-chip-operational">Operacional</span>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Fila prioritária operacional</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="subtle">Quem precisa de ação para não perder o ciclo atual.</div>',
            unsafe_allow_html=True,
        )
        if patients_df.empty:
            st.info("Nenhum paciente encontrado com os filtros selecionados.")
        else:
            watch = patients_df[patients_df["Severidade operacional"] > 0].sort_values(
                ["Severidade operacional", "Dias para ciclo"], ascending=[False, True]
            ).head(10)
            if watch.empty:
                st.success("Nenhuma pendência operacional prioritária neste filtro.")
            for _, row in watch.iterrows():
                reason, css_class, _ = evaluate_patient_alerts(row)
                delta_label, _ = build_status_flag(row["Dias para ciclo"])
                st.markdown(
                    f"""
                    <div style="padding: 14px 0; border-bottom: 1px solid rgba(18, 56, 71, 0.08);">
                        <div style="display:flex; justify-content:space-between; gap: 12px; align-items:flex-start;">
                            <div>
                                <div style="font-weight:800; color:#113847;">{row["name"]}</div>
                                <div style="color:#56707a; font-size:0.92rem;">{row["doctor_name"]} | {row["regimen"]}</div>
                                <div style="color:#56707a; font-size:0.9rem;">Próximo ciclo: {row["Próxima quimio"]} ({delta_label})</div>
                                <div style="color:#8a4b07; font-size:0.9rem; margin-top:6px;">{reason}</div>
                            </div>
                            <span class="flag {css_class}">{format_status(row["scheduling_status"], SCHEDULING_LABELS)}</span>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<span class="section-chip section-chip-protocol">Protocolo</span>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Janela de protocolo</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="subtle">Pacientes que entraram na janela de 21 dias para solicitar o próximo ciclo.</div>',
            unsafe_allow_html=True,
        )
        if patients_df.empty:
            st.info("Sem dados para resumir.")
        else:
            protocol_watch = patients_df[patients_df["Severidade protocolo"] > 0].sort_values(
                ["Severidade protocolo", "Dias para ciclo"], ascending=[False, True]
            ).copy()
            if protocol_watch.empty:
                st.success("Nenhum paciente dentro da janela de protocolo neste filtro.")
            else:
                for _, row in protocol_watch.head(8).iterrows():
                    protocol_message, protocol_flag_class, _ = evaluate_protocol_alert(row)
                    st.markdown(
                        f"""
                        <div class="protocol-card">
                            <div style="display:flex; justify-content:space-between; gap: 12px; align-items:flex-start;">
                                <div>
                                    <div class="protocol-card-title">{row["name"]}</div>
                                    <div class="protocol-card-copy">{row["doctor_name"]} | {row["regimen"]}</div>
                                    <div class="protocol-card-copy">Próxima quimio: {row["Próxima quimio"]}</div>
                                    <div class="protocol-card-copy">{protocol_message}</div>
                                </div>
                                <span class="flag {protocol_flag_class}">21 dias</span>
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("")
    table_left, table_right = st.columns([1.35, 1])
    with table_left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Painel operacional do ciclo</div>', unsafe_allow_html=True)
        if patients_df.empty:
            st.info("Nenhum paciente para consolidar.")
        else:
            display = patients_df.sort_values(["Severidade operacional", "Dias para ciclo"], ascending=[False, True]).copy()
            st.dataframe(
                display[
                    [
                        "name",
                        "doctor_name",
                        "insurance_name",
                        "Próxima quimio",
                        "Status prescrição",
                        "Status autorização",
                        "Status agenda",
                        "Alerta operacional",
                    ]
                ].rename(
                    columns={
                        "name": "Paciente",
                        "doctor_name": "Médico",
                        "insurance_name": "Convênio",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

    with table_right:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Agenda de quimioterapia</div>', unsafe_allow_html=True)
        if sessions_df.empty:
            st.info("Nenhuma sessao cadastrada.")
        else:
            display = sessions_df.copy()
            display["Data"] = display["scheduled_date"].apply(format_date)
            display["Status"] = display["status"].map(STATUS_LABELS).fillna(display["status"])
            st.dataframe(
                display[["Data", "patient_name", "doctor_name", "cycle_label", "Status", "notes"]],
                use_container_width=True,
                hide_index=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

    if not support_df.empty:
        st.markdown("")
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Suportes em atencao</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="subtle">Apoio para nao perder medicacoes de suporte nos proximos dias.</div>',
            unsafe_allow_html=True,
        )
        watch = support_df.sort_values("next_due_date").head(8).copy()
        watch["Proxima data"] = watch["next_due_date"].apply(format_date)
        watch["Status"] = watch["status"].map(STATUS_LABELS).fillna(watch["status"])
        st.dataframe(
            watch[["patient_name", "doctor_name", "medication_name", "Proxima data", "Status", "notes"]].rename(
                columns={
                    "patient_name": "Paciente",
                    "doctor_name": "Medico",
                    "medication_name": "Medicacao",
                    "notes": "Observacoes",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)


def render_register_tab(doctors_df: pd.DataFrame, patients_df: pd.DataFrame) -> None:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Novos cadastros</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtle">Registre dados do ciclo para acompanhar prescrição, convênio e agenda no mesmo lugar.</div>',
        unsafe_allow_html=True,
    )

    doctor_col, patient_col = st.columns(2)
    with doctor_col:
        with st.form("doctor_form", clear_on_submit=True):
            st.markdown("**Cadastrar medico**")
            doctor_name = st.text_input("Nome do medico")
            specialty = st.text_input("Especialidade ou foco")
            doctor_submitted = st.form_submit_button("Salvar medico", use_container_width=True)
            if doctor_submitted:
                if not doctor_name.strip():
                    st.warning("Informe o nome do medico.")
                else:
                    insert_doctor(doctor_name, specialty)
                    st.success("Médico cadastrado com sucesso.")
                    st.rerun()

    with patient_col:
        with st.form("patient_form", clear_on_submit=True):
            st.markdown("**Cadastrar paciente e situacao do ciclo**")
            doctor_options = {
                f'{row["name"]} | {row["specialty"] or "Sem especialidade"}': int(row["id"])
                for _, row in doctors_df.iterrows()
            }
            if not doctor_options:
                st.info("Cadastre um medico antes de cadastrar pacientes.")
            else:
                selected_doctor = st.selectbox("Médico responsável", list(doctor_options.keys()))
                patient_name = st.text_input("Nome do paciente")
                diagnosis = st.text_input("Diagnostico")
                regimen = st.text_input("Protocolo de quimioterapia")
                insurance_name = st.text_input("Convenio")
                cycle_interval_days = st.number_input("Intervalo entre ciclos (dias)", min_value=1, max_value=60, value=21)
                next_cycle_alert_days = st.number_input("Avisar para nova prescrição com quantos dias de antecedência", min_value=1, max_value=60, value=21)
                last_chemo_date = st.date_input("Última quimioterapia", value=None)
                next_chemo_date = st.date_input("Próxima quimioterapia prevista", value=None)
                support_plan = st.text_input("Plano de suporte")
                prescription_status = st.selectbox("Status da prescrição", list(PRESCRIPTION_LABELS.keys()), format_func=lambda x: PRESCRIPTION_LABELS[x])
                prescription_requested_date = st.date_input("Data da solicitacao ao medico", value=None)
                authorization_status = st.selectbox("Status da autorização", list(AUTHORIZATION_LABELS.keys()), format_func=lambda x: AUTHORIZATION_LABELS[x])
                authorization_submission_date = st.date_input("Data de envio ao convênio", value=None)
                authorization_valid_until = st.date_input("Autorizacao valida ate", value=None)
                scheduling_status = st.selectbox("Status do agendamento", list(SCHEDULING_LABELS.keys()), format_func=lambda x: SCHEDULING_LABELS[x])
                scheduled_cycle_date = st.date_input("Data agendada para infusao", value=None)
                notes = st.text_area("Observacoes")
                patient_submitted = st.form_submit_button("Salvar paciente", use_container_width=True)
                if patient_submitted:
                    if not patient_name.strip():
                        st.warning("Informe o nome do paciente.")
                    else:
                        insert_patient(
                            {
                                "doctor_id": doctor_options[selected_doctor],
                                "name": patient_name,
                                "diagnosis": diagnosis,
                                "regimen": regimen,
                                "cycle_interval_days": int(cycle_interval_days),
                                "last_chemo_date": last_chemo_date.strftime(DATE_FMT) if last_chemo_date else None,
                                "next_chemo_date": next_chemo_date.strftime(DATE_FMT) if next_chemo_date else None,
                                "support_plan": support_plan,
                                "notes": notes,
                                "insurance_name": insurance_name,
                                "prescription_status": prescription_status,
                                "prescription_requested_date": prescription_requested_date.strftime(DATE_FMT) if prescription_requested_date else None,
                                "authorization_status": authorization_status,
                                "authorization_submission_date": authorization_submission_date.strftime(DATE_FMT) if authorization_submission_date else None,
                                "authorization_valid_until": authorization_valid_until.strftime(DATE_FMT) if authorization_valid_until else None,
                                "scheduling_status": scheduling_status,
                                "scheduled_cycle_date": scheduled_cycle_date.strftime(DATE_FMT) if scheduled_cycle_date else None,
                                "next_cycle_alert_days": int(next_cycle_alert_days),
                            }
                        )
                        st.success("Paciente cadastrado com sucesso.")
                        st.rerun()

    st.markdown("---")
    support_col, session_col = st.columns(2)
    patient_options = {
        f'{row["name"]} | {row["doctor_name"]}': int(row["id"])
        for _, row in patients_df.iterrows()
    }
    with support_col:
        with st.form("support_form", clear_on_submit=True):
            st.markdown("**Cadastrar medicacao de suporte**")
            if not patient_options:
                st.info("Cadastre um paciente antes de registrar suporte.")
            else:
                patient_key = st.selectbox("Paciente", list(patient_options.keys()))
                medication_name = st.text_input("Medicacao")
                purpose = st.text_input("Finalidade")
                frequency_label = st.text_input("Frequencia ou orientacao")
                next_due_date = st.date_input("Proxima data", value=date.today())
                support_status = st.selectbox("Status", list(STATUS_LABELS.keys()), format_func=lambda x: STATUS_LABELS[x])
                support_notes = st.text_area("Observacoes do suporte")
                support_submitted = st.form_submit_button("Salvar suporte", use_container_width=True)
                if support_submitted:
                    if not medication_name.strip():
                        st.warning("Informe o nome da medicacao.")
                    else:
                        insert_support_med(
                            {
                                "patient_id": patient_options[patient_key],
                                "medication_name": medication_name,
                                "purpose": purpose,
                                "frequency_label": frequency_label,
                                "next_due_date": next_due_date.strftime(DATE_FMT),
                                "status": support_status,
                                "notes": support_notes,
                            }
                        )
                        st.success("Medicacao de suporte salva.")
                        st.rerun()

    with session_col:
        with st.form("session_form", clear_on_submit=True):
            st.markdown("**Cadastrar sessao de quimioterapia**")
            if not patient_options:
                st.info("Cadastre um paciente antes de registrar sessao.")
            else:
                patient_key = st.selectbox("Paciente ", list(patient_options.keys()))
                scheduled_date = st.date_input("Data programada", value=date.today())
                cycle_label = st.text_input("Ciclo ou observacao")
                session_status = st.selectbox("Status da sessao", ["scheduled", "done", "attention"], format_func=lambda x: STATUS_LABELS[x])
                session_notes = st.text_area("Observacoes da sessao")
                session_submitted = st.form_submit_button("Salvar sessao", use_container_width=True)
                if session_submitted:
                    insert_chemo_session(
                        {
                            "patient_id": patient_options[patient_key],
                            "scheduled_date": scheduled_date.strftime(DATE_FMT),
                            "cycle_label": cycle_label,
                            "status": session_status,
                            "notes": session_notes,
                        }
                    )
                    st.success("Sessao registrada.")
                    st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)


def render_import_tab() -> None:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Importar planilha de pacientes</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtle">Você pode subir uma planilha do Excel Online em formato .xlsx ou .csv para cadastrar pacientes em lote e gerar os agendamentos no app.</div>',
        unsafe_allow_html=True,
    )

    st.markdown("**Colunas esperadas na planilha**")
    template_df = pd.DataFrame(
        [
            {
                "medico": "Dra. Camila Torres",
                "especialidade": "Oncologia Clinica",
                "paciente": "Ana Paula Lima",
                "diagnostico": "CA colorretal",
                "protocolo": "FOLFOX",
                "intervalo_dias": 14,
                "ultima_quimio": "2026-04-20",
                "proxima_quimio": "2026-05-04",
                "convenio": "Bradesco Saude",
                "status_prescricao": "prescribed",
                "data_solicitacao_prescricao": "2026-04-25",
                "status_autorizacao": "authorized",
                "data_envio_convenio": "2026-04-26",
                "autorizacao_valida_ate": "2026-05-30",
                "status_agendamento": "confirmed",
                "data_agendada": "2026-05-04",
                "alerta_novo_ciclo_dias": 21,
                "ciclo": "Ciclo 4",
                "suporte": "Ondansetrona + dexametasona",
                "observacoes": "Paciente em seguimento semanal.",
                "observacoes_agenda": "Sessao importada da planilha.",
            }
        ]
    )
    st.dataframe(template_df, use_container_width=True, hide_index=True)

    st.caption(
        "Status aceitos: prescrição = not_requested, requested, prescribed, sent_to_insurance | autorização = not_sent, pending, authorized, denied | agendamento = not_booked, awaiting_slot, scheduled, confirmed"
    )

    uploaded_file = st.file_uploader(
        "Selecione a planilha",
        type=["xlsx", "csv"],
        accept_multiple_files=False,
    )

    if uploaded_file is not None:
        try:
            preview_df = load_uploaded_dataframe(uploaded_file)
            st.markdown("**Previa da planilha**")
            st.dataframe(preview_df.head(20), use_container_width=True, hide_index=True)

            if st.button("Importar planilha para o app", use_container_width=True):
                imported_patients, imported_sessions = import_patients_dataframe(preview_df)
                st.success(
                    f"Importacao concluida: {imported_patients} paciente(s) cadastrados e {imported_sessions} agendamento(s) criado(s)."
                )
                st.rerun()
        except Exception as exc:
            st.error(f"Não consegui ler a planilha. Verifique as colunas e o formato do arquivo. Detalhe: {exc}")

    st.markdown("</div>", unsafe_allow_html=True)


def render_google_sync_tab(patients_df: pd.DataFrame) -> None:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Planilha principal</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtle">Sincronize as abas dos médicos da planilha principal e mantenha o painel alinhado ao arquivo fonte.</div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns([1, 1.3])
    with col1:
        if st.button("Sincronizar agora com a planilha principal", use_container_width=True):
            try:
                imported, updated = sync_google_sheets_to_db()
                st.success(f"Sincronização concluída: {imported} novo(s) e {updated} atualizado(s).")
                st.rerun()
            except Exception as exc:
                st.error(f"Não consegui sincronizar com a planilha principal. Detalhe: {exc}")
        st.caption("A sincronizacao usa as abas com nome de medico, como `Dr.` e `Dra.`.")

    with col2:
        workbook_file = find_primary_workbook_file()
        last_sync = get_app_state("last_google_sync_at")
        st.write("**Arquivo fonte**")
        st.write(f"Planilha: `{workbook_file.name if workbook_file else 'não encontrada'}`")
        if workbook_file is not None:
            st.write(f"Caminho: `{workbook_file}`")
        st.write(f"Última sincronização: `{format_sync_timestamp(last_sync)}`")

    st.markdown("---")
    st.markdown("**Trocar a fonte da planilha**")
    st.caption(
        "Para uso na web, você pode enviar um novo arquivo .xlsx aqui e ele passa a ser a fonte principal do app."
    )
    uploaded_workbook = st.file_uploader(
        "Enviar nova planilha principal",
        type=["xlsx"],
        accept_multiple_files=False,
        key="primary_workbook_uploader",
    )
    if uploaded_workbook is not None:
        if st.button("Usar esta planilha como fonte principal", use_container_width=True, key="save_primary_workbook"):
            try:
                saved_path = save_uploaded_primary_workbook(uploaded_workbook)
                imported, updated = sync_google_sheets_to_db()
                st.success(
                    f"Nova fonte salva em `{saved_path.name}` e sincronizada com sucesso: {imported} novo(s) e {updated} atualizado(s)."
                )
                st.rerun()
            except Exception as exc:
                st.error(f"Não consegui trocar a planilha principal. Detalhe: {exc}")

    google_patients = patients_df[patients_df["source_sheet_name"].notna()].copy() if not patients_df.empty else pd.DataFrame()
    st.markdown("---")
    st.markdown("**Editar paciente sincronizado**")
    if google_patients.empty:
        st.info("Nenhum paciente sincronizado da planilha ainda. Clique em sincronizar primeiro.")
    else:
        patient_options = {
            f'{row["name"]} | {row["doctor_name"]}': int(row["id"])
            for _, row in google_patients.iterrows()
        }
        selected_label = st.selectbox("Paciente importado da planilha", list(patient_options.keys()))
        selected_id = patient_options[selected_label]
        patient_row = google_patients[google_patients["id"] == selected_id].iloc[0]

        with st.form("google_patient_edit_form"):
            diagnosis = st.text_input("Diagnóstico", value=patient_row["diagnosis"] or "")
            regimen = st.text_input("Tratamento / protocolo", value=patient_row["regimen"] or "")
            next_chemo = st.date_input("Próxima infusão", value=parse_date(patient_row["next_chemo_date"]))
            protocol_next_cycle = st.date_input("Próxima data do protocolo", value=parse_date(patient_row["protocol_next_cycle_date"]))
            protocol_alert_days = st.number_input("Janela do protocolo (dias antes)", min_value=1, max_value=90, value=int(patient_row["next_cycle_alert_days"] or 21))
            insurance_name = st.text_input("Convênio", value=patient_row["insurance_name"] or "")
            notes = st.text_area("Observações", value=patient_row["notes"] or "")
            st.caption("Os flags de prescrição, autorização e agendamento agora são controlados por ciclo dentro do calendário.")
            submitted = st.form_submit_button("Salvar no app e na planilha", use_container_width=True)
            if submitted:
                updates = {
                    "diagnosis": diagnosis.strip(),
                    "regimen": regimen.strip(),
                    "next_chemo_date": next_chemo.strftime(DATE_FMT) if next_chemo else None,
                    "scheduled_cycle_date": next_chemo.strftime(DATE_FMT) if next_chemo else None,
                    "protocol_next_cycle_date": protocol_next_cycle.strftime(DATE_FMT) if protocol_next_cycle else None,
                    "next_cycle_alert_days": int(protocol_alert_days),
                    "insurance_name": insurance_name.strip(),
                    "notes": notes.strip(),
                }
                try:
                    update_patient_record(selected_id, updates)
                    fresh_patients = load_patients()
                    fresh_row = fresh_patients[fresh_patients["id"] == selected_id].iloc[0]
                    update_google_sheet_patient_row(fresh_row, updates)
                    st.success("Paciente atualizado no app e gravado na planilha local.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Não consegui salvar a atualização. Detalhe: {exc}")

    st.markdown("</div>", unsafe_allow_html=True)


def render_patients_tab(filtered_patients: pd.DataFrame) -> None:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Base de pacientes</div>', unsafe_allow_html=True)
    if filtered_patients.empty:
        st.info("Nenhum paciente encontrado.")
    else:
        display = build_operational_table(filtered_patients)
        st.dataframe(
            display[
                [
                    "name",
                    "doctor_name",
                    "insurance_name",
                    "diagnosis",
                    "regimen",
                    "Próxima quimio",
                    "Próximo protocolo",
                    "Status prescrição",
                    "Status autorização",
                    "Status agenda",
                    "Alerta operacional",
                    "Alerta de protocolo",
                    "support_plan",
                    "notes",
                ]
            ].rename(
                columns={
                    "name": "Paciente",
                    "doctor_name": "Médico",
                    "insurance_name": "Convênio",
                    "diagnosis": "Diagnóstico",
                    "regimen": "Protocolo",
                    "support_plan": "Suporte",
                    "notes": "Observações",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)


def render_alerts_tab(filtered_patients: pd.DataFrame) -> None:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Alertas e priorização</div>', unsafe_allow_html=True)
    if filtered_patients.empty:
        st.info("Nenhum paciente encontrado.")
    else:
        display = build_operational_table(filtered_patients)
        operational_alerts = display[display["Severidade operacional"] > 0].sort_values(
            ["Severidade operacional", "Dias para ciclo"], ascending=[False, True]
        )
        protocol_alerts = display[display["Severidade protocolo"] > 0].sort_values(
            ["Severidade protocolo", "Dias para ciclo"], ascending=[False, True]
        )

        st.markdown('<span class="section-chip section-chip-operational">Operacional</span>', unsafe_allow_html=True)
        st.markdown("**Operacional agora**")
        st.caption("Aqui ficam apenas os riscos do ciclo atual: agenda, autorização e atraso.")
        if operational_alerts.empty:
            st.success("Nenhum alerta operacional importante neste filtro.")
        else:
            operational_alerts["Dias para ciclo"] = operational_alerts["Dias para ciclo"].fillna("-")
            st.dataframe(
                operational_alerts[
                    [
                        "name",
                        "doctor_name",
                        "Próxima quimio",
                        "Dias para ciclo",
                        "Status prescrição",
                        "Status autorização",
                        "Status agenda",
                        "Alerta operacional",
                    ]
                ].rename(
                    columns={
                        "name": "Paciente",
                        "doctor_name": "Médico",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

        st.markdown("---")
        st.markdown('<span class="section-chip section-chip-protocol">Protocolo</span>', unsafe_allow_html=True)
        st.markdown("**Protocolos em janela de prescrição**")
        st.caption("Aqui ficam os pacientes que atingiram a janela de 21 dias para solicitar o próximo ciclo.")
        if protocol_alerts.empty:
            st.success("Nenhum protocolo entrou na janela de prescrição neste filtro.")
        else:
            protocol_alerts["Dias para ciclo"] = protocol_alerts["Dias para ciclo"].fillna("-")
            st.dataframe(
                protocol_alerts[
                    [
                        "name",
                        "doctor_name",
                        "Próxima quimio",
                        "Dias para ciclo",
                        "Status prescrição",
                        "Alerta de protocolo",
                    ]
                ].rename(
                    columns={
                        "name": "Paciente",
                        "doctor_name": "Médico",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )
    st.markdown("</div>", unsafe_allow_html=True)


def render_support_tab(filtered_support: pd.DataFrame) -> None:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Medicacoes de suporte</div>', unsafe_allow_html=True)
    if filtered_support.empty:
        st.info("Nenhuma medicacao de suporte encontrada.")
    else:
        display = filtered_support.copy()
        display["Proxima data"] = display["next_due_date"].apply(format_date)
        display["Status"] = display["status"].map(STATUS_LABELS).fillna(display["status"])
        st.dataframe(
            display[
                [
                    "patient_name",
                    "doctor_name",
                    "medication_name",
                    "purpose",
                    "frequency_label",
                    "Proxima data",
                    "Status",
                    "notes",
                ]
            ].rename(
                columns={
                    "patient_name": "Paciente",
                    "doctor_name": "Medico",
                    "medication_name": "Medicacao",
                    "purpose": "Finalidade",
                    "frequency_label": "Frequencia",
                    "notes": "Observacoes",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)


def render_doctors_tab(doctors_df: pd.DataFrame, patients_df: pd.DataFrame) -> None:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Medicos e carteiras</div>', unsafe_allow_html=True)
    if doctors_df.empty:
        st.info("Nenhum medico cadastrado.")
    else:
        if patients_df.empty:
            merged = doctors_df.copy()
            merged["total_pacientes"] = 0
            merged["sem_agenda"] = 0
            merged["pendentes_convenio"] = 0
        else:
            base = build_operational_table(patients_df)
            count_by_doctor = (
                base.groupby("doctor_id", as_index=False)
                .agg(
                    total_pacientes=("id", "count"),
                    sem_agenda=("scheduling_status", lambda s: s.isin(["not_booked", "awaiting_slot"]).sum()),
                    pendentes_convenio=("authorization_status", lambda s: s.isin(["not_sent", "pending", "denied"]).sum()),
                )
            )
            merged = doctors_df.merge(count_by_doctor, how="left", left_on="id", right_on="doctor_id").fillna(
                {"total_pacientes": 0, "sem_agenda": 0, "pendentes_convenio": 0}
            )
        for column in ["total_pacientes", "sem_agenda", "pendentes_convenio"]:
            merged[column] = merged[column].astype(int)
        st.dataframe(
            merged[["name", "specialty", "total_pacientes", "sem_agenda", "pendentes_convenio"]].rename(
                columns={
                    "name": "Médico",
                    "specialty": "Especialidade",
                    "total_pacientes": "Pacientes",
                    "sem_agenda": "Sem agenda",
                    "pendentes_convenio": "Pendentes convênio",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)


def maybe_auto_sync_google() -> None:
    if find_primary_workbook_file() is None:
        return

    now = datetime.now(APP_TIMEZONE)
    last_sync_raw = get_app_state("last_google_sync_at")
    should_sync = False

    if not last_sync_raw:
        should_sync = True
    else:
        try:
            last_sync = datetime.fromisoformat(last_sync_raw)
            if last_sync.tzinfo is None:
                last_sync = last_sync.replace(tzinfo=APP_TIMEZONE)
            should_sync = now - last_sync >= timedelta(minutes=AUTO_SYNC_MINUTES)
        except ValueError:
            should_sync = True

    if should_sync:
        try:
            sync_google_sheets_to_db()
        except Exception:
            pass


def inject_auto_refresh() -> None:
    refresh_ms = AUTO_SYNC_MINUTES * 60 * 1000
    components.html(
        f"""
        <script>
            setTimeout(function() {{
                window.parent.location.reload();
            }}, {refresh_ms});
        </script>
        """,
        height=0,
    )


def main() -> None:
    init_db()
    ensure_auth_session_state()

    st.markdown(APP_CSS, unsafe_allow_html=True)
    if not st.session_state.get("auth_user"):
        render_login_gate()
        return

    maybe_auto_sync_google()
    inject_auto_refresh()
    sync_navigation_state_from_query_params()

    st.markdown(
        """
        <div class="hero">
            <h1 style="color:#ffffff !important;">Navegação Oncológica</h1>
            <p style="color:#f3fbff !important;">
                <span style="color:#f3fbff !important;">
                    Painel para antecipar o próximo ciclo dos pacientes, conferir prescrição,
                    autorização do convênio e agendamento da quimioterapia, reduzindo o risco
                    de o paciente ficar fora da agenda.
                </span>
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    doctors_df = load_doctors()
    patients_df = load_patients()
    support_df = load_support_medications()
    sessions_df = load_chemo_sessions()

    with st.sidebar:
        st.caption(f"Acesso: {st.session_state.get('auth_user')}")
        if st.button("Encerrar sessão", use_container_width=True):
            st.session_state["auth_user"] = None
            close_patient_detail()
            st.rerun()
        st.markdown("---")
        st.markdown("### Filtros")
        doctor_options = ["Todos"] + doctors_df["name"].tolist()
        selected_doctor = st.selectbox("Médico", doctor_options)

        patient_names = patients_df["name"].tolist()
        selected_patient = st.selectbox("Paciente", ["Todos"] + patient_names)

        show_only_attention = st.toggle("Mostrar apenas pacientes com alerta", value=False)
        st.markdown("---")
        st.caption("Versão focada em prescrição, convênio e agenda do ciclo.")
        if find_primary_workbook_file() is not None:
            last_sync = get_app_state("last_google_sync_at")
            st.caption("Planilha principal conectada.")
            st.caption(f"Última sincronização: {last_sync or 'ainda não sincronizado'}")
            st.caption(f"Atualização automática a cada {AUTO_SYNC_MINUTES} minutos.")

    filtered_patients = patients_df.copy()
    filtered_support = support_df.copy()
    filtered_sessions = sessions_df.copy()

    if selected_doctor != "Todos":
        filtered_patients = filtered_patients[filtered_patients["doctor_name"] == selected_doctor]
        filtered_support = filtered_support[filtered_support["doctor_name"] == selected_doctor]
        filtered_sessions = filtered_sessions[filtered_sessions["doctor_name"] == selected_doctor]

    if selected_patient != "Todos":
        filtered_patients = filtered_patients[filtered_patients["name"] == selected_patient]
        filtered_support = filtered_support[filtered_support["patient_name"] == selected_patient]
        filtered_sessions = filtered_sessions[filtered_sessions["patient_name"] == selected_patient]

    if show_only_attention and not filtered_patients.empty:
        alerts = build_operational_table(filtered_patients)
        alert_names = alerts[
            (alerts["Severidade operacional"] > 0) | (alerts["Severidade protocolo"] > 0)
        ]["name"].tolist()
        filtered_patients = filtered_patients[filtered_patients["name"].isin(alert_names)]
        filtered_support = filtered_support[filtered_support["patient_name"].isin(alert_names)]
        filtered_sessions = filtered_sessions[filtered_sessions["patient_name"].isin(alert_names)]

    if st.session_state.get("current_view") == "patient_detail":
        render_calendar_patient_detail_page(patients_df, sessions_df)
        return

    tabs = st.tabs(["Visão simples", "Painel operacional", "Alertas", "Pacientes", "Médicos", "Cadastros", "Importação", "Planilha principal"])
    with tabs[0]:
        render_simple_dashboard(filtered_patients, filtered_support, filtered_sessions)
    with tabs[1]:
        render_dashboard(filtered_patients, filtered_support, filtered_sessions)
    with tabs[2]:
        render_alerts_tab(filtered_patients)
    with tabs[3]:
        render_patients_tab(filtered_patients)
    with tabs[4]:
        render_doctors_tab(doctors_df, filtered_patients)
    with tabs[5]:
        render_register_tab(doctors_df, patients_df)
    with tabs[6]:
        render_import_tab()
    with tabs[7]:
        render_google_sync_tab(patients_df)


if __name__ == "__main__":
    main()
