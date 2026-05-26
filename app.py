from __future__ import annotations

import html
import hashlib
import hmac
import os
import json
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path
import re
import unicodedata
from urllib.parse import parse_qs, quote, unquote, urlencode, urlparse
from urllib.request import Request, urlopen
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
PENDING_WORKBOOK_NAME = "pending_primary_workbook.xlsx"
MICROSOFT_WORKBOOK_URL_KEY = "microsoft_online_workbook_url"
LAST_MICROSOFT_DOWNLOAD_KEY = "last_microsoft_download_at"
LOCAL_WORKBOOK_PATH_KEY = "local_workbook_path"
INCLUDED_WORKBOOK_SHEETS_KEY = "included_workbook_sheets"
REMEMBERED_AUTH_USER_KEY = "remembered_auth_user"
ONEDRIVE_CLOUDSTORAGE_DIR = Path.home() / "Library" / "CloudStorage"
APP_TIMEZONE = ZoneInfo("America/Sao_Paulo")
PRODUCT_NAME = "OncoNavega"
PRODUCT_TAGLINE = "Navegação oncológica para proteger ciclos, agenda e receita."
PRODUCT_PROMISE = "Da planilha à fila prioritária: prescrição, autorização e agenda no mesmo fluxo."
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
    page_title=PRODUCT_NAME,
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
    background:
        linear-gradient(135deg, rgba(11, 59, 72, 0.98) 0%, rgba(22, 105, 122, 0.98) 58%, rgba(42, 157, 143, 0.96) 100%);
    border-radius: 26px;
    color: #ffffff !important;
    padding: 30px 32px;
    box-shadow: 0 22px 48px rgba(15, 61, 76, 0.18);
    margin-bottom: 18px;
    position: relative;
    overflow: hidden;
}

.hero, .hero * {
    color: #ffffff !important;
}

.hero h1 {
    margin: 0;
    font-size: 2.72rem;
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

.brand-row {
    display: flex;
    gap: 12px;
    align-items: center;
    margin-bottom: 12px;
}

.brand-mark {
    width: 42px;
    height: 42px;
    border-radius: 14px;
    background: rgba(255, 255, 255, 0.16);
    border: 1px solid rgba(255, 255, 255, 0.24);
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-weight: 800;
    color: #ffffff;
}

.brand-kicker {
    color: #d9fbff !important;
    font-size: 0.82rem;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}

.hero-actions {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin-top: 18px;
}

.hero-pill {
    background: rgba(255, 255, 255, 0.14);
    border: 1px solid rgba(255, 255, 255, 0.22);
    border-radius: 999px;
    color: #ffffff !important;
    font-size: 0.86rem;
    font-weight: 800;
    padding: 8px 12px;
}

.product-strip {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 12px;
    margin: 0 0 18px 0;
}

.product-proof {
    background: rgba(255, 255, 255, 0.88);
    border: 1px solid rgba(15, 61, 76, 0.08);
    border-radius: 18px;
    padding: 14px 16px;
    box-shadow: 0 12px 26px rgba(26, 55, 77, 0.06);
}

.product-proof-label {
    color: #5b7681;
    font-size: 0.78rem;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

.product-proof-value {
    color: #123847;
    font-size: 1.05rem;
    font-weight: 800;
    margin-top: 5px;
}

@media (max-width: 900px) {
    .product-strip {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }
}

@media (max-width: 560px) {
    .product-strip {
        grid-template-columns: 1fr;
    }
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

.commercial-card {
    background: rgba(255, 255, 255, 0.92);
    border: 1px solid rgba(15, 61, 76, 0.08);
    border-radius: 18px;
    padding: 16px 18px;
    min-height: 170px;
    box-shadow: 0 12px 28px rgba(26, 55, 77, 0.07);
}

.commercial-card h3 {
    color: #123847;
    font-size: 1rem;
    margin: 0 0 8px 0;
    font-weight: 800;
}

.commercial-card p,
.commercial-card li {
    color: #45636f;
    font-size: 0.92rem;
    line-height: 1.48;
}

.commercial-card ul {
    padding-left: 18px;
    margin: 8px 0 0 0;
}

.price-card {
    background: linear-gradient(180deg, rgba(255, 255, 255, 0.96) 0%, rgba(239, 248, 251, 0.96) 100%);
    border: 1px solid rgba(15, 61, 76, 0.10);
    border-radius: 18px;
    padding: 18px;
    min-height: 255px;
    box-shadow: 0 14px 32px rgba(26, 55, 77, 0.08);
}

.price-name {
    font-size: 0.82rem;
    font-weight: 800;
    color: #16697a;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

.price-value {
    color: #123847;
    font-size: 1.8rem;
    font-weight: 800;
    margin-top: 6px;
}

.price-note {
    color: #5b7681;
    font-size: 0.9rem;
    margin: 4px 0 12px 0;
}

.sales-step {
    border-left: 4px solid #16697a;
    padding: 4px 0 4px 12px;
    margin-bottom: 12px;
}

.sales-step-title {
    color: #123847;
    font-weight: 800;
}

.sales-step-copy {
    color: #56707a;
    font-size: 0.92rem;
    margin-top: 2px;
}

.playbook-card {
    background: rgba(255, 255, 255, 0.94);
    border: 1px solid rgba(15, 61, 76, 0.08);
    border-radius: 18px;
    padding: 18px;
    box-shadow: 0 12px 28px rgba(26, 55, 77, 0.07);
}

.playbook-card h3 {
    margin: 0 0 8px 0;
    color: #123847;
    font-size: 1.02rem;
    font-weight: 800;
}

.playbook-card p,
.playbook-card li {
    color: #4f6d78;
    font-size: 0.92rem;
    line-height: 1.48;
}

.pilot-week {
    background: #f7fbfd;
    border: 1px solid rgba(15, 61, 76, 0.08);
    border-left: 4px solid #2a9d8f;
    border-radius: 14px;
    padding: 12px 14px;
    margin-bottom: 10px;
}

.pilot-week-title {
    color: #123847;
    font-weight: 800;
}

.pilot-week-copy {
    color: #56707a;
    font-size: 0.92rem;
    margin-top: 3px;
}

.calendar-scroll {
    width: 100%;
    overflow-x: auto;
    padding-bottom: 8px;
}

.calendar-grid {
    display: grid;
    grid-template-columns: repeat(7, minmax(150px, 1fr));
    gap: 10px;
    min-width: 1120px;
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
    height: 270px;
    background: rgba(255, 255, 255, 0.92);
    border: 1px solid rgba(15, 61, 76, 0.08);
    border-radius: 18px;
    padding: 10px;
    box-shadow: 0 10px 24px rgba(26, 55, 77, 0.05);
    overflow: hidden;
    display: flex;
    flex-direction: column;
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

.calendar-events {
    flex: 1;
    min-height: 0;
    overflow-y: auto;
    padding-right: 3px;
}

.calendar-events::-webkit-scrollbar {
    width: 6px;
}

.calendar-events::-webkit-scrollbar-thumb {
    background: rgba(15, 61, 76, 0.18);
    border-radius: 999px;
}

.calendar-doctor-block {
    border-left: 4px solid var(--doctor-color, #2a9d8f);
    background: color-mix(in srgb, var(--doctor-color, #2a9d8f) 9%, #ffffff);
    border-radius: 12px;
    padding: 6px 7px;
    margin-bottom: 7px;
}

.calendar-doctor-group {
    font-size: 0.66rem;
    font-weight: 800;
    color: var(--doctor-color, #2a9d8f);
    letter-spacing: 0.01em;
    margin: 0 0 4px 0;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

.calendar-patient-link {
    display: block;
    color: #123847 !important;
    font-size: 0.73rem;
    font-weight: 800;
    line-height: 1.2;
    text-decoration: none !important;
    background: rgba(255, 255, 255, 0.72);
    border: 1px solid rgba(15, 61, 76, 0.07);
    border-radius: 9px;
    padding: 5px 6px;
    margin-top: 4px;
    overflow-wrap: anywhere;
}

.calendar-patient-link:hover {
    background: #ffffff;
    border-color: rgba(15, 61, 76, 0.18);
}

.calendar-more {
    display: inline-block;
    background: #eef2f5;
    color: #4b6370;
    font-size: 0.72rem;
    font-weight: 800;
    border-radius: 999px;
    padding: 4px 7px;
    margin-top: 3px;
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
.stMultiSelect > div > div,
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

.stMultiSelect [data-baseweb="tag"] {
    background: #e7f4f7 !important;
    border: 1px solid rgba(15, 61, 76, 0.16) !important;
    border-radius: 10px !important;
    color: #123847 !important;
}

.stMultiSelect [data-baseweb="tag"] span,
.stMultiSelect [data-baseweb="tag"] p,
.stMultiSelect [data-baseweb="tag"] div {
    color: #123847 !important;
    font-weight: 800 !important;
}

.stMultiSelect [data-baseweb="tag"] svg {
    fill: #31556a !important;
}

.stMultiSelect [data-baseweb="select"] input {
    color: #123847 !important;
}

.stMultiSelect svg {
    fill: #31556a !important;
}

.stCheckbox [data-baseweb="checkbox"] {
    background: #ffffff !important;
}

.stCheckbox svg {
    fill: #0f4c5c !important;
}

[data-testid="stTabs"] {
    background: rgba(255, 255, 255, 0.86);
    border: 1px solid rgba(15, 61, 76, 0.08);
    border-radius: 18px;
    padding: 8px 10px 12px 10px;
    box-shadow: 0 12px 26px rgba(26, 55, 77, 0.06);
}

[data-testid="stTabs"] [role="tablist"] {
    gap: 6px;
    border-bottom: 1px solid rgba(15, 61, 76, 0.10);
    padding-bottom: 8px;
}

[data-testid="stTabs"] [role="tab"] {
    background: #f6fbfd !important;
    border: 1px solid rgba(15, 61, 76, 0.12) !important;
    border-radius: 999px !important;
    color: #214755 !important;
    min-height: 38px !important;
    padding: 7px 13px !important;
    box-shadow: none !important;
}

[data-testid="stTabs"] [role="tab"] p {
    color: #214755 !important;
    font-weight: 800 !important;
    font-size: 0.86rem !important;
}

[data-testid="stTabs"] [role="tab"]:hover {
    background: #e8f4f8 !important;
    border-color: rgba(15, 61, 76, 0.18) !important;
}

[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    background: linear-gradient(135deg, #0f3d4c 0%, #16697a 100%) !important;
    border-color: rgba(15, 61, 76, 0.34) !important;
    box-shadow: 0 10px 18px rgba(15, 61, 76, 0.16) !important;
}

[data-testid="stTabs"] [role="tab"][aria-selected="true"] p {
    color: #ffffff !important;
}

[data-testid="stTabs"] [data-baseweb="tab-highlight"] {
    display: none !important;
}

[data-testid="stTabs"] [role="tabpanel"] {
    padding-top: 18px;
}

.stButton button:not([kind="tertiary"]),
.stForm button:not([kind="tertiary"]),
.stForm [data-testid="stFormSubmitButton"] button {
    background: linear-gradient(135deg, #ffffff 0%, #eef8fc 100%) !important;
    color: #0f3d4c !important;
    border: 1px solid rgba(15, 61, 76, 0.22) !important;
    box-shadow: 0 10px 20px rgba(15, 61, 76, 0.10) !important;
    font-weight: 800 !important;
}

.stButton button:not([kind="tertiary"]):hover,
.stButton button:not([kind="tertiary"]):focus,
.stForm button:not([kind="tertiary"]):hover,
.stForm button:not([kind="tertiary"]):focus,
.stForm [data-testid="stFormSubmitButton"] button:hover,
.stForm [data-testid="stFormSubmitButton"] button:focus {
    background: linear-gradient(135deg, #ffffff 0%, #e2f1f8 100%) !important;
    color: #0f3d4c !important;
    border: 1px solid rgba(15, 61, 76, 0.28) !important;
}

.stButton button:not([kind="tertiary"]) p,
.stButton button:not([kind="tertiary"]) span,
.stForm button:not([kind="tertiary"]) p,
.stForm button:not([kind="tertiary"]) span,
.stForm [data-testid="stFormSubmitButton"] button p,
.stForm [data-testid="stFormSubmitButton"] button span {
    color: #0f3d4c !important;
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

.hero p,
.hero span,
.hero div,
.hero h1 {
    color: #ffffff !important;
}

.hero .brand-kicker {
    color: #d9fbff !important;
}

@media (max-width: 560px) {
    .hero {
        padding: 26px 24px;
        border-radius: 22px;
    }

    .hero h1 {
        font-size: 2.18rem;
        line-height: 1.16;
    }
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


def build_auth_token(username: str) -> str:
    credentials = load_access_credentials()
    secret = os.getenv("ONCO_APP_AUTH_SECRET") or credentials["password"]
    message = f"{credentials['username'].strip().lower()}:{username.strip().lower()}".encode("utf-8")
    return hmac.new(secret.encode("utf-8"), message, hashlib.sha256).hexdigest()


def restore_auth_from_query_params() -> None:
    if st.session_state.get("auth_user"):
        return
    remembered_user = get_app_state(REMEMBERED_AUTH_USER_KEY)
    if remembered_user:
        st.session_state["auth_user"] = remembered_user
        return
    auth_user = read_query_param("auth_user")
    auth_token = read_query_param("auth_token")
    if not auth_user or not auth_token:
        return
    expected_token = build_auth_token(auth_user)
    if hmac.compare_digest(auth_token, expected_token):
        st.session_state["auth_user"] = auth_user.strip()


def ensure_persistent_auth_query_params() -> None:
    auth_user = st.session_state.get("auth_user")
    if not auth_user:
        return
    current_user = read_query_param("auth_user")
    current_token = read_query_param("auth_token")
    expected_token = build_auth_token(str(auth_user))
    if current_user != str(auth_user) or current_token != expected_token:
        persist_auth_in_query_params(str(auth_user))
        components.html(
            f"""
            <script>
                const url = new URL(window.parent.location.href);
                url.searchParams.set("auth_user", {json.dumps(str(auth_user))});
                url.searchParams.set("auth_token", {json.dumps(expected_token)});
                window.parent.history.replaceState(null, "", url.toString());
            </script>
            """,
            height=0,
        )


def persist_auth_in_query_params(username: str) -> None:
    st.query_params.update(
        {
            "auth_user": username.strip(),
            "auth_token": build_auth_token(username),
        }
    )


def clear_auth_query_params() -> None:
    for key in ["auth_user", "auth_token"]:
        if key in st.query_params:
            del st.query_params[key]


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
            <div class="brand-row">
                <div class="brand-mark">ON</div>
                <div>
                    <div class="brand-kicker">OncoNavega</div>
                    <div style="color:#d9fbff !important; font-weight:700;">Navegação oncológica operacional</div>
                </div>
            </div>
            <h1 style="color:#ffffff !important;">Proteja ciclos, agenda e receita.</h1>
            <p style="color:#f3fbff !important;">
                Acesso protegido para organizar prescrição, autorização e agendamento dos pacientes em tratamento oncológico.
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
                normalized_username = username.strip()
                st.session_state["auth_user"] = normalized_username
                set_app_state(REMEMBERED_AUTH_USER_KEY, normalized_username)
                persist_auth_in_query_params(normalized_username)
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


def delete_app_state(key: str) -> None:
    conn = get_connection()
    conn.execute("DELETE FROM app_state WHERE key = ?", (key,))
    conn.commit()
    conn.close()


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
    local_workbook_path = get_app_state(LOCAL_WORKBOOK_PATH_KEY)
    if local_workbook_path:
        local_path = clean_local_workbook_path(local_workbook_path)
        if local_path.exists() and local_path.suffix.lower() == ".xlsx":
            return local_path

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


def save_pending_primary_workbook(uploaded_file) -> Path:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    target = DATA_DIR / PENDING_WORKBOOK_NAME
    target.write_bytes(uploaded_file.getbuffer())
    return target


def promote_pending_primary_workbook() -> Path:
    pending = DATA_DIR / PENDING_WORKBOOK_NAME
    if not pending.exists():
        raise FileNotFoundError("Nenhuma planilha pendente encontrada para confirmar.")
    target = DATA_DIR / UPLOADED_WORKBOOK_NAME
    pending.replace(target)
    return target


def clean_local_workbook_path(value: str) -> Path:
    cleaned = (value or "").strip().strip("\"'")
    return Path(cleaned).expanduser()


def microsoft_download_candidates(shared_url: str) -> list[str]:
    url = shared_url.strip()
    if not url:
        return []

    candidates: list[str] = []
    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    host = parsed.netloc.lower()
    path = parsed.path

    if "onedrive.live.com" in host and "resid" in query:
        download_query = {"resid": query["resid"][0]}
        if "authkey" in query:
            download_query["authkey"] = query["authkey"][0]
        candidates.append(f"https://onedrive.live.com/download?{urlencode(download_query)}")

    if "sharepoint.com" in host:
        if "sourcedoc" in query:
            candidates.append(f"{parsed.scheme}://{parsed.netloc}/_layouts/15/download.aspx?{urlencode({'UniqueId': query['sourcedoc'][0].strip('{}')})}")

        if "/:x:/r/" in path:
            source_path = unquote(path.split("/:x:/r", 1)[1])
            candidates.append(
                f"{parsed.scheme}://{parsed.netloc}/_layouts/15/download.aspx?{urlencode({'SourceUrl': source_path})}"
            )
            candidates.append(
                f"{parsed.scheme}://{parsed.netloc}/_layouts/15/download.aspx?SourceUrl={quote(source_path, safe='/')}"
            )

    separator = "&" if parsed.query else "?"
    if "download=1" not in parsed.query.lower():
        candidates.append(f"{url}{separator}download=1")

    if url not in candidates:
        candidates.append(url)

    return candidates


def download_microsoft_workbook(shared_url: str) -> Path:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    last_error: Exception | None = None

    for candidate_url in microsoft_download_candidates(shared_url):
        try:
            request = Request(
                candidate_url,
                headers={
                    "User-Agent": "Mozilla/5.0 OncoNavega/1.0",
                    "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*",
                },
            )
            with urlopen(request, timeout=30) as response:
                content = response.read()
                content_type = response.headers.get("content-type", "")
        except Exception as exc:
            last_error = exc
            continue

        if len(content) < 200 or content[:2] != b"PK":
            last_error = ValueError(
                "O link respondeu uma página web em vez de um arquivo .xlsx. "
                "No Excel Online, use Compartilhar > Qualquer pessoa com o link pode exibir, "
                "ou crie um link de download."
            )
            if "html" in content_type.lower():
                continue
            continue

        target = DATA_DIR / UPLOADED_WORKBOOK_NAME
        target.write_bytes(content)
        set_app_state(MICROSOFT_WORKBOOK_URL_KEY, shared_url.strip())
        set_app_state(LAST_MICROSOFT_DOWNLOAD_KEY, datetime.now(APP_TIMEZONE).isoformat(timespec="seconds"))
        return target

    if last_error:
        raise last_error
    raise ValueError("Informe um link compartilhado do Excel Online, OneDrive ou SharePoint.")


def refresh_microsoft_workbook_if_configured() -> Path | None:
    workbook_url = get_app_state(MICROSOFT_WORKBOOK_URL_KEY)
    if not workbook_url:
        return None
    return download_microsoft_workbook(workbook_url)


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


def get_included_workbook_sheets(available_sheets: list[str]) -> list[str]:
    raw_value = get_app_state(INCLUDED_WORKBOOK_SHEETS_KEY)
    if not raw_value:
        return available_sheets
    try:
        selected = json.loads(raw_value)
    except json.JSONDecodeError:
        return available_sheets
    if not isinstance(selected, list):
        return available_sheets
    valid = [sheet for sheet in selected if sheet in available_sheets]
    return valid or available_sheets


def set_included_workbook_sheets(sheet_names: list[str]) -> None:
    set_app_state(INCLUDED_WORKBOOK_SHEETS_KEY, json.dumps(sheet_names, ensure_ascii=False))


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


def normalized_patient_key(doctor_name: str, patient_name: str) -> str:
    return f"{normalize_header(doctor_name)}::{normalize_header(patient_name)}"


def read_workbook_patient_payloads(workbook_path: Path, doctor_sheets: list[str]) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    for sheet_name in doctor_sheets:
        dataframe = pd.read_excel(workbook_path, sheet_name=sheet_name, header=None, dtype=object)
        values = dataframe.fillna("").values.tolist()
        headers, data_rows = detect_sheet_header_and_rows(values)
        if not headers or "patient_name" not in headers:
            continue

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
            if not payload["name"]:
                continue
            records.append(
                {
                    "sheet_name": sheet_name,
                    "row_number": row_number,
                    "row_data": row_data,
                    "payload": payload,
                    "cycle_dates": extract_cycle_dates(row_data),
                    "source_key": (sheet_name, row_number),
                    "patient_key": normalized_patient_key(sheet_name, str(payload["name"])),
                }
            )
    return records


def analyze_workbook_sync(workbook_path: Path | None = None, selected_sheets: list[str] | None = None) -> dict[str, object]:
    workbook_path = workbook_path or find_primary_workbook_file()
    if workbook_path is None:
        raise FileNotFoundError("Planilha principal .xlsx não encontrada na pasta do projeto.")

    available_sheets = get_workbook_doctor_sheet_titles(workbook_path)
    doctor_sheets = selected_sheets or get_included_workbook_sheets(available_sheets)
    workbook_records = read_workbook_patient_payloads(workbook_path, doctor_sheets)
    workbook_source_keys = {record["source_key"] for record in workbook_records}
    workbook_patient_keys = {str(record["patient_key"]) for record in workbook_records}

    conn = get_connection()
    existing_rows = conn.execute(
        """
        SELECT
            p.id,
            p.name,
            d.name AS doctor_name,
            p.active,
            p.source_sheet_name,
            p.source_row_number
        FROM patients p
        JOIN doctors d ON d.id = p.doctor_id
        """
    ).fetchall()
    conn.close()

    existing_source_keys = {
        (row["source_sheet_name"], int(row["source_row_number"]))
        for row in existing_rows
        if row["source_sheet_name"] is not None and row["source_row_number"] is not None
    }
    existing_patient_keys = {
        normalized_patient_key(row["doctor_name"], row["name"])
        for row in existing_rows
    }
    existing_sourced_rows = [
        row
        for row in existing_rows
        if row["source_sheet_name"] is not None and row["source_row_number"] is not None
    ]
    missing_from_workbook = [
        row
        for row in existing_sourced_rows
        if (row["source_sheet_name"], int(row["source_row_number"])) not in workbook_source_keys
        and normalized_patient_key(row["doctor_name"], row["name"]) not in workbook_patient_keys
    ]

    new_records = [
        record
        for record in workbook_records
        if record["source_key"] not in existing_source_keys
        and str(record["patient_key"]) not in existing_patient_keys
    ]
    matched_records = [record for record in workbook_records if record not in new_records]
    duplicates = int(len(workbook_records) - len(workbook_patient_keys))
    active_existing = sum(1 for row in existing_rows if int(row["active"] or 0) == 1)
    delta = len(workbook_records) - active_existing

    return {
        "workbook_path": str(workbook_path),
        "workbook_name": workbook_path.name,
        "available_sheets": available_sheets,
        "doctor_sheets": doctor_sheets,
        "source_total": len(workbook_records),
        "existing_active_total": active_existing,
        "new_total": len(new_records),
        "matched_total": len(matched_records),
        "missing_total": len(missing_from_workbook),
        "duplicates_total": duplicates,
        "delta": delta,
        "new_examples": [
            {"Médico": record["sheet_name"], "Paciente": record["payload"]["name"]}
            for record in new_records[:12]
        ],
        "missing_examples": [
            {"Médico": row["doctor_name"], "Paciente": row["name"]}
            for row in missing_from_workbook[:12]
        ],
    }


def sync_google_sheets_to_db() -> tuple[int, int]:
    workbook_path = find_primary_workbook_file()
    if workbook_path is None:
        raise FileNotFoundError("Planilha principal .xlsx não encontrada na pasta do projeto.")

    available_doctor_sheets = get_workbook_doctor_sheet_titles(workbook_path)
    doctor_sheets = get_included_workbook_sheets(available_doctor_sheets)
    workbook_records = read_workbook_patient_payloads(workbook_path, doctor_sheets)
    conn = get_connection()
    existing_patient_state = {
        (row["source_sheet_name"], int(row["source_row_number"])): {
            "active": row["active"],
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
                active,
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

    for record in workbook_records:
            sheet_name = str(record["sheet_name"])
            row_number = int(record["row_number"])
            row_data = record["row_data"]
            payload = dict(record["payload"])
            doctor_id = get_or_create_doctor_id_in_conn(conn, sheet_name)
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
                    "active",
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
                    active, insurance_name, prescription_status, prescription_requested_date,
                    authorization_status, authorization_submission_date,
                    authorization_valid_until, scheduling_status, scheduled_cycle_date,
                    next_cycle_alert_days, protocol_next_cycle_date, source_sheet_name, source_row_number
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                    int(payload.get("active", 1)),
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

            cycle_dates = record["cycle_dates"]
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


DOCTOR_COLOR_PALETTE = [
    "#0f766e",
    "#1d4ed8",
    "#b45309",
    "#be123c",
    "#7c3aed",
    "#047857",
    "#0e7490",
    "#a21caf",
    "#c2410c",
]

DOCTOR_COLOR_OVERRIDES = {
    "dr_rafael_schmerling": "#1d4ed8",
    "dra_carolina_kawamura": "#be123c",
    "dra_cynthia_lemos": "#0f766e",
    "dra_juliana_pimenta_e_buzaid": "#7c3aed",
    "dr_tiago_kenji": "#b45309",
    "dr_marcos_magalhaes_e_buzaid": "#0e7490",
}


def doctor_calendar_color(doctor_name: str) -> str:
    normalized = normalize_header(doctor_name)
    if normalized in DOCTOR_COLOR_OVERRIDES:
        return DOCTOR_COLOR_OVERRIDES[normalized]
    if not normalized:
        return DOCTOR_COLOR_PALETTE[0]
    color_index = sum(ord(char) for char in normalized) % len(DOCTOR_COLOR_PALETTE)
    return DOCTOR_COLOR_PALETTE[color_index]


def calendar_patient_url(patient_id: int, cycle_date: date) -> str:
    return "?" + urlencode(
        {
            "view": "patient_detail",
            "patient_id": str(patient_id),
            "cycle_date": cycle_date.strftime(DATE_FMT),
        }
    )


def render_calendar_day_html(cell_date: date, in_month: bool, group: pd.DataFrame | None) -> str:
    count = 0 if group is None else len(group)
    day_classes = "calendar-day" if in_month else "calendar-day muted"
    count_html = f'<span class="calendar-count">{count} infusao(oes)</span>' if count else ""
    event_blocks: list[str] = []

    if group is not None and not group.empty:
        sorted_group = group.sort_values(["doctor_name", "patient_name"])
        for doctor_name, doctor_group in sorted_group.groupby("doctor_name", sort=False):
            color = doctor_calendar_color(str(doctor_name))
            patient_links = []
            for _, event in doctor_group.iterrows():
                patient_name = html.escape(abbreviate_patient_name(str(event["patient_name"])))
                patient_url = html.escape(calendar_patient_url(int(event["patient_id"]), event["scheduled_date"]))
                patient_links.append(f'<a class="calendar-patient-link" href="{patient_url}">{patient_name}</a>')
            if patient_links:
                safe_doctor = html.escape(str(doctor_name))
                event_blocks.append(
                    f'<div class="calendar-doctor-block" style="--doctor-color:{color};">'
                    f'<div class="calendar-doctor-group">{safe_doctor}</div>'
                    f'{"".join(patient_links)}</div>'
                )

    return (
        f'<div class="{day_classes}">'
        f'<div class="calendar-date">{cell_date.day}</div>'
        f'{count_html}'
        f'<div class="calendar-events">{"".join(event_blocks)}</div>'
        f'</div>'
    )


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

    day_headers = "".join(f'<div class="calendar-head">{day_name}</div>' for day_name in ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"])
    day_cells = []
    for slot in range(total_slots):
        cell_date = month_start - timedelta(days=start_offset) + timedelta(days=slot)
        in_month = month_start <= cell_date <= month_end
        day_cells.append(render_calendar_day_html(cell_date, in_month, day_map.get(cell_date)))

    st.markdown(
        f'<div class="calendar-scroll"><div class="calendar-grid">{day_headers}{"".join(day_cells)}</div></div>',
        unsafe_allow_html=True,
    )

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
        "doctor_id",
        "diagnosis",
        "regimen",
        "cycle_interval_days",
        "last_chemo_date",
        "next_chemo_date",
        "protocol_next_cycle_date",
        "support_plan",
        "notes",
        "insurance_name",
        "prescription_status",
        "prescription_requested_date",
        "authorization_status",
        "authorization_submission_date",
        "authorization_valid_until",
        "scheduling_status",
        "scheduled_cycle_date",
        "next_cycle_alert_days",
        "active",
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


def render_commercial_tab(filtered_patients: pd.DataFrame, filtered_sessions: pd.DataFrame) -> None:
    active_patients = len(filtered_patients)
    monthly_sessions = 0
    if not filtered_sessions.empty:
        sessions_df = filtered_sessions.copy()
        sessions_df["scheduled_dt"] = pd.to_datetime(sessions_df["scheduled_date"], errors="coerce")
        today = pd.Timestamp(date.today())
        next_30_days = today + pd.Timedelta(days=30)
        monthly_sessions = int(sessions_df["scheduled_dt"].between(today, next_30_days).sum())

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_metric("Entrada", "Profissional", "Navegador ou coordenador sente a dor e puxa o uso.", "metric-a")
    with col2:
        render_metric("Carteira no filtro", str(active_patients), "Base atual usada para dimensionar proposta.", "metric-d")
    with col3:
        render_metric("Infusões em 30 dias", str(monthly_sessions), "Volume ajuda a calcular ganho operacional.", "metric-b")
    with col4:
        render_metric("Receita principal", "Clínica", "Contrato institucional com usuários, governança e ROI.", "metric-protocol")

    st.markdown("")
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Estratégia: profissional adota, clínica contrata</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="subtle">
            O profissional de navegação sente a dor todos os dias e pode começar usando o app para organizar
            a própria carteira. A clínica compra quando percebe que precisa centralizar os dados, liberar acesso
            para a equipe, padronizar o processo e proteger agenda, receita e governança.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("")
    product_page_html = f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{PRODUCT_NAME} - Navegação oncológica operacional</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; font-family: Arial, sans-serif; background: #f5fafb; color: #123847; line-height: 1.5; }}
    main {{ max-width: 1040px; margin: 0 auto; padding: 34px 24px 46px; }}
    .hero {{ background: linear-gradient(135deg, #0f3d4c 0%, #16697a 58%, #2a9d8f 100%); color: white; border-radius: 24px; padding: 36px; box-shadow: 0 22px 48px rgba(15,61,76,.16); }}
    .brand {{ display: flex; gap: 12px; align-items: center; margin-bottom: 18px; }}
    .mark {{ width: 46px; height: 46px; border-radius: 15px; background: rgba(255,255,255,.16); border: 1px solid rgba(255,255,255,.25); display: flex; align-items: center; justify-content: center; font-weight: 800; }}
    .kicker {{ color: #d9fbff; text-transform: uppercase; letter-spacing: .08em; font-size: 12px; font-weight: 800; }}
    h1 {{ margin: 0; font-size: 38px; line-height: 1.08; max-width: 780px; }}
    .hero p {{ max-width: 760px; color: #eefcff; }}
    .chips {{ display: flex; flex-wrap: wrap; gap: 10px; margin-top: 18px; }}
    .chip {{ border: 1px solid rgba(255,255,255,.25); background: rgba(255,255,255,.14); border-radius: 999px; padding: 8px 12px; font-weight: 700; }}
    .grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-top: 18px; }}
    section {{ background: white; border: 1px solid #dce9ee; border-radius: 18px; padding: 22px; margin-top: 18px; }}
    h2 {{ margin: 0 0 10px 0; color: #0f3d4c; }}
    .card h3 {{ margin: 0 0 8px 0; color: #123847; }}
    .price {{ font-size: 28px; font-weight: 800; color: #0f3d4c; margin: 8px 0; }}
    li {{ margin: 6px 0; }}
    @media (max-width: 820px) {{ .grid {{ grid-template-columns: 1fr; }} h1 {{ font-size: 30px; }} }}
  </style>
</head>
<body>
  <main>
    <div class="hero">
      <div class="brand">
        <div class="mark">ON</div>
        <div>
          <div class="kicker">{PRODUCT_NAME}</div>
          <div>{PRODUCT_PROMISE}</div>
        </div>
      </div>
      <h1>{PRODUCT_TAGLINE}</h1>
      <p>Produto para clínicas e profissionais de navegação que precisam transformar planilhas, mensagens e cobranças soltas em uma operação rastreável por ciclo.</p>
      <div class="chips">
        <div class="chip">Fila prioritária</div>
        <div class="chip">Agenda de infusão</div>
        <div class="chip">Autorização e prescrição</div>
        <div class="chip">Relatório para gestão</div>
      </div>
    </div>
    <section>
      <h2>Para quem é</h2>
      <div class="grid">
        <div class="card"><h3>Profissional navegador</h3><p>Organiza a carteira, prioriza cobranças e leva um resumo claro para a clínica.</p></div>
        <div class="card"><h3>Clínica oncológica</h3><p>Centraliza operação, reduz ciclos em risco e ganha governança sobre agenda e autorizações.</p></div>
        <div class="card"><h3>Gestão e faturamento</h3><p>Enxerga gargalos antes que virem atraso, retrabalho ou perda de previsibilidade.</p></div>
      </div>
    </section>
    <section>
      <h2>Oferta inicial</h2>
      <div class="grid">
        <div class="card"><h3>Profissional</h3><div class="price">R$ 99 a R$ 299/mês</div><p>Carteira individual, alertas e resumo para apresentar à clínica.</p></div>
        <div class="card"><h3>Clínica</h3><div class="price">R$ 1,5k a R$ 6k/mês</div><p>Multiusuário, relatórios, suporte, backup e governança.</p></div>
        <div class="card"><h3>Piloto de 45 dias</h3><div class="price">R$ 3k a R$ 8k</div><p>Implantação assistida, relatório semanal e reunião de fechamento.</p></div>
      </div>
    </section>
    <section>
      <h2>Como prova valor</h2>
      <ul>
        <li>Mostra pacientes com risco de atraso por prescrição, autorização ou agenda.</li>
        <li>Converte a carteira em fila operacional de ação diária.</li>
        <li>Ancora a mensalidade em ciclos protegidos e receita operacional preservada.</li>
      </ul>
    </section>
  </main>
</body>
</html>
"""
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Kit de produto</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="subtle">{PRODUCT_NAME} agora tem nome, promessa, oferta inicial e uma one-page comercial para abrir conversa com clínicas e profissionais.</div>',
        unsafe_allow_html=True,
    )
    st.download_button(
        "Baixar one-page do produto",
        data=product_page_html,
        file_name="onconavega_one_page.html",
        mime="text/html",
        use_container_width=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("")
    mode_left, mode_right = st.columns([0.9, 1.1])
    with mode_left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Modo de venda</div>', unsafe_allow_html=True)
        commercial_mode = st.radio(
            "Escolha o discurso principal",
            ["Profissional navegador", "Clínica / gestor"],
            horizontal=True,
            label_visibility="collapsed",
        )
        if commercial_mode == "Profissional navegador":
            mode_pitch = (
                "Organize sua carteira em um lugar só e gere um resumo objetivo para mostrar à clínica "
                "onde existem riscos de prescrição, autorização, agenda e próximo ciclo."
            )
            mode_cta = "Comece com o plano profissional e use a apresentação externa para abrir a conversa interna."
            mode_bullets = [
                "Menos dependência de planilhas soltas.",
                "Mais clareza sobre quem cobrar hoje.",
                "Material pronto para pedir apoio da clínica.",
            ]
        else:
            mode_pitch = (
                "Centralize a operação de navegação oncológica, reduza ciclos em risco e acompanhe a agenda "
                "com indicadores para gestão, médicos, navegação e faturamento."
            )
            mode_cta = "Contrate um piloto institucional de 45 dias com meta operacional e relatório semanal."
            mode_bullets = [
                "Multiusuário e governança dos dados.",
                "Fila prioritária por risco operacional.",
                "ROI ancorado em receita preservada.",
            ]
        st.markdown("</div>", unsafe_allow_html=True)
    with mode_right:
        bullet_html = "".join(f"<li>{item}</li>" for item in mode_bullets)
        st.markdown(
            f"""
            <div class="playbook-card">
                <h3>{commercial_mode}</h3>
                <p>{mode_pitch}</p>
                <ul>{bullet_html}</ul>
                <p><strong>Próxima ação:</strong> {mode_cta}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("")
    fit_cols = st.columns(3)
    commercial_cards = [
        (
            "Porta de entrada",
            [
                "Navegadora oncológica, enfermeira coordenadora ou secretária especializada.",
                "Profissional que acompanha carteiras em uma ou mais clínicas.",
                "Uso inicial com dados mínimos, autorizados ou anonimizados.",
            ],
        ),
        (
            "Momento de conversão",
            [
                "O profissional mostra uma fila real de pendências para a gestão.",
                "A clínica vê risco de atraso, autorização e agenda em um painel simples.",
                "O gestor entende que o problema é institucional, não individual.",
            ],
        ),
        (
            "Contrato que sustenta",
            [
                "Plano clínica com multiusuário, relatórios e governança.",
                "Implantação com a planilha atual e rotina da equipe.",
                "Mensalidade ancorada em receita preservada e menor retrabalho.",
            ],
        ),
    ]
    for column, (title, bullets) in zip(fit_cols, commercial_cards):
        with column:
            bullet_html = "".join(f"<li>{item}</li>" for item in bullets)
            st.markdown(
                f"""
                <div class="commercial-card">
                    <h3>{title}</h3>
                    <ul>{bullet_html}</ul>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("")
    price_cols = st.columns(3)
    price_cards = [
        (
            "Profissional",
            "R$ 99 a R$ 299",
            "por usuário/mês",
            ["Carteira individual", "Alertas e calendário", "Resumo para apresentar à clínica"],
        ),
        (
            "Clínica",
            "R$ 1,5k a R$ 6k",
            "por unidade/mês",
            ["Multiusuário e perfis", "Relatórios gerenciais", "Suporte, backup e governança"],
        ),
        (
            "Implantação",
            "R$ 15k+",
            "setup, treino e integrações",
            ["Carga de base e planilhas", "Desenho da rotina operacional", "Integração com agenda, BI ou ERP"],
        ),
    ]
    for column, (name, value, note, bullets) in zip(price_cols, price_cards):
        with column:
            bullet_html = "".join(f"<li>{item}</li>" for item in bullets)
            st.markdown(
                f"""
                <div class="price-card">
                    <div class="price-name">{name}</div>
                    <div class="price-value">{value}</div>
                    <div class="price-note">{note}</div>
                    <ul>{bullet_html}</ul>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("")
    funnel_left, funnel_right = st.columns([1, 1])
    with funnel_left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Convite interno para a clínica</div>', unsafe_allow_html=True)
        pending_patients = 0
        if not filtered_patients.empty:
            operational = build_operational_table(filtered_patients)
            pending_patients = int(
                (
                    (operational["Severidade operacional"] > 0)
                    | (operational["Severidade protocolo"] > 0)
                ).sum()
            )
        external_presentation = f"""# Apresentação Externa - {PRODUCT_NAME}

{PRODUCT_TAGLINE}

## Oportunidade identificada

Minha carteira acompanha {active_patients} paciente(s) ativo(s), com {monthly_sessions} infusão(ões) previstas nos próximos 30 dias.

Neste momento, {pending_patients} paciente(s) têm algum ponto de atenção operacional ou de protocolo, como prescrição, autorização, agendamento ou janela do próximo ciclo.

## Por que isso importa para a clínica

Quando esses pontos ficam espalhados em planilhas, mensagens e memória da equipe, a clínica corre risco de atraso de ciclo, retrabalho administrativo, pior experiência do paciente e perda de previsibilidade da agenda de infusão.

## Proposta

Usar o {PRODUCT_NAME} como ferramenta institucional para centralizar a carteira, acompanhar pendências por ciclo, priorizar pacientes em risco e dar visibilidade para navegação, médicos, faturamento e gestão.

## Próximo passo sugerido

Fazer um piloto institucional de 30 a 60 dias, com uma meta simples:

- reduzir ciclos com pendência perto da data de infusão
- antecipar cobranças de prescrição e autorização
- melhorar previsibilidade da agenda
- gerar um relatório semanal de gargalos operacionais

## Modelo comercial

- Plano clínica: R$ 1.500 a R$ 6.000 por unidade/mês
- Implantação assistida: R$ 3.000 a R$ 8.000 no piloto
- Integrações e expansão: orçamento conforme escopo
"""
        st.markdown(
            f"""
            <div class="subtle">
                Minha carteira tem {active_patients} paciente(s) ativos e {monthly_sessions} infusão(ões)
                previstas nos próximos 30 dias. Hoje existem {pending_patients} paciente(s) com algum ponto
                de atenção operacional ou de protocolo. Se a clínica contratar o plano institucional, a equipe
                passa a acompanhar isso com acesso compartilhado, rotina padronizada e relatórios para gestão.
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Apresentação Externa", use_container_width=True):
            st.session_state["external_presentation_text"] = external_presentation
        if st.session_state.get("external_presentation_text"):
            st.text_area(
                "Texto pronto para enviar",
                st.session_state["external_presentation_text"],
                height=360,
            )
            st.download_button(
                "Baixar apresentação externa",
                data=st.session_state["external_presentation_text"],
                file_name="apresentacao_externa_navegacao_oncologica.md",
                mime="text/markdown",
                use_container_width=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)
    with funnel_right:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Regras comerciais para dados sensíveis</div>', unsafe_allow_html=True)
        st.markdown(
            """
            <div class="subtle">
                Para uso profissional individual, trabalhe com autorização da clínica ou com dados mínimos:
                iniciais, datas, status e códigos internos. Dados identificáveis e acesso multiusuário devem
                migrar para o plano clínica, com responsável institucional, controle de acesso e rotina de backup.
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("")
    outreach_left, outreach_right = st.columns([1, 1])
    whatsapp_message = f"""Olá, tudo bem?

Estou organizando a navegação oncológica com o {PRODUCT_NAME}, um produto simples que mostra carteira, próximos ciclos, pendências de prescrição, autorização e agenda.

Na carteira analisada hoje temos {active_patients} paciente(s) ativo(s), {monthly_sessions} infusão(ões) nos próximos 30 dias e {pending_patients} paciente(s) com algum ponto de atenção operacional ou de protocolo.

Acho que vale uma conversa rápida para avaliar um piloto de 45 dias na clínica, com relatório semanal de gargalos e foco em reduzir risco de atraso de ciclo.
"""
    email_message = f"""Assunto: Piloto de navegação oncológica para reduzir atrasos de ciclo

Olá,

Estou estruturando o {PRODUCT_NAME} para dar visibilidade aos próximos ciclos, pendências de prescrição, autorização do convênio e agendamento de quimioterapia.

Na carteira analisada, temos {active_patients} paciente(s) ativo(s), {monthly_sessions} infusão(ões) previstas nos próximos 30 dias e {pending_patients} paciente(s) com algum ponto de atenção operacional ou de protocolo.

Minha sugestão é fazermos um piloto institucional de 45 dias, com:

- carga da planilha atual
- fila prioritária de pacientes em risco
- acompanhamento de prescrição, autorização e agenda
- relatório semanal de gargalos
- reunião de fechamento com indicadores e próximos passos

O objetivo é reduzir risco de atraso de ciclo, melhorar previsibilidade da agenda e dar mais governança para navegação, médicos, faturamento e gestão.
"""
    with outreach_left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Mensagem rápida para WhatsApp</div>', unsafe_allow_html=True)
        st.text_area("WhatsApp", whatsapp_message, height=210, label_visibility="collapsed")
        st.download_button(
            "Baixar mensagem WhatsApp",
            data=whatsapp_message,
            file_name="mensagem_whatsapp_navegacao_oncologica.txt",
            mime="text/plain",
            use_container_width=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
    with outreach_right:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">E-mail de abordagem</div>', unsafe_allow_html=True)
        st.text_area("E-mail", email_message, height=210, label_visibility="collapsed")
        st.download_button(
            "Baixar e-mail de abordagem",
            data=email_message,
            file_name="email_abordagem_navegacao_oncologica.txt",
            mime="text/plain",
            use_container_width=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("")
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Pacote de piloto de 45 dias</div>', unsafe_allow_html=True)
    pilot_cols = st.columns(3)
    pilot_steps = [
        ("Semana 1", "Carga da planilha, ajuste dos campos e definição dos indicadores do piloto."),
        ("Semanas 2-3", "Uso assistido da fila prioritária, calendário e pendências por ciclo."),
        ("Semanas 4-5", "Relatórios semanais de gargalos, atrasos evitáveis e oportunidades de processo."),
        ("Semana 6", "Reunião de fechamento com ROI estimado, decisão de assinatura e próximos incrementos."),
        ("Entregáveis", "Painel em uso, rotina documentada, proposta de contrato mensal e backlog de melhorias."),
        ("Meta do piloto", "Reduzir ciclos em risco e provar valor antes da contratação mensal."),
    ]
    for index, (title, copy) in enumerate(pilot_steps):
        with pilot_cols[index % 3]:
            st.markdown(
                f"""
                <div class="pilot-week">
                    <div class="pilot-week-title">{title}</div>
                    <div class="pilot-week-copy">{copy}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("")
    calc_left, calc_right = st.columns([1, 1.2])
    with calc_left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Simulador de ROI para venda</div>', unsafe_allow_html=True)
        patients_input = st.number_input("Pacientes ativos na clínica", min_value=10, max_value=2000, value=max(active_patients, 120), step=10)
        monthly_ticket = st.number_input("Receita média por ciclo/infusão (R$)", min_value=500, max_value=50000, value=4500, step=500)
        avoided_delay_rate = st.slider("Ciclos protegidos por mês (%)", min_value=1, max_value=20, value=5)
        monthly_price = st.number_input("Mensalidade clínica proposta (R$)", min_value=500, max_value=30000, value=3500, step=500)
        protected_cycles = max(1, round(patients_input * avoided_delay_rate / 100))
        protected_revenue = protected_cycles * monthly_ticket
        roi_multiple = protected_revenue / monthly_price if monthly_price else 0
        st.markdown("</div>", unsafe_allow_html=True)
    with calc_right:
        roi_cols = st.columns(3)
        with roi_cols[0]:
            render_metric("Ciclos protegidos", str(protected_cycles), "Estimativa mensal para defender valor.", "metric-d")
        with roi_cols[1]:
            render_metric("Receita preservada", f"R$ {protected_revenue:,.0f}".replace(",", "."), "Valor de agenda que o painel ajuda a proteger.", "metric-b")
        with roi_cols[2]:
            render_metric("ROI potencial", f"{roi_multiple:.1f}x", "Relação entre valor protegido e mensalidade.", "metric-a")
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Frase de venda</div>', unsafe_allow_html=True)
        st.markdown(
            f"""
            <div class="subtle">
                Se o profissional mostrar que a clínica pode proteger apenas {protected_cycles} ciclo(s) por mês, com ticket médio de
                R$ {monthly_ticket:,.0f}, o sistema ajuda a defender cerca de
                R$ {protected_revenue:,.0f} em receita operacional. Uma mensalidade de
                R$ {monthly_price:,.0f} para o plano clínica fica ancorada em ROI potencial de {roi_multiple:.1f}x.
            </div>
            """.replace(",", "."),
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("")
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Gerar proposta comercial</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtle">Monte uma proposta objetiva para enviar ao gestor da clínica após a conversa inicial.</div>',
        unsafe_allow_html=True,
    )
    with st.form("commercial_proposal_form"):
        proposal_col1, proposal_col2, proposal_col3 = st.columns(3)
        with proposal_col1:
            client_name = st.text_input("Nome da clínica", value="Clínica de Oncologia")
            contact_name = st.text_input("Contato ou gestor", value="Gestão da clínica")
        with proposal_col2:
            proposal_plan = st.selectbox(
                "Plano proposto",
                ["Piloto institucional", "Plano clínica mensal", "Implantação enterprise"],
            )
            pilot_days = st.number_input("Duração do piloto (dias)", min_value=15, max_value=120, value=45, step=15)
        with proposal_col3:
            setup_value = st.number_input("Implantação / piloto (R$)", min_value=0, max_value=100000, value=5000, step=500)
            proposal_monthly_value = st.number_input("Mensalidade após piloto (R$)", min_value=0, max_value=50000, value=monthly_price, step=500)
        proposal_notes = st.text_area(
            "Escopo e observações",
            value="Carga da planilha atual, treinamento da equipe, painel operacional, relatório semanal de gargalos e reunião de fechamento do piloto.",
            height=90,
        )
        generate_proposal = st.form_submit_button("Gerar Proposta Comercial", use_container_width=True)

    if generate_proposal:
        setup_label = f"R$ {setup_value:,.0f}".replace(",", ".")
        monthly_label = f"R$ {proposal_monthly_value:,.0f}".replace(",", ".")
        revenue_label = f"R$ {protected_revenue:,.0f}".replace(",", ".")
        monthly_ticket_label = f"R$ {monthly_ticket:,.0f}".replace(",", ".")
        proposal_text = f"""# Proposta Comercial - {PRODUCT_NAME}

{PRODUCT_TAGLINE}

## Cliente

{client_name}

Contato: {contact_name}

## Resumo executivo

Propomos um {proposal_plan.lower()} com o {PRODUCT_NAME} para organizar a navegação oncológica da clínica, centralizando carteira de pacientes, próximos ciclos, prescrição, autorização, agendamento e alertas operacionais.

A análise inicial indica {active_patients} paciente(s) na carteira filtrada, {monthly_sessions} infusão(ões) nos próximos 30 dias e {pending_patients} paciente(s) com algum ponto de atenção operacional ou de protocolo.

## Objetivo do piloto

Durante {pilot_days} dias, a meta será reduzir riscos de atraso de ciclo e aumentar previsibilidade da agenda de infusão, com foco em:

- antecipar cobrança de prescrição
- acompanhar autorização do convênio
- identificar pacientes sem agenda confirmada
- consolidar a fila prioritária da equipe
- gerar relatório semanal de gargalos

## Tese de retorno

Se a clínica proteger {protected_cycles} ciclo(s) por mês, com receita média estimada de {monthly_ticket_label} por ciclo, o valor operacional preservado pode chegar a {revenue_label}/mês.

Com mensalidade de {monthly_label}, o ROI potencial estimado é de {roi_multiple:.1f}x.

## Investimento

- Implantação / piloto: {setup_label}
- Mensalidade após piloto: {monthly_label}
- Duração inicial: {pilot_days} dias

## Escopo incluído

{proposal_notes}

## Próximo passo

Realizar uma reunião de alinhamento operacional, validar a planilha-base e definir os indicadores de acompanhamento do piloto.
"""
        proposal_html = f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Proposta Comercial - {PRODUCT_NAME}</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{ font-family: Arial, sans-serif; margin: 0; background: #f4f8fa; color: #123847; line-height: 1.5; }}
    main {{ max-width: 980px; margin: 0 auto; padding: 36px 24px 48px; }}
    header {{
      background: linear-gradient(135deg, #0f3d4c 0%, #16697a 58%, #2a9d8f 100%);
      color: white;
      padding: 34px;
      border-radius: 22px;
      box-shadow: 0 22px 45px rgba(15, 61, 76, 0.16);
    }}
    .brand {{ display: flex; align-items: center; gap: 12px; margin-bottom: 18px; }}
    .mark {{
      width: 44px;
      height: 44px;
      border-radius: 14px;
      background: rgba(255,255,255,.16);
      border: 1px solid rgba(255,255,255,.25);
      display: flex;
      align-items: center;
      justify-content: center;
      font-weight: 800;
    }}
    .kicker {{ font-size: 12px; text-transform: uppercase; letter-spacing: .08em; color: #d9fbff; font-weight: 800; }}
    h1 {{ margin: 0 0 8px 0; font-size: 34px; line-height: 1.1; }}
    h2 {{ margin: 0 0 12px 0; color: #0f3d4c; }}
    section {{ background: white; border: 1px solid #dce9ee; border-radius: 18px; padding: 22px; margin-top: 18px; }}
    .grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 16px; }}
    .metric {{ background: #eef8fc; border-radius: 14px; padding: 16px; min-height: 98px; }}
    .metric strong {{ display: block; font-size: 27px; color: #0f3d4c; margin-bottom: 4px; }}
    .investment {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }}
    .investment div {{ background: #f7fbfd; border: 1px solid #dce9ee; border-radius: 14px; padding: 14px; }}
    .investment strong {{ display: block; color: #0f3d4c; }}
    li {{ margin: 6px 0; }}
    footer {{ color: #67808a; font-size: 13px; margin-top: 18px; text-align: center; }}
    @media (max-width: 760px) {{ .grid, .investment {{ grid-template-columns: 1fr; }} h1 {{ font-size: 28px; }} }}
  </style>
</head>
<body>
  <main>
    <header>
      <div class="brand">
        <div class="mark">ON</div>
        <div>
          <div class="kicker">{PRODUCT_NAME}</div>
          <div>{PRODUCT_PROMISE}</div>
        </div>
      </div>
      <h1>{PRODUCT_TAGLINE}</h1>
      <div>Proposta comercial para {client_name}</div>
    </header>
    <section>
      <h2>Resumo executivo</h2>
      <p>Propomos um {proposal_plan.lower()} com o {PRODUCT_NAME} para organizar carteira, próximos ciclos, prescrição, autorização, agendamento e alertas operacionais.</p>
      <div class="grid">
        <div class="metric"><strong>{active_patients}</strong>pacientes na carteira</div>
        <div class="metric"><strong>{monthly_sessions}</strong>infusões em 30 dias</div>
        <div class="metric"><strong>{pending_patients}</strong>pacientes com atenção</div>
      </div>
    </section>
    <section>
      <h2>Tese de retorno</h2>
      <p>Protegendo {protected_cycles} ciclo(s) por mês, com ticket médio estimado de {monthly_ticket_label}, a clínica pode preservar cerca de {revenue_label}/mês em operação.</p>
      <p>Mensalidade proposta: <strong>{monthly_label}</strong>. ROI potencial estimado: <strong>{roi_multiple:.1f}x</strong>.</p>
    </section>
    <section>
      <h2>Investimento e escopo</h2>
      <div class="investment">
        <div><strong>{setup_label}</strong>Implantação / piloto</div>
        <div><strong>{monthly_label}</strong>Mensalidade após piloto</div>
        <div><strong>{pilot_days} dias</strong>Duração inicial</div>
      </div>
      <p>{proposal_notes}</p>
    </section>
    <footer>{PRODUCT_NAME} - Navegação oncológica operacional</footer>
  </main>
</body>
</html>
"""
        st.session_state["commercial_proposal_text"] = proposal_text
        st.session_state["commercial_proposal_html"] = proposal_html

    if st.session_state.get("commercial_proposal_text"):
        st.text_area(
            "Proposta pronta para enviar",
            st.session_state["commercial_proposal_text"],
            height=420,
        )
        download_col1, download_col2 = st.columns(2)
        with download_col1:
            st.download_button(
                "Baixar proposta em Markdown",
                data=st.session_state["commercial_proposal_text"],
                file_name="proposta_comercial_navegacao_oncologica.md",
                mime="text/markdown",
                use_container_width=True,
            )
        with download_col2:
            st.download_button(
                "Baixar proposta visual em HTML",
                data=st.session_state["commercial_proposal_html"],
                file_name="proposta_comercial_navegacao_oncologica.html",
                mime="text/html",
                use_container_width=True,
            )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("")
    sales_left, sales_right = st.columns([1, 1])
    with sales_left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Roteiro de venda em 6 passos</div>', unsafe_allow_html=True)
        steps = [
            ("1. Profissional testa", "Navegador ou coordenador organiza uma carteira e sente alívio operacional."),
            ("2. Relatório de valor", "O app mostra pendências, ciclos próximos e riscos que a clínica deveria enxergar."),
            ("3. Convite interno", "O profissional apresenta a visão para gestor, médico líder ou faturamento."),
            ("4. Piloto institucional", "A clínica usa por 30 a 60 dias com meta operacional objetiva."),
            ("5. Contrato mensal", "A assinatura entra com multiusuário, governança, suporte e relatórios."),
            ("6. Expansão", "Adicionar unidades, integrações, indicadores financeiros e treinamento recorrente."),
        ]
        for title, copy in steps:
            st.markdown(
                f"""
                <div class="sales-step">
                    <div class="sales-step-title">{title}</div>
                    <div class="sales-step-copy">{copy}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)
    with sales_right:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Próximos incrementos que aumentam preço</div>', unsafe_allow_html=True)
        roadmap = pd.DataFrame(
            [
                ["Curto prazo", "Exportar fila prioritária", "Facilita rotina da equipe e vira entrega semanal."],
                ["Curto prazo", "Resumo para apresentar à clínica", "Transforma o profissional em canal de venda."],
                ["Curto prazo", "Alertas por e-mail/WhatsApp", "Aumenta percepção de automação e reduz dependência manual."],
                ["Médio prazo", "Multiusuário com perfis", "Permite vender para clínicas maiores com governança."],
                ["Médio prazo", "Indicadores financeiros", "Conecta cuidado, agenda e receita protegida."],
                ["Longo prazo", "Integração com Tasy/PEP/ERP", "Abre venda enterprise e contratos mais altos."],
            ],
            columns=["Prazo", "Entrega", "Por que monetiza"],
        )
        st.dataframe(roadmap, use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)


def render_patient_management_tab(doctors_df: pd.DataFrame, patients_df: pd.DataFrame) -> None:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Gestão de pacientes</div>', unsafe_allow_html=True)

    if doctors_df.empty:
        with st.form("patient_management_doctor_form", clear_on_submit=True):
            st.markdown("**Cadastrar médico**")
            doctor_name = st.text_input("Nome do médico")
            specialty = st.text_input("Especialidade")
            submitted = st.form_submit_button("Salvar médico", use_container_width=True)
            if submitted:
                if not doctor_name.strip():
                    st.warning("Informe o nome do médico.")
                else:
                    insert_doctor(doctor_name, specialty)
                    st.success("Médico cadastrado.")
                    st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
        return

    active_patients = patients_df[patients_df["active"].fillna(1).astype(int) == 1].copy()
    archived_patients = patients_df[patients_df["active"].fillna(1).astype(int) == 0].copy()
    new_col, edit_col = st.columns([1, 1.1])

    doctor_options = {
        f'{row["name"]} | {row["specialty"] or "Sem especialidade"}': int(row["id"])
        for _, row in doctors_df.iterrows()
    }

    with new_col:
        with st.form("patient_management_create_form", clear_on_submit=True):
            st.markdown("**Novo paciente**")
            selected_doctor = st.selectbox("Médico responsável", list(doctor_options.keys()), key="create_patient_doctor")
            patient_name = st.text_input("Nome do paciente", key="create_patient_name")
            diagnosis = st.text_input("Diagnóstico", key="create_patient_diagnosis")
            regimen = st.text_input("Protocolo", key="create_patient_regimen")
            insurance_name = st.text_input("Convênio", key="create_patient_insurance")
            cycle_interval_days = st.number_input("Intervalo entre ciclos", min_value=1, max_value=90, value=21, key="create_patient_interval")
            next_cycle_alert_days = st.number_input("Janela de alerta", min_value=1, max_value=90, value=21, key="create_patient_alert")
            last_chemo_date = st.date_input("Última quimioterapia", value=None, key="create_patient_last_chemo")
            next_chemo_date = st.date_input("Próxima quimioterapia", value=None, key="create_patient_next_chemo")
            protocol_next_cycle = st.date_input("Próxima data do protocolo", value=None, key="create_patient_protocol_date")
            support_plan = st.text_input("Plano de suporte", key="create_patient_support")
            notes = st.text_area("Observações", key="create_patient_notes")
            prescription_status = st.selectbox("Prescrição", list(PRESCRIPTION_LABELS.keys()), format_func=lambda x: PRESCRIPTION_LABELS[x], key="create_patient_prescription")
            authorization_status = st.selectbox("Autorização", list(AUTHORIZATION_LABELS.keys()), format_func=lambda x: AUTHORIZATION_LABELS[x], key="create_patient_authorization")
            scheduling_status = st.selectbox("Agenda", list(SCHEDULING_LABELS.keys()), format_func=lambda x: SCHEDULING_LABELS[x], key="create_patient_scheduling")
            scheduled_cycle_date = st.date_input("Data agendada", value=None, key="create_patient_scheduled")
            submitted = st.form_submit_button("Salvar paciente", use_container_width=True)
            if submitted:
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
                            "prescription_requested_date": None,
                            "authorization_status": authorization_status,
                            "authorization_submission_date": None,
                            "authorization_valid_until": None,
                            "scheduling_status": scheduling_status,
                            "scheduled_cycle_date": scheduled_cycle_date.strftime(DATE_FMT) if scheduled_cycle_date else None,
                            "next_cycle_alert_days": int(next_cycle_alert_days),
                        }
                    )
                    new_patient_id = load_patients().sort_values("id").iloc[-1]["id"]
                    if protocol_next_cycle:
                        update_patient_record(int(new_patient_id), {"protocol_next_cycle_date": protocol_next_cycle.strftime(DATE_FMT)})
                    st.success("Paciente cadastrado.")
                    st.rerun()

    with edit_col:
        patient_pool = pd.concat([active_patients, archived_patients], ignore_index=True)
        if patient_pool.empty:
            st.info("Nenhum paciente cadastrado.")
        else:
            patient_options = {
                f'{row["name"]} | {row["doctor_name"]}{" | arquivado" if int(row["active"] or 0) == 0 else ""}': int(row["id"])
                for _, row in patient_pool.sort_values(["active", "name"], ascending=[False, True]).iterrows()
            }
            selected_label = st.selectbox("Paciente", list(patient_options.keys()), key="manage_patient_selector")
            selected_id = patient_options[selected_label]
            patient_row = patient_pool[patient_pool["id"] == selected_id].iloc[0]
            current_doctor_label = next(
                (
                    label
                    for label, doctor_id in doctor_options.items()
                    if doctor_id == int(patient_row["doctor_id"])
                ),
                list(doctor_options.keys())[0],
            )

            with st.form("patient_management_edit_form"):
                st.markdown("**Editar paciente**")
                prescription_keys = list(PRESCRIPTION_LABELS.keys())
                prescription_value = patient_row["prescription_status"] if patient_row["prescription_status"] in prescription_keys else "not_requested"
                authorization_keys = list(AUTHORIZATION_LABELS.keys())
                authorization_value = patient_row["authorization_status"] if patient_row["authorization_status"] in authorization_keys else "not_sent"
                scheduling_keys = list(SCHEDULING_LABELS.keys())
                scheduling_value = patient_row["scheduling_status"] if patient_row["scheduling_status"] in scheduling_keys else "not_booked"
                selected_doctor = st.selectbox(
                    "Médico responsável",
                    list(doctor_options.keys()),
                    index=list(doctor_options.keys()).index(current_doctor_label),
                    key="edit_patient_doctor",
                )
                diagnosis = st.text_input("Diagnóstico", value=patient_row["diagnosis"] or "", key="edit_patient_diagnosis")
                regimen = st.text_input("Protocolo", value=patient_row["regimen"] or "", key="edit_patient_regimen")
                insurance_name = st.text_input("Convênio", value=patient_row["insurance_name"] or "", key="edit_patient_insurance")
                cycle_interval_days = st.number_input(
                    "Intervalo entre ciclos",
                    min_value=1,
                    max_value=90,
                    value=int(patient_row["cycle_interval_days"] or 21),
                    key="edit_patient_interval",
                )
                next_cycle_alert_days = st.number_input(
                    "Janela de alerta",
                    min_value=1,
                    max_value=90,
                    value=int(patient_row["next_cycle_alert_days"] or 21),
                    key="edit_patient_alert",
                )
                last_chemo_date = st.date_input("Última quimioterapia", value=parse_date(patient_row["last_chemo_date"]), key="edit_patient_last_chemo")
                next_chemo_date = st.date_input("Próxima quimioterapia", value=parse_date(patient_row["next_chemo_date"]), key="edit_patient_next_chemo")
                protocol_next_cycle = st.date_input("Próxima data do protocolo", value=parse_date(patient_row["protocol_next_cycle_date"]), key="edit_patient_protocol_date")
                scheduled_cycle_date = st.date_input("Data agendada", value=parse_date(patient_row["scheduled_cycle_date"]), key="edit_patient_scheduled")
                support_plan = st.text_input("Plano de suporte", value=patient_row["support_plan"] or "", key="edit_patient_support")
                notes = st.text_area("Observações", value=patient_row["notes"] or "", key="edit_patient_notes")
                prescription_status = st.selectbox(
                    "Prescrição",
                    prescription_keys,
                    index=prescription_keys.index(prescription_value),
                    format_func=lambda x: PRESCRIPTION_LABELS[x],
                    key="edit_patient_prescription",
                )
                authorization_status = st.selectbox(
                    "Autorização",
                    authorization_keys,
                    index=authorization_keys.index(authorization_value),
                    format_func=lambda x: AUTHORIZATION_LABELS[x],
                    key="edit_patient_authorization",
                )
                scheduling_status = st.selectbox(
                    "Agenda",
                    scheduling_keys,
                    index=scheduling_keys.index(scheduling_value),
                    format_func=lambda x: SCHEDULING_LABELS[x],
                    key="edit_patient_scheduling",
                )
                active_value = st.checkbox("Paciente ativo", value=int(patient_row["active"] or 0) == 1, key="edit_patient_active")
                submitted = st.form_submit_button("Salvar alterações", use_container_width=True)
                if submitted:
                    update_patient_record(
                        selected_id,
                        {
                            "doctor_id": doctor_options[selected_doctor],
                            "diagnosis": diagnosis.strip(),
                            "regimen": regimen.strip(),
                            "cycle_interval_days": int(cycle_interval_days),
                            "last_chemo_date": last_chemo_date.strftime(DATE_FMT) if last_chemo_date else None,
                            "next_chemo_date": next_chemo_date.strftime(DATE_FMT) if next_chemo_date else None,
                            "protocol_next_cycle_date": protocol_next_cycle.strftime(DATE_FMT) if protocol_next_cycle else None,
                            "scheduled_cycle_date": scheduled_cycle_date.strftime(DATE_FMT) if scheduled_cycle_date else None,
                            "support_plan": support_plan.strip(),
                            "notes": notes.strip(),
                            "insurance_name": insurance_name.strip(),
                            "prescription_status": prescription_status,
                            "authorization_status": authorization_status,
                            "scheduling_status": scheduling_status,
                            "next_cycle_alert_days": int(next_cycle_alert_days),
                            "active": 1 if active_value else 0,
                        },
                    )
                    st.success("Paciente atualizado.")
                    st.rerun()

    st.markdown("---")
    metric_col1, metric_col2 = st.columns(2)
    with metric_col1:
        render_metric("Pacientes ativos", str(len(active_patients)), "Entram no painel e no calendário.", "metric-a")
    with metric_col2:
        render_metric("Pacientes arquivados", str(len(archived_patients)), "Ficam fora da rotina diária.", "metric-d")
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


def render_sync_preview(preview: dict[str, object]) -> None:
    st.markdown("**Prévia da sincronização**")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_metric("Na planilha", str(preview["source_total"]), "Registros lidos nas abas escolhidas.", "metric-a")
    with col2:
        render_metric("Base atual", str(preview["existing_active_total"]), "Pacientes ativos antes de aplicar.", "metric-b")
    with col3:
        render_metric("Novos", str(preview["new_total"]), "Entrariam como novos registros.", "metric-c")
    with col4:
        render_metric("Não encontrados", str(preview["missing_total"]), "Estão no app, mas não apareceram na planilha.", "metric-d")

    delta = int(preview["delta"])
    duplicates = int(preview["duplicates_total"])
    if abs(delta) >= 5 or int(preview["missing_total"]) > 0:
        st.warning(
            "Atenção: a planilha tem diferença relevante em relação à base atual. "
            "Confira os exemplos abaixo antes de confirmar."
        )
    elif duplicates:
        st.info("A planilha tem pacientes repetidos. Isso pode ser correto quando há mais de um ciclo/registro.")
    else:
        st.success("A prévia não encontrou queda relevante de pacientes.")

    if preview["new_examples"]:
        st.write("**Exemplos de novos registros**")
        st.dataframe(pd.DataFrame(preview["new_examples"]), use_container_width=True, hide_index=True)
    if preview["missing_examples"]:
        st.write("**Exemplos que não aparecem na nova planilha**")
        st.dataframe(pd.DataFrame(preview["missing_examples"]), use_container_width=True, hide_index=True)


def render_google_sync_tab(patients_df: pd.DataFrame) -> None:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Planilha principal</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtle">Sincronize as abas dos médicos da planilha principal e mantenha o painel alinhado ao arquivo fonte.</div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns([1, 1.3])
    with col1:
        if st.button("Analisar planilha principal", use_container_width=True):
            try:
                refresh_microsoft_workbook_if_configured()
                st.session_state["primary_sync_preview"] = analyze_workbook_sync()
            except Exception as exc:
                st.error(f"Não consegui analisar a planilha principal. Detalhe: {exc}")
        st.caption("Primeiro analise. O app só sincroniza depois da confirmação.")
        if st.session_state.get("primary_sync_preview"):
            render_sync_preview(st.session_state["primary_sync_preview"])
            if st.button("Confirmar sincronização analisada", use_container_width=True, type="primary"):
                try:
                    imported, updated = sync_google_sheets_to_db()
                    st.session_state.pop("primary_sync_preview", None)
                    st.success(f"Sincronização concluída: {imported} registro(s) importado(s).")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Não consegui sincronizar com a planilha principal. Detalhe: {exc}")

    with col2:
        workbook_file = find_primary_workbook_file()
        last_sync = get_app_state("last_google_sync_at")
        st.write("**Arquivo fonte**")
        st.write(f"Planilha: `{workbook_file.name if workbook_file else 'não encontrada'}`")
        if workbook_file is not None:
            st.write(f"Caminho: `{workbook_file}`")
        st.write(f"Última sincronização: `{format_sync_timestamp(last_sync)}`")
        microsoft_url = get_app_state(MICROSOFT_WORKBOOK_URL_KEY)
        last_microsoft_download = get_app_state(LAST_MICROSOFT_DOWNLOAD_KEY)
        if microsoft_url:
            st.write("**Fonte online Microsoft:** vinculada")
            st.write(f"Último download online: `{format_sync_timestamp(last_microsoft_download)}`")

    st.markdown("---")
    st.markdown("**Abas sincronizadas**")
    st.caption(
        "Escolha quais abas de médicos entram no painel. Isso evita importar carteiras de outros médicos sem querer."
    )
    workbook_file_for_sheets = find_primary_workbook_file()
    if workbook_file_for_sheets is None:
        st.info("Nenhuma planilha principal encontrada para listar abas.")
    else:
        try:
            available_sheets = get_workbook_doctor_sheet_titles(workbook_file_for_sheets)
            included_sheets = get_included_workbook_sheets(available_sheets)
            selected_sheets = st.multiselect(
                "Abas de médicos para sincronizar",
                available_sheets,
                default=included_sheets,
            )
            if st.button("Analisar abas selecionadas", use_container_width=True):
                if not selected_sheets:
                    st.warning("Selecione pelo menos uma aba.")
                else:
                    st.session_state["sheet_sync_selection"] = selected_sheets
                    st.session_state["sheet_sync_preview"] = analyze_workbook_sync(
                        workbook_file_for_sheets,
                        selected_sheets,
                    )
            if st.session_state.get("sheet_sync_preview"):
                render_sync_preview(st.session_state["sheet_sync_preview"])
                if st.button("Confirmar abas e sincronizar", use_container_width=True, type="primary"):
                    try:
                        selected_for_sync = st.session_state.get("sheet_sync_selection", selected_sheets)
                        set_included_workbook_sheets(selected_for_sync)
                        imported, updated = sync_google_sheets_to_db()
                        st.session_state.pop("sheet_sync_preview", None)
                        st.session_state.pop("sheet_sync_selection", None)
                        st.success(
                            f"Abas salvas e base sincronizada com {len(selected_for_sync)} aba(s): {imported} registro(s)."
                        )
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Não consegui confirmar a sincronização. Detalhe: {exc}")
        except Exception as exc:
            st.error(f"Não consegui listar as abas da planilha. Detalhe: {exc}")

    st.markdown("---")
    st.markdown("**Vincular planilha online Microsoft**")
    st.caption(
        "Cole um link compartilhado do Excel Online, OneDrive ou SharePoint. O app baixa uma cópia .xlsx e sincroniza os dados."
    )
    saved_microsoft_url = get_app_state(MICROSOFT_WORKBOOK_URL_KEY) or ""
    with st.form("microsoft_workbook_form"):
        microsoft_workbook_url = st.text_input(
            "Link compartilhado da planilha Microsoft",
            value=saved_microsoft_url,
            placeholder="https://...onedrive... ou https://...sharepoint...",
        )
        microsoft_submit = st.form_submit_button("Vincular e sincronizar planilha online", use_container_width=True)
    if microsoft_submit:
        try:
            downloaded_path = download_microsoft_workbook(microsoft_workbook_url)
            imported, updated = sync_google_sheets_to_db()
            st.success(
                f"Planilha online vinculada em `{downloaded_path.name}` e sincronizada: {imported} novo(s) e {updated} atualizado(s)."
            )
            st.rerun()
        except Exception as exc:
            st.error(f"Não consegui baixar/sincronizar a planilha Microsoft. Detalhe: {exc}")

    if saved_microsoft_url:
        if st.button("Atualizar agora a partir do link Microsoft", use_container_width=True):
            try:
                downloaded_path = refresh_microsoft_workbook_if_configured()
                imported, updated = sync_google_sheets_to_db()
                st.success(
                    f"Fonte online atualizada em `{downloaded_path.name if downloaded_path else UPLOADED_WORKBOOK_NAME}`: {imported} novo(s) e {updated} atualizado(s)."
                )
                st.rerun()
            except Exception as exc:
                st.error(f"Não consegui atualizar pelo link Microsoft. Detalhe: {exc}")

    st.markdown("---")
    st.markdown("**Usar arquivo sincronizado no Mac**")
    st.caption(
        "Se o SharePoint exigir login, sincronize a pasta pelo OneDrive no Finder e cole aqui o caminho local do arquivo .xlsx."
    )
    saved_local_path = get_app_state(LOCAL_WORKBOOK_PATH_KEY) or ""
    with st.form("local_workbook_path_form"):
        local_workbook_path = st.text_input(
            "Caminho local da planilha sincronizada",
            value=saved_local_path,
            placeholder="/Users/seu-usuario/Library/CloudStorage/OneDrive-.../PLANILHA.xlsx",
        )
        local_submit = st.form_submit_button("Analisar caminho local", use_container_width=True)
    if local_submit:
        try:
            candidate_path = clean_local_workbook_path(local_workbook_path)
            if candidate_path.suffix.lower() == ".url":
                raise ValueError(
                    "Esse caminho aponta para um atalho .url do OneDrive, não para a planilha. "
                    "No Finder, escolha o arquivo com ícone XLSX e extensão .xlsx, não o atalho .url."
                )
            if not candidate_path.exists():
                raise FileNotFoundError("Não encontrei esse arquivo no Mac. Copie o caminho completo pelo Finder.")
            if candidate_path.suffix.lower() != ".xlsx":
                raise ValueError("O arquivo precisa estar no formato .xlsx.")
            set_app_state(LOCAL_WORKBOOK_PATH_KEY, str(candidate_path))
            st.session_state["local_workbook_preview"] = analyze_workbook_sync(candidate_path)
            st.success("Caminho local salvo. Confira a prévia antes de sincronizar.")
        except Exception as exc:
            st.error(f"Não consegui usar esse caminho local. Detalhe: {exc}")
    if st.session_state.get("local_workbook_preview"):
        render_sync_preview(st.session_state["local_workbook_preview"])
        if st.button("Confirmar caminho local e sincronizar", use_container_width=True, type="primary"):
            try:
                imported, updated = sync_google_sheets_to_db()
                st.session_state.pop("local_workbook_preview", None)
                st.success(f"Caminho local sincronizado: {imported} registro(s).")
                st.rerun()
            except Exception as exc:
                st.error(f"Não consegui sincronizar o caminho local. Detalhe: {exc}")

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
        if st.button("Analisar esta planilha antes de trocar", use_container_width=True, key="analyze_primary_workbook"):
            try:
                pending_path = save_pending_primary_workbook(uploaded_workbook)
                st.session_state["pending_workbook_preview"] = analyze_workbook_sync(pending_path)
                st.session_state["pending_workbook_name"] = uploaded_workbook.name
            except Exception as exc:
                st.error(f"Não consegui analisar a nova planilha. Detalhe: {exc}")
    if st.session_state.get("pending_workbook_preview"):
        st.info(f"Arquivo pendente: `{st.session_state.get('pending_workbook_name', PENDING_WORKBOOK_NAME)}`")
        render_sync_preview(st.session_state["pending_workbook_preview"])
        if st.button("Confirmar troca da fonte e sincronizar", use_container_width=True, type="primary"):
            try:
                saved_path = promote_pending_primary_workbook()
                imported, updated = sync_google_sheets_to_db()
                st.session_state.pop("pending_workbook_preview", None)
                st.session_state.pop("pending_workbook_name", None)
                st.success(
                    f"Nova fonte salva em `{saved_path.name}` e sincronizada com sucesso: {imported} registro(s)."
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
            refresh_microsoft_workbook_if_configured()
            sync_google_sheets_to_db()
        except Exception:
            pass


def main() -> None:
    init_db()
    ensure_auth_session_state()
    restore_auth_from_query_params()

    st.markdown(APP_CSS, unsafe_allow_html=True)
    if not st.session_state.get("auth_user"):
        render_login_gate()
        return
    ensure_persistent_auth_query_params()

    maybe_auto_sync_google()
    sync_navigation_state_from_query_params()

    st.markdown(
        f"""
        <div class="hero">
            <div class="brand-row">
                <div class="brand-mark">ON</div>
                <div>
                    <div class="brand-kicker">{PRODUCT_NAME}</div>
                    <div style="color:#d9fbff !important; font-weight:700;">{PRODUCT_PROMISE}</div>
                </div>
            </div>
            <h1 style="color:#ffffff !important;">{PRODUCT_TAGLINE}</h1>
            <p style="color:#f3fbff !important;">
                <span style="color:#f3fbff !important;">
                    Produto para clínicas e profissionais de navegação que precisam transformar
                    planilhas, mensagens e cobranças soltas em uma operação rastreável por ciclo.
                </span>
            </p>
            <div class="hero-actions">
                <span class="hero-pill">Fila prioritária</span>
                <span class="hero-pill">Agenda de infusão</span>
                <span class="hero-pill">Autorização e prescrição</span>
                <span class="hero-pill">Proposta comercial</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="product-strip">
            <div class="product-proof">
                <div class="product-proof-label">Usuário inicial</div>
                <div class="product-proof-value">Profissional navegador</div>
            </div>
            <div class="product-proof">
                <div class="product-proof-label">Cliente pagante</div>
                <div class="product-proof-value">Clínica oncológica</div>
            </div>
            <div class="product-proof">
                <div class="product-proof-label">Primeira oferta</div>
                <div class="product-proof-value">Piloto de 45 dias</div>
            </div>
            <div class="product-proof">
                <div class="product-proof-label">Tese de valor</div>
                <div class="product-proof-value">Ciclos protegidos</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    doctors_df = load_doctors()
    all_patients_df = load_patients()
    patients_df = all_patients_df.copy()
    support_df = load_support_medications()
    sessions_df = load_chemo_sessions()

    with st.sidebar:
        st.caption(f"Acesso: {st.session_state.get('auth_user')}")
        if st.button("Encerrar sessão", use_container_width=True):
            st.session_state["auth_user"] = None
            delete_app_state(REMEMBERED_AUTH_USER_KEY)
            clear_auth_query_params()
            close_patient_detail()
            st.rerun()
        st.markdown("---")
        st.markdown("### Filtros")
        doctor_options = ["Todos"] + doctors_df["name"].tolist()
        selected_doctor = st.selectbox("Médico", doctor_options)

        show_archived_patients = st.toggle("Incluir pacientes arquivados", value=False)
        sidebar_patients_df = patients_df.copy()
        if not show_archived_patients and not sidebar_patients_df.empty:
            sidebar_patients_df = sidebar_patients_df[sidebar_patients_df["active"].fillna(1).astype(int) == 1]

        patient_names = sidebar_patients_df["name"].tolist()
        selected_patient = st.selectbox("Paciente", ["Todos"] + patient_names)

        show_only_attention = st.toggle("Mostrar apenas pacientes com alerta", value=False)
        st.markdown("---")
        st.caption("Versão focada em prescrição, convênio e agenda do ciclo.")
        if find_primary_workbook_file() is not None:
            last_sync = get_app_state("last_google_sync_at")
            st.caption("Planilha principal conectada.")
            if get_app_state(MICROSOFT_WORKBOOK_URL_KEY):
                st.caption("Fonte Microsoft online vinculada.")
            st.caption(f"Última sincronização: {last_sync or 'ainda não sincronizado'}")
            st.caption(f"Sincroniza ao navegar ou ao clicar em sincronizar, sem derrubar o login.")

    patients_df = sidebar_patients_df.copy()
    filtered_patients = patients_df.copy()
    filtered_support = support_df[support_df["patient_name"].isin(patients_df["name"])].copy() if not patients_df.empty else support_df.iloc[0:0].copy()
    filtered_sessions = sessions_df[sessions_df["patient_id"].isin(patients_df["id"])].copy() if not patients_df.empty else sessions_df.iloc[0:0].copy()

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

    tabs = st.tabs(["Visão simples", "Painel operacional", "Gestão de pacientes", "Alertas", "Pacientes", "Médicos", "Planilha principal", "Importação", "Modelo comercial"])
    with tabs[0]:
        render_simple_dashboard(filtered_patients, filtered_support, filtered_sessions)
    with tabs[1]:
        render_dashboard(filtered_patients, filtered_support, filtered_sessions)
    with tabs[2]:
        render_patient_management_tab(doctors_df, all_patients_df)
    with tabs[3]:
        render_alerts_tab(filtered_patients)
    with tabs[4]:
        render_patients_tab(filtered_patients)
    with tabs[5]:
        render_doctors_tab(doctors_df, filtered_patients)
    with tabs[6]:
        render_google_sync_tab(patients_df)
    with tabs[7]:
        render_import_tab()
    with tabs[8]:
        render_commercial_tab(filtered_patients, filtered_sessions)


if __name__ == "__main__":
    main()
