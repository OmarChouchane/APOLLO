#!/usr/bin/env python3
"""
Email Outreach Automation
PyQt6 desktop app for LLM-powered personalized email campaigns.
"""

import sys
import os
import csv
import json
import smtplib
import time
import re
import threading
from datetime import datetime
import unicodedata
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from typing import Optional, List, Dict, Any, Tuple

import requests

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit, QProgressBar,
    QFileDialog, QFrame, QCheckBox, QSpinBox, QGroupBox,
    QMessageBox, QDialog, QDialogButtonBox, QSplitter, QComboBox,
    QScrollArea, QFormLayout, QSizePolicy,
    QTableWidget, QTableWidgetItem, QHeaderView,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QMutex, QWaitCondition
from PyQt6.QtGui import QFont, QTextCursor, QPalette, QColor

# ── Optional deps ──────────────────────────────────────────────────────────────
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    from PyPDF2 import PdfReader
    HAS_PDF = True
except ImportError:
    try:
        from pypdf import PdfReader  # type: ignore
        HAS_PDF = True
    except ImportError:
        HAS_PDF = False

try:
    from docx import Document as DocxDoc  # type: ignore
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ── Constants ──────────────────────────────────────────────────────────────────
APP_NAME    = "قل الحمد لله"
APP_VERSION = "1.0.0"
LOG_FILE    = Path("outreach_log.json")
CONFIG_FILE = Path("config.local.json")
MAX_CV_CHARS = 6000
EMAIL_RE = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")
EMAIL_IN_TEXT_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
FIXED_EMAIL_SUBJECT = "Candidature de stage d’été – DevOps / Cloud / Software Development"

LEGACY_SIGNATURE_BLOCK = (
    "Je serais ravi d'échanger avec vous au sujet de toute opportunité de stage.\n\n"
    "LinkedIn : https://www.linkedin.com/in/omar-chouchane/\n"
    "GitHub : https://github.com/OmarChouchane\n"
    "Portfolio : https://portfolio-omarchouchane.vercel.app/\n\n"
    "Bien cordialement,\n"
    "Omar Chouchane\n\n"
    "omar.chouchane@insat.ucar.tn | +216 52 834 833"
)

# ── Palette (Catppuccin Mocha) ─────────────────────────────────────────────────
C = {
    "base":    "#111318",
    "mantle":  "#171a21",
    "crust":   "#0b0d11",
    "surface0":"#20242d",
    "surface1":"#2b313c",
    "surface2":"#394252",
    "overlay0":"#717987",
    "overlay2":"#a2aab8",
    "text":    "#edf1f7",
    "subtext": "#b6bfcc",
    "blue":    "#6ea8fe",
    "lavender":"#9fb7ff",
    "mauve":   "#a78bfa",
    "pink":    "#f0abfc",
    "red":     "#ff6b7a",
    "maroon":  "#f08a9a",
    "peach":   "#f6b26b",
    "yellow":  "#f7d774",
    "green":   "#67d391",
    "teal":    "#5dd4c2",
    "sky":     "#68c7ee",
    "sapphire":"#4fa4d8",
    "accent":  "#67d391",
}

DARK_STYLE = f"""
QWidget {{
    background-color: {C['base']};
    color: {C['text']};
    font-family: 'Segoe UI Variable', 'Segoe UI', 'SF Pro Display', 'Ubuntu', Arial, sans-serif;
    font-size: 13px;
}}
QMainWindow {{ background-color: {C['base']}; }}

/* ── Buttons ── */
QPushButton {{
    background-color: {C['surface0']}; color: {C['text']};
    border: 1px solid {C['surface1']}; border-radius: 6px;
    padding: 8px 16px; font-weight: 600; min-height: 28px;
}}
QPushButton:hover   {{ background-color: {C['surface1']}; border-color: {C['overlay2']}; color: #ffffff; }}
QPushButton:pressed {{ background-color: {C['surface2']}; padding-top: 9px; padding-bottom: 7px; }}
QPushButton:disabled {{ color: {C['surface1']}; border-color: {C['surface0']}; background-color: {C['mantle']}; }}

QPushButton#btnStart {{
    background-color: {C['accent']};
    color: {C['crust']}; border: none; font-weight: 700; border-radius: 6px;
}}
QPushButton#btnStart:hover    {{ background-color: {C['teal']}; color: {C['crust']}; }}
QPushButton#btnStart:disabled {{ background: {C['surface0']}; color: {C['surface2']}; }}

QPushButton#btnPause {{
    background-color: {C['surface0']};
    color: {C['yellow']}; border: 1px solid {C['surface1']}; font-weight: 700; border-radius: 6px;
}}
QPushButton#btnPause:hover {{ background-color: rgba(247,215,116,0.12); border-color: {C['yellow']}; }}

QPushButton#btnStop {{
    background-color: {C['surface0']};
    color: {C['red']}; border: 1px solid {C['surface1']}; font-weight: 700; border-radius: 6px;
}}
QPushButton#btnStop:hover {{ background-color: rgba(255,107,122,0.12); border-color: {C['red']}; }}

QPushButton#btnBrowse {{
    background-color: transparent;
    color: {C['blue']}; border: 1px solid {C['surface1']}; padding: 5px 14px; font-weight: 600; border-radius: 6px;
}}
QPushButton#btnBrowse:hover {{ background-color: rgba(110,168,254,0.10); border-color: {C['blue']}; }}

QPushButton#btnTest {{
    background-color: {C['surface0']};
    color: {C['text']}; border: 1px solid {C['surface1']}; font-weight: 600; border-radius: 6px;
}}
QPushButton#btnTest:hover {{ background-color: {C['surface1']}; border-color: {C['blue']}; }}

QPushButton#btnGenerate {{
    background-color: {C['blue']};
    color: {C['crust']}; border: none; font-weight: 700; border-radius: 6px;
}}
QPushButton#btnGenerate:hover {{ background-color: {C['sky']}; color: {C['crust']}; }}

/* ── Inputs ── */
QLineEdit, QSpinBox, QComboBox {{
    background-color: {C['crust']}; border: 1px solid {C['surface1']};
    border-radius: 6px; padding: 8px 11px;
    color: {C['text']}; selection-background-color: {C['blue']}; min-height: 22px;
}}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus {{
    border-color: {C['accent']}; background-color: {C['mantle']};
}}
QLineEdit:hover, QSpinBox:hover, QComboBox:hover {{ border-color: {C['overlay0']}; }}

QComboBox::drop-down {{ border: none; width: 22px; }}
QComboBox QAbstractItemView {{
    background-color: {C['surface0']}; border: 1px solid {C['surface1']};
    color: {C['text']}; selection-background-color: {C['blue']};
    selection-color: {C['base']}; outline: none;
}}
QSpinBox::up-button, QSpinBox::down-button {{
    width: 20px; background: {C['surface1']}; border: none; border-radius: 3px;
}}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {{ background: {C['surface2']}; }}

/* ── Text edit / log ── */
QTextEdit {{
    background-color: {C['crust']}; border: 1px solid {C['surface0']};
    border-radius: 6px; padding: 8px; color: {C['text']};
    font-family: 'Consolas', 'JetBrains Mono', 'Fira Code', 'Courier New', monospace;
    font-size: 12px; line-height: 1.4;
}}

/* ── Progress bar ── */
QProgressBar {{
    border: 1px solid {C['surface1']}; border-radius: 6px;
    background-color: {C['mantle']}; text-align: center;
    color: {C['text']}; min-height: 24px; font-weight: 500;
}}
QProgressBar::chunk {{
    background-color: {C['accent']};
    border-radius: 5px;
}}

/* ── Group boxes ── */
QGroupBox {{
    border: 1px solid {C['surface0']}; border-radius: 8px;
    margin-top: 1.25em; padding-top: 0.85em; padding-bottom: 6px;
    font-weight: 700; color: {C['text']};
    background-color: {C['mantle']};
}}
QGroupBox::title {{
    subcontrol-origin: margin; left: 14px; padding: 2px 6px;
    color: {C['overlay2']}; background-color: {C['base']}; border-radius: 4px;
}}

/* ── Checkboxes ── */
QCheckBox {{ color: {C['text']}; spacing: 8px; }}
QCheckBox::indicator {{
    width: 17px; height: 17px;
    border: 1px solid {C['surface1']}; border-radius: 4px;
    background-color: {C['mantle']};
}}
QCheckBox::indicator:hover {{ border-color: {C['blue']}; }}
QCheckBox::indicator:checked {{
    background-color: {C['blue']}; border-color: {C['blue']};
    image: none;
}}

/* ── Scrollbars ── */
QScrollBar:vertical {{ background: {C['base']}; width: 8px; border-radius: 4px; }}
QScrollBar::handle:vertical {{
    background: {C['surface1']}; border-radius: 4px; min-height: 30px;
}}
QScrollBar::handle:vertical:hover {{ background: {C['surface2']}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{ background: {C['base']}; height: 8px; border-radius: 4px; }}
QScrollBar::handle:horizontal {{ background: {C['surface1']}; border-radius: 4px; }}

/* ── Labels ── */
QLabel#lblTitle    {{ font-size: 22px; font-weight: bold; color: {C['blue']}; }}
QLabel#lblSubtitle {{ font-size: 11px; color: {C['overlay2']}; letter-spacing: 0px; }}
QLabel#lblStatus   {{ color: {C['text']}; font-style: normal; }}
QLabel#lblBadge    {{
    background-color: {C['surface0']}; border: 1px solid {C['surface1']};
    border-radius: 10px; padding: 2px 10px; color: {C['text']}; font-size: 11px;
}}
QLabel#lblBadgeGood {{
    background-color: rgba(166,227,161,0.15); border: 1px solid {C['green']};
    border-radius: 10px; padding: 2px 10px; color: {C['green']}; font-size: 11px;
}}

/* ── Dialog ── */
QDialog {{ background-color: {C['base']}; }}
QDialogButtonBox QPushButton {{ min-width: 90px; }}

/* ── Splitter ── */
QSplitter::handle {{ background-color: {C['surface0']}; }}
QSplitter::handle:horizontal {{ width: 1px; }}

/* ── Scroll area ── */
QScrollArea {{ border: none; background-color: transparent; }}
"""

# ── Helpers ────────────────────────────────────────────────────────────────────

def valid_email(addr: str) -> bool:
    return bool(EMAIL_RE.match(addr.strip()))


def _normalize_key_name(value: str) -> str:
    value = unicodedata.normalize("NFKD", value)
    value = value.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\W+", "_", value.strip().lower()).strip("_")


def parse_cv(path: str) -> str:
    suffix = Path(path).suffix.lower()
    text = ""
    if suffix == ".pdf":
        if not HAS_PDF:
            return "[PDF parsing unavailable — install PyPDF2 or pypdf]"
        try:
            reader = PdfReader(path)
            text = " ".join(p.extract_text() or "" for p in reader.pages)
        except Exception as e:
            return f"[PDF error: {e}]"
    elif suffix in (".docx", ".doc"):
        if not HAS_DOCX:
            return "[DOCX parsing unavailable — install python-docx]"
        try:
            doc = DocxDoc(path)
            text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        except Exception as e:
            return f"[DOCX error: {e}]"
    else:
        return "[Unsupported format — use PDF or DOCX]"
    text = re.sub(r"\s+", " ", text).strip()
    return text[:MAX_CV_CHARS] + ("…" if len(text) > MAX_CV_CHARS else "")


def _normalize_contact_record(row: Dict[str, str]) -> Dict[str, str]:
    normalized = dict(row)

    # French CSV aliases used by the real contacts file.
    if not normalized.get("company"):
        for key in (
            "nom_de_l_entreprise",
            "entreprise",
            "company",
            "nom_entreprise",
        ):
            if normalized.get(key):
                normalized["company"] = normalized[key]
                break

    if not normalized.get("email"):
        for key in ("e_mail", "email", "e_mail_address", "mail"):
            if normalized.get(key):
                normalized["email"] = normalized[key]
                break

    if not normalized.get("phone"):
        for key in ("telephone", "téléphone", "phone", "tel"):
            if normalized.get(key):
                normalized["phone"] = normalized[key]
                break

    if not normalized.get("company_context"):
        for key in (
            "description_de_l_activite_anglais",
            "description",
            "activity_description",
            "company_description",
        ):
            if normalized.get(key):
                normalized["company_context"] = normalized[key]
                break

    if not normalized.get("sector"):
        for key in ("secteur_majeur_bvd", "secteur", "sector"):
            if normalized.get(key):
                normalized["sector"] = normalized[key]
                break

    if not normalized.get("country"):
        for key in ("pays_du_stage", "country", "pays", "stage_country"):
            if normalized.get(key):
                normalized["country"] = normalized[key]
                break

    contact_field = normalized.get("contact_entreprise", "").strip()
    if contact_field and not normalized.get("email"):
        m = EMAIL_IN_TEXT_RE.search(contact_field)
        if m:
            normalized["email"] = m.group(0)

    # Preserve a human-readable label for previews when the company name is missing.
    if not normalized.get("name") and normalized.get("contact_entreprise"):
        normalized["name"] = normalized["contact_entreprise"]

    return normalized


def _contacts_language_for_path(path: str) -> str:
    name = Path(path).name.lower()
    if name in {"tech_companies.csv", "contacts_entreprises_2024.csv"}:
        return "fr"
    return ""


def load_contacts(path: str) -> List[Dict[str, str]]:
    # Try to read with pandas if available, letting it infer separators.
    if HAS_PANDAS:
        try:
            df = pd.read_csv(path, sep=None, engine="python")
        except Exception:
            df = pd.read_csv(path)
        # normalize column names to snake_case
        df.columns = [_normalize_key_name(c) for c in df.columns]
        df = df.where(pd.notna(df), "").astype(str)
        records = df.to_dict(orient="records")
        return [_normalize_contact_record(r) for r in records]

    rows: List[Dict[str, str]] = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        sample = f.read(4096)
        f.seek(0)
        try:
            sniffer = csv.Sniffer()
            dialect = sniffer.sniff(sample)
            delimiter = dialect.delimiter
        except Exception:
            delimiter = ","
        reader = csv.DictReader(f, delimiter=delimiter)
        for row in reader:
            normalized: Dict[str, str] = {}
            for k, v in row.items():
                if k is None:
                    continue
                key = _normalize_key_name(k)
                normalized[key] = (v or "").strip()
            rows.append(_normalize_contact_record(normalized))
    return rows


def get_field(row: Dict[str, str], *keys: str, default: str = "") -> str:
    for k in keys:
        v = row.get(k, "").strip()
        if v:
            return v
    return default


def infer_name_from_email(email: str) -> str:
    local = email.split("@", 1)[0].split("+", 1)[0].strip()
    if not local:
        return ""
    parts = [p for p in re.split(r"[._\-]+", local) if p]
    if not parts:
        return ""
    return " ".join(p.capitalize() for p in parts)


def resolve_sender_name(cv_text: str, sender_name: str = "",
                        sender_email: str = "") -> str:
    name = sender_name.strip()
    if name:
        return name
    name = extract_sender_info_from_cv(cv_text).get("name", "").strip()
    if name:
        return name
    return infer_name_from_email(sender_email)


def detect_language_from_country(contact: Dict[str, str]) -> str:
    """Return 'fr' for French or 'en' for English based on contact['country']."""
    country = (contact.get("country") or "").strip().lower()
    if not country:
        return "en"

    # Two-letter ISO codes mapping
    iso_map = {
        "fr": "fr", "tn": "fr", "ma": "fr", "dz": "fr",
        "be": "fr", "ch": "fr", "lu": "fr",
    }
    if len(country) == 2 and country in iso_map:
        return iso_map[country]

    francophone = (
        "france tunisia tunisie morocco maroc algeria algerie belgium belgique "
        "switzerland suisse luxembourg"
    )
    for token in francophone.split():
        if token in country:
            return "fr"
    return "en"


def build_signature_block(sender_name: str, sender_email: str = "",
                          sender_phone: str = "") -> str:
    display_name = (sender_name or "Omar Chouchane").strip() or "Omar Chouchane"
    contact_parts = []
    if sender_email:
        contact_parts.append(sender_email.strip())
    if sender_phone:
        contact_parts.append(sender_phone.strip())
    contact_line = " | ".join(contact_parts)
    if contact_line:
        return f"Bien cordialement,\n{display_name}\n\n{contact_line}"
    return f"Bien cordialement,\n{display_name}"


def enforce_signature(body: str, signature_block: str) -> str:
    body = body.strip()
    signature_block = signature_block.strip()
    if not signature_block:
        return body

    if body.endswith(signature_block):
        return body

    signoff_pattern = (
        r"\n\s*(?:best regards|kind regards|regards|sincerely|best|thanks|thank you|"
        r"cordialement|bien a vous|bien a toi|bien a vous|salutations),?"
        r"\s*\n[\s\S]*$"
    )
    body = re.sub(signoff_pattern, "", body, count=1, flags=re.IGNORECASE).rstrip()
    return f"{body}\n\n{signature_block}"


def sanitize_company_context(ctx: str) -> str:
    """Return a short French paraphrase for company research context.

    If the scraped/researched context appears English-heavy, map common keywords
    to short French phrases to avoid leaking raw English into the email.
    """
    if not ctx:
        return ""
    s = ctx.strip()
    # quick heuristic: presence of common English function words
    english_indicators = (" the ", " and ", " is ", " of ", " for ", "in the", "company", "editor", "bank", "airline", "headquarters")
    low = s.lower()
    score = sum(1 for t in english_indicators if t in low)
    # if many english indicators, produce a short french paraphrase using keywords
    if score >= 2:
        # keyword mapping
        if "bank" in low or "banking" in low or "banque" in low:
            return "éditeur de solutions bancaires en France"
        if "airline" in low or ("air" in low and "line" in low):
            return "acteur du transport aérien"
        if "software" in low or "éditeur" in low or "solutions" in low:
            return "éditeur de logiciels ou de solutions techniques"
        if "consult" in low or "services" in low or "service" in low:
            return "fournisseur de services techniques"
        # fallback short french summary
        words = re.findall(r"[A-Za-z]+", low)[:6]
        brief = " ".join(words).strip()
        return (brief[:200] + "...") if brief else "votre environnement technique"
    # If looks French or neutral, collapse to one cleaned paragraph and strip odd markup
    cleaned = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s)).strip()
    # keep only first 250 chars to avoid long dumps
    return cleaned if len(cleaned) <= 250 else cleaned[:247] + "..."


def ensure_portfolio_links(body: str, sender_extra: str = "") -> str:
    """Replace [lien] placeholders with actual URLs and ensure all links are present."""
    if not body:
        return body
    # Replace all [lien] / [link] placeholders with portfolio URL
    body = re.sub(r"\[\s*lien\s*\]|\[\s*link\s*\]", "https://portfolio-omarchouchane.vercel.app", body, flags=re.IGNORECASE)
    
    low = body.lower()
    links_to_add = []
    
    # Check for and add missing links (only if not already present with actual URL)
    if "linkedin" not in low or "linkedin.com" not in low:
        links_to_add.append("LinkedIn : https://linkedin.com/in/omar-chouchane")
    if "github" not in low or "github.com" not in low:
        links_to_add.append("GitHub : https://github.com/OmarChouchane")
    if "portfolio" not in low or "portfolio-omarchouchane" not in low:
        links_to_add.append("Portfolio : https://portfolio-omarchouchane.vercel.app")
    
    if not links_to_add:
        return body
    
    # remove existing signoff if present (similar to enforce_signature)
    signoff_pattern = (
        r"\n\s*(?:best regards|kind regards|regards|sincerely|best|thanks|thank you|"
        r"cordialement|bien a vous|bien a toi|bien a vous|salutations),?"
        r"\s*\n[\s\S]*$"
    )
    stripped = re.sub(signoff_pattern, "", body, count=1, flags=re.IGNORECASE).rstrip()
    added = stripped + "\n\n" + "\n".join(links_to_add)
    return added


def build_prompt(cv_text: str, contact: Dict[str, str], goal: str,
                 company_context: str = "",
                 sender_name: str = "", sender_email: str = "",
                 sender_phone: str = "", sender_extra: str = "",
                 language: str = "fr") -> str:
    language = "fr"
    name    = get_field(contact, "name", "full_name", "first_name", "contact",
                        default=get_field(contact, "company", "organization", "employer", "firm", default="there"))
    company = get_field(contact, "company", "organization", "employer", "firm",
                        default="your organization")
    role    = get_field(contact, "role", "position", "title", "job_title")

    skip = {"name", "full_name", "first_name", "email", "e_mail", "e-mail", "email_address",
            "company", "organization", "employer", "firm", "role", "position",
            "title", "job_title", "contact", "contact_entreprise", "filiere",
            "pays_du_stage", "pays", "pays_stage", "stage_country", "country",
            "nom_de_l_entreprise", "nom_entreprise", "phone", "telephone", "tel",
            "description", "company_context", "description_de_l_activite_anglais",
            "secteur_majeur_bvd", "sector", "section_nace_rev_2"}
    extras = "\n".join(
        f"- {k.replace('_', ' ').title()}: {v}"
        for k, v in contact.items()
        if k not in skip and v.strip()
    )
    role_line = f"- Role: {role}\n" if role else ""
    context_block = (
        f'\nCompany Research (reformule en francais naturel et ne recopie jamais mot a mot):\n"""\n{company_context}\n"""\n'
        if company_context else ""
    )

    sender_label = resolve_sender_name(cv_text, sender_name, sender_email)
    contact_lines = []
    if sender_phone:
        contact_lines.append(sender_phone)
    if sender_email:
        contact_lines.append(sender_email)
    contact_line = " | ".join(contact_lines)
    contact_info = f"- Contact: {contact_line}\n" if contact_line else ""
    extra_info = f"- Additional info: {sender_extra}\n" if sender_extra.strip() else ""

    signature = build_signature_block(sender_label, sender_email, sender_phone)

    if language and language.lower().startswith("en"):
        # English prompt aligned with the sample email structure and tone
        return f"""You are an expert copywriter for internship outreach emails. Your goal is to help {sender_label} write a standout cold email for a summer internship, with the option of an extended internship or part-time role.

Candidate ({sender_label}):
{contact_info}{extra_info}
Full CV (single source of truth for skills, experience and education):
\"\"\"
{cv_text}
\"\"\"

Recipient:
- Name: {name}
- Company: {company}
{role_line}{extras}{context_block}

Task: write a top-tier cold email in English (en-US) that follows this structure: short greeting, brief interest sentence, one paragraph about profile and skills, one paragraph about why this company, one paragraph about internship interest and fit, then a short attachment / links block, then the exact signature.

Rules (must follow):
1. Use first person ("I") and keep the tone human, concise, and confident.
2. Base everything only on the CV above. Do NOT invent projects, dates, degrees, employers, or metrics not present in the CV.
3. Keep the body powerful but short: 4 paragraphs max, short sentences, no fluff.
4. Mention broad strengths only, not a long technical inventory. Focus on 2-3 areas such as DevOps, Cloud, software development, infrastructure, automation, CI/CD, Kubernetes, Docker, Terraform, backend, reliability.
5. Do not name specific employers or detailed projects unless they are explicitly in the CV and essential. Prefer general framing such as hands-on projects, academic work, and infrastructure-oriented experience.
6. Make the company-specific paragraph concrete and genuine using the company description / sector, and explain why it stands out.
7. Explicitly say the user is looking for a summer internship and is also open to an extended internship or part-time position.
8. Include a short line inviting the reader to review the attached CV and the portfolio / LinkedIn / GitHub links if present.
9. The subject must be short and compelling, for example: "Summer Internship Application – DevOps / Cloud / Software Development".
10. End the body with this exact signature block and nothing after it:
{signature}

Respond ONLY with valid JSON (no surrounding text): {{"subject": "...", "body": "..."}}"""

    # Default: French prompt aligned with the sample email structure and tone
    return f"""Tu es un expert en copywriting d'emails de candidature. Ton objectif est d'aider {sender_label} à rédiger un cold email qui se démarque pour un stage d'été, avec aussi la possibilité d'un stage prolongé ou d'un poste à temps partiel si cela intéresse l'entreprise.

A propos du candidat ({sender_label}) :
{contact_info}{extra_info}
CV complet du candidat (source unique de verite pour ses competences, experiences et formation) :
\"\"\"
{cv_text}
\"\"\"

Destinataire :
- Nom: {name}
- Entreprise: {company}
{role_line}{extras}{context_block}

Tache : redige un email de prospection en francais qui suit exactement cette structure :
1. Une salutation simple et professionnelle.
2. Une phrase d'accroche courte et directe.
3. Un paragraphe sur le profil, la formation et les competences principales.
4. Un paragraphe sur ce qui t'attire chez l'entreprise en t'appuyant sur sa description.
5. Un paragraphe sur le stage recherche et la valeur que tu peux apporter.
6. Un court bloc final AVEC LES TROIS LIENS COMPLETS (pas de placeholders) :
   Portfolio : https://portfolio-omarchouchane.vercel.app | LinkedIn : https://linkedin.com/in/omar-chouchane | GitHub : https://github.com/OmarChouchane
7. La signature exacte, sans rien ajouter apres.

Regles obligatoires :
1. Rédige tout l'email en français naturel (fr-FR), fluide, humain et professionnel.
2. Adopte un ton proche de l'exemple fourni : direct, sincère, sobre, sans formule générique.
3. Commence par une salutation simple et professionnelle. Si le nom est connu, utilise-le ; sinon, utilise "Bonjour {company}," ou "Bonjour l'équipe {company},".
4. La première phrase après la salutation doit dire que tu es étudiant en 4ème année en ingénierie informatique, réseaux et télécommunications à l'INSAT (Institut National des Sciences Appliquées et de Technologie).
5. Utilise "je" à la première personne. Ne parle jamais de "le candidat".
6. Base-toi uniquement sur le CV ci-dessus. N'invente aucun projet, aucune expérience, aucune technologie, aucun chiffre non justifié.
7. Fais ressortir 2 à 3 atouts maximum, en termes larges et puissants : DevOps, Cloud, développement logiciel, automatisation, CI/CD, Docker, Kubernetes, Terraform, backend, fiabilité.
8. Ne cite pas des entreprises ou projets spécifiques comme exemples de réalisations, sauf si cela est explicitement nécessaire et déjà présent dans le CV. Préfère une formulation générale sur les projets, l'infrastructure et l'automatisation.
9. Explique pourquoi l'entreprise t'intéresse en t'appuyant sur sa description ou son secteur, de façon concrète et naturelle.
10. Mentionne clairement que le stage recherché est un stage d'été, que tu es ouvert à un format à distance, et que tu peux aussi envisager une relocalisation si besoin.
11. Le corps du mail doit rester court mais impactant : 4 paragraphes maximum, phrases courtes, sans blabla.
12. Évite strictement les clichés du type "J'espère que vous allez bien".
13. Fais en sorte que le mail paraisse réel, original et pas générique.
14. Le sujet doit être court, pro, accrocheur, et proche du style : "Candidature stage d'été – DevOps / Cloud / Développement logiciel" ou une variante équivalente.
15. N'utilise jamais les libellés du CSV (GL, IIA, IMI, NACE, secteur, etc.) dans le contenu du mail.
16. Les liens Portfolio, LinkedIn et GitHub DOIVENT toujours être des URLs complètes et valides:
   - Portfolio : https://portfolio-omarchouchane.vercel.app
   - LinkedIn : https://linkedin.com/in/omar-chouchane
   - GitHub : https://github.com/OmarChouchane
   Ne utilise JAMAIS de placeholders comme [lien], [link], ou des pointillés.
17. Termine le corps avec ce bloc de signature EXACT, inchangé, sans rien ajouter après :
{signature}
18. Si le contexte entreprise est en anglais, reformule-le en francais naturel et ne recopie jamais de phrases anglaises mot a mot.

Reponds uniquement en JSON valide (sans markdown, sans texte autour).
IMPORTANT: Retourne STRICTEMENT un objet JSON valide et rien d'autre. Exemples ou explications sont interdits.
- L'objet doit avoir exactement deux clés: `subject` et `body`.
- `subject`: la ligne d'objet en français.
- `body`: le contenu complet du message en français, en respectant la structure demandée (salutation, 3 paragraphes principaux + court bloc liens, signature).
- Le `body` doit contenir explicitement (en français) les tokens: "Portfolio", "LinkedIn", "GitHub", "stage d'été", "INSAT".
- Ne recopie jamais des libellés CSV ou des métadonnées brutes; reformule en français naturel.
Retour final (strictement): {{"subject": "...", "body": "..."}}"""


def build_french_fallback_email(cv_text: str, contact: Dict[str, str], goal: str,
                                company_context: str = "",
                                sender_name: str = "", sender_email: str = "",
                                sender_phone: str = "", sender_extra: str = "") -> Tuple[str, str]:
    sender_label = resolve_sender_name(cv_text, sender_name, sender_email)
    company = get_field(contact, "company", "organization", "employer", "firm", default="votre equipe")
    recipient_name = get_field(contact, "name", "full_name", "first_name", "contact", default="")
    greeting = f"Bonjour {recipient_name}," if recipient_name and recipient_name != company else f"Bonjour {company},"

    context = company_context.strip()
    if not context:
        context = "votre environnement technique et votre approche orientée produit"

    subject = FIXED_EMAIL_SUBJECT

    signature = build_signature_block(sender_label, sender_email, sender_phone)
    portfolio_lines = [
        "Portfolio : portfolio-omarchouchanes-projects.vercel.app",
        "LinkedIn : linkedin.com/in/omar-chouchane",
        "GitHub : github.com/OmarChouchane",
    ]
    if sender_extra.strip():
        portfolio_lines.insert(0, sender_extra.strip())

    body = f"""{greeting}

Je suis étudiant en 4ème année en ingénierie informatique, réseaux et télécommunications à l'INSAT (Institut National des Sciences Appliquées et de Technologie).

Je vous contacte parce que {context} correspond exactement au type d'environnement que je souhaite rejoindre.

Je porte un fort intérêt pour le DevOps, le Cloud et le développement logiciel. Mes projets m'ont amené à travailler sur Kubernetes, Docker, Terraform, les pipelines CI/CD et l'automatisation d'infrastructures, avec une attention particulière à la fiabilité et à la simplicité d'exploitation.

Ce qui m'intéresse chez vous, c'est la possibilité de contribuer à des systèmes concrets tout en continuant à apprendre aux côtés d'une équipe qui construit des produits ou services techniques de manière sérieuse. Je recherche un stage d'été, je suis ouvert à un format à distance, et je peux aussi envisager une relocalisation si besoin.

Veuillez trouver mon CV ci-joint. Je serais ravi d'échanger avec vous sur une opportunité qui pourrait vous être utile.

{chr(10).join(portfolio_lines)}

{signature}"""
    return subject, body


def is_complete_french_email(body: str) -> bool:
    text = (body or "").strip()
    # Accept if all required tokens are present regardless of length
    required_tokens = ("Portfolio", "LinkedIn", "GitHub", "stage d'été", "INSAT")
    lowers = text.lower()
    found = sum(1 for t in required_tokens if t.lower() in lowers)
    if found == len(required_tokens):
        return True
    # Relaxed acceptance: if most tokens present and reasonable length
    if found >= 4 and len(text) >= 50:
        return True
    # Otherwise consider incomplete
    return False


# ── Company Researcher ─────────────────────────────────────────────────────────

class CompanyResearcher:
    _UA = "Mozilla/5.0 (compatible; outreach-research/1.0)"

    def __init__(self):
        self._cache: Dict[str, str] = {}

    def research(self, company: str, website: str = "") -> str:
        key = company.strip().lower()
        if key in self._cache:
            return self._cache[key]
        ctx = self._scrape(website) if website else ""
        if not ctx:
            ctx = self._duckduckgo(company)
        self._cache[key] = ctx
        return ctx

    def _scrape(self, url: str) -> str:
        if not url.startswith("http"):
            url = "https://" + url
        try:
            r = requests.get(url, timeout=10, headers={"User-Agent": self._UA})
            r.raise_for_status()
            html = r.text
            for pat in [
                r'<meta[^>]+name=["\']description["\'][^>]+content=["\'](.*?)["\']',
                r'<meta[^>]+content=["\'](.*?)["\'][^>]+name=["\']description["\']',
                r'<meta[^>]+property=["\']og:description["\'][^>]+content=["\'](.*?)["\']',
                r'<meta[^>]+content=["\'](.*?)["\'][^>]+property=["\']og:description["\']',
            ]:
                m = re.search(pat, html, re.IGNORECASE | re.DOTALL)
                if m:
                    desc = m.group(1).strip()
                    if len(desc) > 40:
                        return desc[:600]
            text = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html,
                          flags=re.DOTALL | re.IGNORECASE)
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s+", " ", text).strip()
            return text[:600]
        except Exception:
            return ""

    def _duckduckgo(self, company: str) -> str:
        try:
            r = requests.get(
                "https://api.duckduckgo.com/",
                params={"q": company, "format": "json", "no_html": "1", "skip_disambig": "1"},
                timeout=12,
                headers={"User-Agent": self._UA},
            )
            data = r.json()
            abstract = data.get("AbstractText", "").strip()
            if abstract:
                return abstract[:600]
            topics = data.get("RelatedTopics", [])
            if topics and isinstance(topics[0], dict):
                return topics[0].get("Text", "")[:400]
        except Exception:
            pass
        return ""


def extract_sender_info_from_cv(cv_text: str) -> Dict[str, str]:
    info: Dict[str, str] = {"name": "", "phone": "", "email": "", "extra": ""}
    m = re.search(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", cv_text)
    if m:
        info["email"] = m.group(0)
    m = re.search(r"(?:\+?\d[\d\s\-().]{6,}\d)", cv_text)
    if m:
        info["phone"] = m.group(0).strip()
    for line in cv_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        words = line.split()
        if (2 <= len(words) <= 4
                and all(w[0].isupper() for w in words if w)
                and not any(ch.isdigit() for ch in line)
                and len(line) < 60):
            info["name"] = line
            break
    return info


def _decode_json_string_fragment(value: str) -> str:
    try:
        return json.loads(f'"{value}"')
    except json.JSONDecodeError:
        return (value.replace(r"\n", "\n")
                     .replace(r"\r", "\r")
                     .replace(r"\t", "\t")
                     .replace(r"\"", '"')
                     .replace(r"\\", "\\"))


def _extract_json_string_value(text: str, key: str) -> str:
    m = re.search(rf'"{re.escape(key)}"\s*:\s*"', text)
    if not m:
        return ""

    chars: List[str] = []
    escaped = False
    for ch in text[m.end():]:
        if escaped:
            chars.append("\\" + ch)
            escaped = False
            continue
        if ch == "\\":
            escaped = True
            continue
        if ch == '"':
            break
        chars.append(ch)

    return _decode_json_string_fragment("".join(chars)).strip()


def parse_llm_json(text: str) -> Tuple[str, str]:
    cleaned = re.sub(r"```(?:json)?\s*|\s*```", "", text).strip()
    try:
        data = json.loads(cleaned)
        return data.get("subject", ""), data.get("body", "")
    except json.JSONDecodeError:
        subject = _extract_json_string_value(cleaned, "subject") or "Outreach"
        body = _extract_json_string_value(cleaned, "body")
        if not body and cleaned.lstrip().startswith("{"):
            raise ValueError("LLM returned malformed JSON without a readable body")
        body = body or cleaned
        return subject, body


# ── LLM Client ─────────────────────────────────────────────────────────────────

class LLMProviderError(Exception):
    def __init__(self, provider: str, status_code: int = None, text: str = ""):
        super().__init__(f"Provider {provider} error: {status_code}")
        self.provider = provider
        self.status_code = status_code
        self.text = text


class LLMClient:
    _REGISTRY: Dict[str, Tuple[str, str]] = {
        "OpenAI":      ("https://api.openai.com/v1/chat/completions",
                        "openai/gpt-oss-120b"),
        "Gemini":      ("https://generativelanguage.googleapis.com/v1beta/models",
                        "gemini-2.0-flash"),
        "Groq":        ("https://api.groq.com/openai/v1/chat/completions",
                        "llama-3.3-70b-versatile"),
        "OpenRouter":  ("https://openrouter.ai/api/v1/chat/completions",
                        "google/gemini-2.0-flash-exp:free"),
        "Anthropic":   ("https://api.anthropic.com/v1/messages",
                        "claude-3-5-haiku-20241022"),
        "Mistral":     ("https://api.mistral.ai/v1/chat/completions",
                        "mistral-small-latest"),
        "Together AI": ("https://api.together.xyz/v1/chat/completions",
                        "meta-llama/Llama-3.3-70B-Instruct-Turbo"),
        "Cohere":      ("https://api.cohere.com/v2/chat",
                        "command-r-plus-08-2024"),
    }
    MAX_RETRIES  = 3
    BASE_BACKOFF = 10
    GENERATE_TIMEOUT = 60  # 120s timeout per contact to prevent indefinite hangs

    def __init__(self, provider: str, api_key: str, model: str = "", llm_debug_provider: bool = False):
        self.provider  = provider
        self.api_key   = api_key
        base, default  = self._REGISTRY.get(provider, ("", ""))
        self._base_url = base
        self.model     = model.strip() or default
        self.llm_debug_provider = bool(llm_debug_provider)

    def generate(self, prompt: str) -> str:
        """Generate with timeout wrapper to prevent indefinite hangs."""
        result = [None]
        exception = [None]

        def _generate():
            try:
                result[0] = self._generate_internal(prompt)
            except Exception as e:
                exception[0] = e

        thread = threading.Thread(target=_generate, daemon=True)
        thread.start()
        thread.join(timeout=self.GENERATE_TIMEOUT)

        if thread.is_alive():
            raise TimeoutError(f"LLM generation timed out after {self.GENERATE_TIMEOUT}s for {self.provider}")
        if exception[0]:
            raise exception[0]
        if result[0] is None:
            raise RuntimeError(f"LLM generation returned None for {self.provider}/{self.model}")
        return result[0]

    def _generate_internal(self, prompt: str) -> str:
        """Internal generate logic with retries (moved from original generate)."""
        dispatch = {
            "Gemini":    self._gemini,
            "Anthropic": self._anthropic,
            "Cohere":    self._cohere,
        }
        fn = dispatch.get(self.provider, self._openai_compat)
        for attempt in range(self.MAX_RETRIES):
            try:
                return fn(prompt)
            except requests.HTTPError as e:
                status = None
                text = ""
                if e.response is not None:
                    status = e.response.status_code
                    try:
                        text = e.response.text
                    except Exception:
                        text = ""
                # On 402/429 surface a provider error so callers can attempt fallback
                if status in (402, 429):
                    raise LLMProviderError(self.provider, status, text)
                # Otherwise, handle 429 with backoff
                if status == 429:
                    wait = self.BASE_BACKOFF * (2 ** attempt)
                    ra = e.response.headers.get("Retry-After")
                    if ra:
                        try:
                            wait = max(wait, int(ra))
                        except Exception:
                            pass
                    if attempt < self.MAX_RETRIES - 1:
                        time.sleep(wait)
                        continue
                raise
            except requests.RequestException:
                # Network/connection errors — let retries loop handle
                if attempt < self.MAX_RETRIES - 1:
                    time.sleep(self.BASE_BACKOFF * (2 ** attempt))
                    continue
        raise RuntimeError("LLM retries exhausted")

    def _gemini(self, prompt: str) -> str:
        url = f"{self._base_url}/{self.model}:generateContent?key={self.api_key}"
        try:
            r = requests.post(
                url,
                json={
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generationConfig": {"temperature": 0.7, "maxOutputTokens": 1024},
                },
                timeout=45,
            )
            # Capture provider response for debugging if requested
            if self.llm_debug_provider:
                try:
                    p = Path(f"llm_provider_{self.provider}_resp_{int(time.time())}.txt")
                    p.write_text(f"URL: {url}\nSTATUS: {r.status_code}\n\n{r.text}", encoding="utf-8")
                except Exception:
                    pass
            r.raise_for_status()
            return r.json()["candidates"][0]["content"]["parts"][0]["text"]
        except requests.RequestException as exc:
            # write available response text if present
            try:
                if hasattr(exc, 'response') and exc.response is not None and self.llm_debug_provider:
                    p = Path(f"llm_provider_{self.provider}_error_{int(time.time())}.txt")
                    p.write_text(f"ERROR: {exc}\nSTATUS: {exc.response.status_code}\n\n{exc.response.text}", encoding="utf-8")
            except Exception:
                pass
            raise

    def _openai_compat(self, prompt: str) -> str:
        headers: Dict[str, str] = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type":  "application/json",
        }
        if self.provider == "OpenRouter":
            headers["HTTP-Referer"] = "https://email-outreach-app"
            headers["X-Title"]      = "Email Outreach Automation"
        try:
            r = requests.post(
                self._base_url,
                headers=headers,
                json={
                    "model":       self.model,
                    "messages":    [{"role": "user", "content": prompt}],
                    "temperature": 0.7,
                    "max_tokens":  1024,
                },
                timeout=45,
            )
            if self.llm_debug_provider:
                try:
                    p = Path(f"llm_provider_{self.provider}_resp_{int(time.time())}.txt")
                    p.write_text(f"URL: {self._base_url}\nSTATUS: {r.status_code}\n\n{r.text}", encoding="utf-8")
                except Exception:
                    pass
            r.raise_for_status()
            data = r.json()
            choices = data.get("choices") or []
            if not choices:
                raise LLMProviderError(self.provider, r.status_code, f"OpenAI-compatible response missing choices: {data}")
            message = choices[0].get("message") or {}
            content = message.get("content")
            if content is None:
                finish_reason = choices[0].get("finish_reason")
                detail = {
                    "finish_reason": finish_reason,
                    "message_keys": list(message.keys()),
                    "response_keys": list(data.keys()),
                }
                raise LLMProviderError(self.provider, r.status_code, f"OpenAI-compatible response missing content: {detail}")
            return content
        except requests.RequestException as exc:
            try:
                if hasattr(exc, 'response') and exc.response is not None and self.llm_debug_provider:
                    p = Path(f"llm_provider_{self.provider}_error_{int(time.time())}.txt")
                    p.write_text(f"ERROR: {exc}\nSTATUS: {exc.response.status_code}\n\n{exc.response.text}", encoding="utf-8")
            except Exception:
                pass
            raise

    def _anthropic(self, prompt: str) -> str:
        try:
            r = requests.post(
                self._base_url,
                headers={
                    "x-api-key":         self.api_key,
                    "anthropic-version": "2023-06-01",
                    "Content-Type":      "application/json",
                },
                json={
                    "model":      self.model,
                    "max_tokens": 1024,
                    "messages":   [{"role": "user", "content": prompt}],
                },
                timeout=45,
            )
            if self.llm_debug_provider:
                try:
                    p = Path(f"llm_provider_{self.provider}_resp_{int(time.time())}.txt")
                    p.write_text(f"URL: {self._base_url}\nSTATUS: {r.status_code}\n\n{r.text}", encoding="utf-8")
                except Exception:
                    pass
            r.raise_for_status()
            return r.json()["content"][0]["text"]
        except requests.RequestException as exc:
            try:
                if hasattr(exc, 'response') and exc.response is not None and self.llm_debug_provider:
                    p = Path(f"llm_provider_{self.provider}_error_{int(time.time())}.txt")
                    p.write_text(f"ERROR: {exc}\nSTATUS: {exc.response.status_code}\n\n{exc.response.text}", encoding="utf-8")
            except Exception:
                pass
            raise

    def _cohere(self, prompt: str) -> str:
        try:
            r = requests.post(
                self._base_url,
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type":  "application/json",
                },
                json={
                    "model":    self.model,
                    "messages": [{"role": "user", "content": prompt}],
                },
                timeout=45,
            )
            if self.llm_debug_provider:
                try:
                    p = Path(f"llm_provider_{self.provider}_resp_{int(time.time())}.txt")
                    p.write_text(f"URL: {self._base_url}\nSTATUS: {r.status_code}\n\n{r.text}", encoding="utf-8")
                except Exception:
                    pass
            r.raise_for_status()
            return r.json()["message"]["content"][0]["text"]
        except requests.RequestException as exc:
            try:
                if hasattr(exc, 'response') and exc.response is not None and self.llm_debug_provider:
                    p = Path(f"llm_provider_{self.provider}_error_{int(time.time())}.txt")
                    p.write_text(f"ERROR: {exc}\nSTATUS: {exc.response.status_code}\n\n{exc.response.text}", encoding="utf-8")
            except Exception:
                pass
            raise


# ── Email Sender ───────────────────────────────────────────────────────────────

class EmailSender:
    def __init__(self, sender: str, password: str,
                 host: str = "smtp.gmail.com", port: int = 587):
        self.sender   = sender
        self.password = password
        self.host     = host
        self.port     = port

    def send(self, to: str, subject: str, body: str,
             cv_path: Optional[str] = None) -> None:
        msg = MIMEMultipart("mixed")
        msg["From"]    = self.sender
        msg["To"]      = to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))

        if cv_path and os.path.isfile(cv_path):
            with open(cv_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{Path(cv_path).name}"',
            )
            msg.attach(part)

        with smtplib.SMTP(self.host, self.port, timeout=30) as srv:
            srv.ehlo()
            srv.starttls()
            srv.login(self.sender, self.password)
            srv.sendmail(self.sender, to, msg.as_string())


# ── Preview Dialog ─────────────────────────────────────────────────────────────

class PreviewDialog(QDialog):
    def __init__(self, name: str, email: str, subject: str, body: str,
                 parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Preview — {name} <{email}>")
        self.resize(660, 540)
        lay = QVBoxLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(16, 16, 16, 16)

        to_lbl = QLabel(f"<b>To:</b> {name} &lt;{email}&gt;")
        to_lbl.setStyleSheet(f"color: {C['subtext']}; font-size: 13px;")
        lay.addWidget(to_lbl)

        lay.addWidget(QLabel("Subject:"))
        self.inp_subject = QLineEdit(subject)
        lay.addWidget(self.inp_subject)

        lay.addWidget(QLabel("Body:"))
        self.txt_body = QTextEdit()
        self.txt_body.setPlainText(body)
        self.txt_body.setMinimumHeight(320)
        lay.addWidget(self.txt_body)

        btns = QDialogButtonBox()
        btns.addButton("Send", QDialogButtonBox.ButtonRole.AcceptRole)
        btns.addButton("Skip", QDialogButtonBox.ButtonRole.RejectRole)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    @property
    def subject(self) -> str:
        return self.inp_subject.text().strip()

    @property
    def body(self) -> str:
        return self.txt_body.toPlainText().strip()


# ── Draft Email ────────────────────────────────────────────────────────────────

class DraftEmail:
    __slots__ = ("name", "email", "company", "subject", "body",
                 "contact", "include", "company_context")

    def __init__(self, name: str, email: str, company: str,
                 subject: str, body: str, contact: Dict,
                 company_context: str = ""):
        self.name            = name
        self.email           = email
        self.company         = company
        self.subject         = subject
        self.body            = body
        self.contact         = contact
        self.include         = True
        self.company_context = company_context


# ── CV Info Extraction Thread ──────────────────────────────────────────────────

class CVInfoThread(QThread):
    sig_done = pyqtSignal(dict)

    def __init__(self, cv_text: str,
                 provider: str = "", api_key: str = "", model: str = "",
                 llm_debug_provider: bool = False,
                 parent=None):
        super().__init__(parent)
        self.cv_text  = cv_text
        self.provider = provider
        self.api_key  = api_key
        self.model    = model
        self.llm_debug_provider = bool(llm_debug_provider)

    def run(self):
        info = extract_sender_info_from_cv(self.cv_text)
        if self.api_key and self.provider:
            try:
                llm = LLMClient(self.provider, self.api_key, self.model, llm_debug_provider=self.llm_debug_provider)
                prompt = (
                    "Extract the following fields from this CV text and return ONLY valid JSON "
                    "(no markdown, no extra text):\n"
                    '{"name": "full name of the CV owner", '
                    '"phone": "phone number or empty string", '
                    '"email": "email address or empty string", '
                    '"extra": "LinkedIn URL or portfolio URL if present, else empty string"}\n\n'
                    f'CV:\n"""\n{self.cv_text}\n"""'
                )
                raw = llm.generate(prompt)
                cleaned = re.sub(r"```(?:json)?\s*|\s*```", "", raw).strip()
                parsed = json.loads(cleaned)
                for k in ("name", "phone", "email", "extra"):
                    if parsed.get(k, "").strip():
                        info[k] = parsed[k].strip()
            except Exception:
                pass
        self.sig_done.emit(info)


# ── Generation Thread ──────────────────────────────────────────────────────────

class GenerationThread(QThread):
    sig_log      = pyqtSignal(str, str)
    sig_progress = pyqtSignal(int, int)
    sig_status   = pyqtSignal(str)
    sig_done     = pyqtSignal(list)

    def __init__(self, cfg: Dict[str, Any], contacts: List[Dict],
                 cv_text: str, parent=None):
        super().__init__(parent)
        self.cfg      = cfg
        self.contacts = contacts
        self.cv_text  = cv_text
        self._stopped = False
        self._researcher = CompanyResearcher()

    def stop(self):
        self._stopped = True

    def run(self):
        cfg   = self.cfg
        llm   = LLMClient(cfg["provider"], cfg["api_key"], cfg.get("model", ""), llm_debug_provider=cfg.get("llm_debug_provider", False))
        total = len(self.contacts)
        drafts: List[DraftEmail] = []

        self.sig_log.emit(f"Generating {total} draft(s)…", "info")
        self.sig_progress.emit(0, total)

        for i, contact in enumerate(self.contacts):
            if self._stopped:
                break

            email   = get_field(contact, "email", "e-mail", "email_address", "mail")
            name    = get_field(contact, "name", "full_name", "first_name",
                                default=f"Contact {i + 1}")
            company = get_field(contact, "company", "organization", "employer", "firm",
                                default="")

            if not email or not valid_email(email):
                self.sig_log.emit(
                    f"[{i+1}/{total}] Invalid/missing email for '{name}' — skipped", "warn"
                )
                self.sig_progress.emit(i + 1, total)
                continue

            company_context = get_field(
                contact,
                "company_context",
                "description",
                "description_de_l_activite_anglais",
                default="",
            )
            company_context = sanitize_company_context(company_context)
            if cfg.get("research") and company:
                self.sig_status.emit(f"Researching {company}…")
                website = get_field(contact, "website", "url", "homepage", "site")
                researched_context = self._researcher.research(company, website)
                if researched_context:
                    researched_context = sanitize_company_context(researched_context)
                    company_context = f"{company_context}\n\n{researched_context}".strip() if company_context else researched_context
                    self.sig_log.emit(
                        f"[{i+1}/{total}] Context: {company} ({len(researched_context)} chars)",
                        "info",
                    )
                else:
                    self.sig_log.emit(
                        f"[{i+1}/{total}] No context found for '{company}'", "warn"
                    )

            self.sig_status.emit(f"Generating draft for {name}…")
            self.sig_log.emit(f"[{i+1}/{total}] Generating → {name} <{email}>", "info")
            try:
                lang = "fr"
                retries = int(cfg.get("llm_retries", 2))
                subject = body = ""
                for attempt in range(1, retries + 1):
                    prompt_text = build_prompt(
                        self.cv_text, contact, cfg["goal"], company_context,
                        sender_name=cfg.get("sender_name", ""),
                        sender_email=cfg.get("sender_email", ""),
                        sender_phone=cfg.get("sender_phone", ""),
                        sender_extra=cfg.get("sender_extra", ""),
                        language=lang,
                    )
                    try:
                        raw = llm.generate(prompt_text)
                    except LLMProviderError as ple:
                        self.sig_log.emit(f"[{i+1}/{total}] Provider {ple.provider} returned {ple.status_code} — attempting fallbacks", "warn")
                        # Try configured fallbacks from cfg
                        fallbacks = cfg.get("provider_fallbacks", [])
                        fallback_succeeded = False
                        for fb in fallbacks:
                            try:
                                p_name = fb.get("provider")
                                p_key = fb.get("api_key")
                                p_model = fb.get("model", "")
                                if not p_name or not p_key:
                                    continue
                                self.sig_log.emit(f"[{i+1}/{total}] Trying fallback provider {p_name}", "info")
                                fb_llm = LLMClient(p_name, p_key, p_model, llm_debug_provider=cfg.get("llm_debug_provider", False))
                                raw = fb_llm.generate(prompt_text)
                                fallback_succeeded = True
                                # replace llm so subsequent repair attempts use the successful provider
                                llm = fb_llm
                                break
                            except Exception as e:
                                self.sig_log.emit(f"[{i+1}/{total}] Fallback {fb.get('provider')} failed: {e}", "warn")
                                continue
                        if not fallback_succeeded:
                            raise
                    try:
                        subject, body = parse_llm_json(raw)
                    except Exception as e:
                        self.sig_log.emit(f"[{i+1}/{total}] Parse error from LLM (attempt {attempt}): {e}", "warn")
                        body = ""
                    if is_complete_french_email(body):
                        self.sig_log.emit(f"[{i+1}/{total}] LLM produced complete French body on attempt {attempt}", "info")
                        break
                    else:
                        self.sig_log.emit(f"[{i+1}/{total}] LLM output failed completeness check (attempt {attempt})", "warn")
                        # Emit truncated raw output for quick debugging
                        try:
                            self.sig_log.emit(f"[{i+1}/{total}] Raw LLM output (truncated): {raw[:800]}", "debug")
                        except Exception:
                            pass
                        # If enabled in config, write full raw output to disk for post-mortem
                        try:
                            if cfg.get("llm_debug"):
                                p = Path(f"llm_debug_gen_{i+1}_attempt{attempt}.txt")
                                p.write_text(raw or "", encoding="utf-8")
                                self.sig_log.emit(f"[{i+1}/{total}] Full LLM output written to {p}", "info")
                        except Exception:
                            pass
                        if attempt < retries:
                            time.sleep(1)
                        # If subject present but body incomplete, try a focused repair prompt
                        try:
                            if subject and not is_complete_french_email(body):
                                repair_tries = 2
                                for r_i in range(1, repair_tries + 1):
                                    repair_prompt = (
                                        f"Le modèle a renvoyé ceci:\n{raw}\n\n"
                                        f"Le sujet détecté est : {subject}\n"
                                        "Le champ 'body' est manquant ou incomplet. En te basant uniquement sur le CV et le contexte fournis, retourne STRICTEMENT un objet JSON valide avec exactement les deux clés 'subject' et 'body' (aucun texte supplémentaire). Ne modifie pas le 'subject'. Le 'body' doit être en français et contenir explicitement les tokens : Portfolio, LinkedIn, GitHub, stage d'été, INSAT.\n"
                                        f"CV:\n\"\"\"\n{self.cv_text}\n\"\"\"\n"
                                        f"Contexte entreprise:\n\"\"\"\n{company_context}\n\"\"\"\n"
                                    )
                                    self.sig_log.emit(f"[{i+1}/{total}] Repair attempt {r_i} for body", "info")
                                    repair_raw = llm.generate(repair_prompt)
                                    try:
                                        subject_r, body_r = parse_llm_json(repair_raw)
                                    except Exception as e:
                                        self.sig_log.emit(f"[{i+1}/{total}] Parse error from repair LLM attempt {r_i}: {e}", "warn")
                                        body_r = ""
                                    if is_complete_french_email(body_r):
                                        body = body_r
                                        subject = subject or subject_r or subject
                                        self.sig_log.emit(f"[{i+1}/{total}] Repair succeeded on attempt {r_i}", "info")
                                        break
                                    time.sleep(1)
                        except Exception:
                            pass
                        continue
                if not is_complete_french_email(body):
                    subject, body = build_french_fallback_email(
                        self.cv_text, contact, cfg["goal"], company_context,
                        sender_name=cfg.get("sender_name", ""),
                        sender_email=cfg.get("sender_email", ""),
                        sender_phone=cfg.get("sender_phone", ""),
                        sender_extra=cfg.get("sender_extra", ""),
                    )
                    self.sig_log.emit(
                        f"[{i+1}/{total}] Incomplete French body after {retries} attempts — using fallback French draft", "warn"
                    )
                subject = FIXED_EMAIL_SUBJECT
                sender_name = resolve_sender_name(
                    self.cv_text,
                    cfg.get("sender_name", ""),
                    cfg.get("sender_email", ""),
                )
                # Ensure portfolio / links are present before appending signature
                body = ensure_portfolio_links(body, cfg.get("sender_extra", ""))
                signature = build_signature_block(
                    sender_name,
                    cfg.get("sender_email", ""),
                    cfg.get("sender_phone", ""),
                )
                body = enforce_signature(body, signature)
            except Exception as exc:
                self.sig_log.emit(f"[{i+1}/{total}] LLM error: {exc}", "error")
                self.sig_progress.emit(i + 1, total)
                continue

            drafts.append(
                DraftEmail(name, email, company, subject, body, contact, company_context)
            )
            self.sig_log.emit(
                f"[{i+1}/{total}] ✓ Draft ready: {name} — {subject}", "success"
            )
            self.sig_progress.emit(i + 1, total)

            if i < total - 1 and not self._stopped:
                time.sleep(3)

        status = "All drafts generated." if not self._stopped else "Generation stopped."
        self.sig_status.emit(status)
        self.sig_done.emit(drafts)


# ── Preview-All Dialog ─────────────────────────────────────────────────────────

class PreviewAllDialog(QDialog):
    TABLE_STYLE = f"""
        QTableWidget {{
            background-color: {C['base']};
            alternate-background-color: {C['mantle']};
            gridline-color: transparent;
            border: none;
        }}
        QTableWidget::item {{ padding: 5px 8px; border-radius: 0px; }}
        QTableWidget::item:selected {{ background-color: {C['surface0']}; color: {C['text']}; }}
        QHeaderView::section {{
            background-color: {C['surface0']}; color: {C['blue']};
            border: none; border-bottom: 1px solid {C['surface1']};
            padding: 7px 6px; font-weight: 600;
        }}
    """

    def __init__(self, drafts: List[DraftEmail], parent=None):
        super().__init__(parent)
        self._drafts          = drafts
        self._current_row     = -1
        self._blocking        = False
        self.setWindowTitle(f"Review {len(drafts)} Generated Email(s)")
        self.setMinimumSize(960, 620)
        self.resize(1140, 700)
        self._build_ui()
        if drafts:
            self.table.selectRow(0)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(12)
        root.setContentsMargins(16, 16, 16, 16)

        # Header bar
        hdr = QWidget()
        hdr.setStyleSheet(f"background-color: {C['surface0']}; border-radius: 8px;")
        hh = QHBoxLayout(hdr)
        hh.setContentsMargins(14, 10, 14, 10)
        info_lbl = QLabel(f"  {len(self._drafts)} email draft(s) ready — review, edit, then send")
        info_lbl.setStyleSheet(f"color: {C['blue']}; font-weight: 600; font-size: 13px; background: transparent;")
        hh.addWidget(info_lbl)
        hh.addStretch()
        hint = QLabel("Click a row to preview & edit  ·  uncheck to skip")
        hint.setStyleSheet(f"color: {C['overlay0']}; font-size: 11px; background: transparent;")
        hh.addWidget(hint)
        root.addWidget(hdr)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(1)
        splitter.addWidget(self._build_left())
        splitter.addWidget(self._build_right())
        splitter.setSizes([380, 680])
        root.addWidget(splitter, 1)
        root.addWidget(self._build_bottom())

    def _build_left(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 6, 0)
        v.setSpacing(6)

        self.table = QTableWidget(len(self._drafts), 3)
        self.table.setHorizontalHeaderLabels(["", "Recipient", "Subject"])
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 34)
        self.table.setColumnWidth(1, 186)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet(self.TABLE_STYLE)

        for r, d in enumerate(self._drafts):
            self.table.setRowHeight(r, 50)
            chk = QTableWidgetItem()
            chk.setFlags(
                Qt.ItemFlag.ItemIsUserCheckable |
                Qt.ItemFlag.ItemIsEnabled |
                Qt.ItemFlag.ItemIsSelectable
            )
            chk.setCheckState(Qt.CheckState.Checked)
            self.table.setItem(r, 0, chk)
            self.table.setItem(r, 1, QTableWidgetItem(f"{d.name}\n{d.email}"))
            self.table.setItem(r, 2, QTableWidgetItem(d.subject))

        self.table.currentCellChanged.connect(lambda cur, _cc, _pr, _pc: self._on_row_changed(cur))
        self.table.itemChanged.connect(self._on_item_changed)
        v.addWidget(self.table)

        bw = QWidget()
        bh = QHBoxLayout(bw)
        bh.setContentsMargins(0, 0, 0, 0)
        b_all  = QPushButton("Select All")
        b_none = QPushButton("Deselect All")
        b_all.clicked.connect(lambda: self._bulk_check(True))
        b_none.clicked.connect(lambda: self._bulk_check(False))
        bh.addWidget(b_all)
        bh.addWidget(b_none)
        bh.addStretch()
        v.addWidget(bw)
        return w

    def _build_right(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(6, 0, 0, 0)
        v.setSpacing(10)

        self.lbl_to = QLabel("Select a draft to preview")
        self.lbl_to.setStyleSheet(f"color: {C['text']}; font-weight: 600; font-size: 13px;")
        self.lbl_to.setWordWrap(True)
        v.addWidget(self.lbl_to)

        self.lbl_context = QLabel("")
        self.lbl_context.setStyleSheet(
            f"color: {C['overlay0']}; font-size: 11px; "
            f"background-color: {C['surface0']}; border-radius: 5px; padding: 4px 8px;"
        )
        self.lbl_context.setWordWrap(True)
        self.lbl_context.setMaximumHeight(46)
        self.lbl_context.setVisible(False)
        v.addWidget(self.lbl_context)

        subj_row = QHBoxLayout()
        lbl_subj = QLabel("Subject:")
        lbl_subj.setStyleSheet(f"color: {C['subtext']}; font-weight: 500;")
        subj_row.addWidget(lbl_subj)
        subj_row.addStretch()
        lbl_edit_hint = QLabel("Editable")
        lbl_edit_hint.setStyleSheet(f"color: {C['peach']}; font-size: 11px;")
        subj_row.addWidget(lbl_edit_hint)
        v.addLayout(subj_row)
        self.inp_subject = QLineEdit()
        self.inp_subject.setPlaceholderText("Select a draft to edit its subject…")
        self.inp_subject.setStyleSheet(
            f"QLineEdit {{ border: 1px solid {C['peach']}; background-color: {C['mantle']}; }}"
            f"QLineEdit:focus {{ border-color: {C['yellow']}; background-color: {C['surface0']}; }}"
        )
        self.inp_subject.textEdited.connect(self._on_subject_edited)
        v.addWidget(self.inp_subject)

        lbl_body = QLabel("Body:")
        lbl_body.setStyleSheet(f"color: {C['subtext']}; font-weight: 500;")
        v.addWidget(lbl_body)
        self.txt_body = QTextEdit()
        self.txt_body.setMinimumHeight(360)
        self.txt_body.setPlaceholderText("Select a draft from the list to preview and edit the email body…")
        self.txt_body.setStyleSheet(
            f"QTextEdit {{ border: 1px solid {C['peach']}; background-color: {C['crust']}; }}"
            f"QTextEdit:focus {{ border-color: {C['yellow']}; }}"
        )
        self.txt_body.textChanged.connect(self._on_body_edited)
        v.addWidget(self.txt_body, 1)
        return w

    def _build_bottom(self) -> QWidget:
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 6, 0, 0)
        h.addStretch()
        btn_cancel = QPushButton("Cancel")
        btn_cancel.setMinimumHeight(40)
        btn_cancel.clicked.connect(self.reject)
        self.btn_send = QPushButton("Send 0 Email(s)")
        self.btn_send.setObjectName("btnStart")
        self.btn_send.setMinimumHeight(40)
        self.btn_send.setMinimumWidth(180)
        self.btn_send.clicked.connect(self.accept)
        h.addWidget(btn_cancel)
        h.addSpacing(8)
        h.addWidget(self.btn_send)
        self._refresh_send_btn()
        return w

    def _on_row_changed(self, row: int):
        self._save_edits()
        self._load_draft(row)

    def _on_item_changed(self, item: QTableWidgetItem):
        if self._blocking or item.column() != 0:
            return
        self._drafts[item.row()].include = (
            item.checkState() == Qt.CheckState.Checked
        )
        self._refresh_send_btn()

    def _save_edits(self):
        r = self._current_row
        if r < 0 or r >= len(self._drafts):
            return
        d = self._drafts[r]
        d.subject = self.inp_subject.text().strip()
        d.body    = self.txt_body.toPlainText().strip()
        item = self.table.item(r, 2)
        if item:
            item.setText(d.subject)

    def _load_draft(self, row: int):
        if row < 0 or row >= len(self._drafts):
            return
        d  = self._drafts[row]
        co = f"  @  {d.company}" if d.company else ""
        self.lbl_to.setText(f"To:  {d.name}{co}  ‹{d.email}›")
        if d.company_context:
            snippet = d.company_context[:220].replace("\n", " ")
            ellipsis = "…" if len(d.company_context) > 220 else ""
            self.lbl_context.setText(f"🔍 {snippet}{ellipsis}")
            self.lbl_context.setVisible(True)
        else:
            self.lbl_context.setVisible(False)
        self.inp_subject.blockSignals(True)
        self.txt_body.blockSignals(True)
        self.inp_subject.setText(d.subject)
        self.txt_body.setPlainText(d.body)
        self.inp_subject.blockSignals(False)
        self.txt_body.blockSignals(False)
        self._current_row = row

    def _on_subject_edited(self, text: str):
        if self._current_row >= 0:
            self._drafts[self._current_row].subject = text
            item = self.table.item(self._current_row, 2)
            if item:
                item.setText(text)

    def _on_body_edited(self):
        if self._current_row >= 0:
            self._drafts[self._current_row].body = self.txt_body.toPlainText().strip()

    def _bulk_check(self, state: bool):
        self._blocking = True
        cs = Qt.CheckState.Checked if state else Qt.CheckState.Unchecked
        for r, d in enumerate(self._drafts):
            d.include = state
            item = self.table.item(r, 0)
            if item:
                item.setCheckState(cs)
        self._blocking = False
        self._refresh_send_btn()

    def _refresh_send_btn(self):
        n = sum(1 for d in self._drafts if d.include)
        self.btn_send.setText(f"Send {n} Email{'s' if n != 1 else ''}")
        self.btn_send.setEnabled(n > 0)

    def selected_drafts(self) -> List[DraftEmail]:
        self._save_edits()
        return [d for d in self._drafts if d.include]


# ── Campaign Thread ────────────────────────────────────────────────────────────

class CampaignThread(QThread):
    sig_log      = pyqtSignal(str, str)
    sig_progress = pyqtSignal(int, int)
    sig_status   = pyqtSignal(str)
    sig_preview  = pyqtSignal(str, str, str, str)
    sig_done     = pyqtSignal(bool, str)

    def __init__(self, cfg: Dict[str, Any], contacts: List[Dict],
                 cv_text: str, cv_path: str, parent=None,
                 drafts: Optional[List[DraftEmail]] = None):
        super().__init__(parent)
        self.cfg      = cfg
        self.contacts = contacts
        self.cv_text  = cv_text
        self.cv_path  = cv_path
        self.drafts   = drafts

        self._mutex   = QMutex()
        self._wait    = QWaitCondition()
        self._paused  = False
        self._stopped = False

        self._preview_event  = threading.Event()
        self._preview_result: Tuple[bool, str, str] = (True, "", "")
        self._log_records: List[Dict] = []
        self._researcher = CompanyResearcher()

    def pause(self):
        self._mutex.lock(); self._paused = True; self._mutex.unlock()

    def resume(self):
        self._mutex.lock()
        self._paused = False; self._wait.wakeAll()
        self._mutex.unlock()

    def stop(self):
        self._mutex.lock()
        self._stopped = True; self._paused = False; self._wait.wakeAll()
        self._mutex.unlock()

    def set_preview_result(self, send: bool, subject: str, body: str):
        self._preview_result = (send, subject, body)
        self._preview_event.set()

    def _check_paused(self) -> bool:
        self._mutex.lock()
        while self._paused and not self._stopped:
            self._wait.wait(self._mutex)
        stopped = self._stopped
        self._mutex.unlock()
        return stopped

    def _log(self, msg: str, level: str = "info", recipient: str = "",
             success: Optional[bool] = None):
        self.sig_log.emit(msg, level)
        rec: Dict[str, Any] = {
            "ts": datetime.now().isoformat(timespec="seconds"),
            "level": level, "msg": msg, "recipient": recipient,
        }
        if success is not None:
            rec["success"] = success
        self._log_records.append(rec)

    def _flush_log(self):
        try:
            existing: List = []
            if LOG_FILE.exists():
                try:
                    existing = json.loads(LOG_FILE.read_text(encoding="utf-8"))
                except Exception:
                    pass
            existing.extend(self._log_records)
            LOG_FILE.write_text(json.dumps(existing, indent=2, ensure_ascii=False), encoding="utf-8")
        except Exception as e:
            self.sig_log.emit(f"Log save error: {e}", "warn")

    def run(self):
        cfg     = self.cfg
        mailer  = EmailSender(cfg["email"], cfg["password"])
        delay_s = cfg["delay_min"] * 60
        sent = failed = 0

        if self.drafts is not None:
            total = len(self.drafts)
            self._log(f"Sending {total} pre-reviewed email(s), {cfg['delay_min']} min delay")
            self.sig_progress.emit(0, total)
            for i, draft in enumerate(self.drafts):
                if self._check_paused():
                    self._log("Stopped by user.", "warn"); break
                self.sig_status.emit(f"Sending to {draft.name} <{draft.email}>…")
                try:
                    mailer.send(
                        to=draft.email, subject=draft.subject, body=draft.body,
                        cv_path=self.cv_path if cfg["attach_cv"] else None,
                    )
                    self._log(
                        f"[{i+1}/{total}] ✓ Sent → {draft.name} <{draft.email}> | {draft.subject}",
                        "success", draft.email, True,
                    )
                    sent += 1
                except smtplib.SMTPAuthenticationError:
                    self._log("SMTP auth failed — check credentials.", "error", draft.email, False)
                    failed += 1; break
                except (smtplib.SMTPException, OSError) as exc:
                    self._log(f"[{i+1}/{total}] SMTP error: {exc}", "error", draft.email, False)
                    failed += 1
                except Exception as exc:
                    self._log(f"[{i+1}/{total}] Error: {exc}", "error", draft.email, False)
                    failed += 1
                self.sig_progress.emit(i + 1, total)
                if i < total - 1 and not self._stopped:
                    remaining = delay_s
                    while remaining > 0 and not self._stopped:
                        if self._check_paused(): break
                        tick = min(remaining, 5)
                        time.sleep(tick); remaining -= tick
                        m, s = divmod(int(remaining), 60)
                        self.sig_status.emit(f"Next email in {m}m {s:02d}s…")
            self._flush_log()
            summary = f"Done: {sent} sent, {failed} failed / {total} total."
            self._log(summary); self.sig_status.emit("Idle")
            self.sig_done.emit(failed == 0, summary)
            return

        llm   = LLMClient(cfg["provider"], cfg["api_key"], cfg.get("model", ""), llm_debug_provider=cfg.get("llm_debug_provider", False))
        total = len(self.contacts)
        sent = failed = 0

        self._log(f"Campaign started — {total} contact(s), {cfg['delay_min']} min delay")
        self.sig_progress.emit(0, total)

        for i, contact in enumerate(self.contacts):
            if self._check_paused():
                self._log("Stopped by user.", "warn")
                break

            email = get_field(contact, "email", "e-mail", "email_address", "mail")
            name  = get_field(contact, "name", "full_name", "first_name",
                              default=f"Contact {i + 1}")

            if not email:
                self._log(f"[{i+1}/{total}] No email for '{name}' — skipped", "warn")
                self.sig_progress.emit(i + 1, total)
                continue
            if not valid_email(email):
                self._log(f"[{i+1}/{total}] Invalid address '{email}' — skipped", "warn", email)
                failed += 1
                self.sig_progress.emit(i + 1, total)
                continue

            company_context = get_field(
                contact,
                "company_context",
                "description",
                "description_de_l_activite_anglais",
                default="",
            )
            company_context = sanitize_company_context(company_context)
            if cfg.get("research"):
                company = get_field(contact, "company", "organization", "employer", "firm",
                                    default="")
                website = get_field(contact, "website", "url", "homepage", "site")
                if company:
                    self.sig_status.emit(f"Researching {company}…")
                    self._log(f"[{i+1}/{total}] Researching company: {company}", "info", email)
                    researched_context = self._researcher.research(company, website)
                    if researched_context:
                        researched_context = sanitize_company_context(researched_context)
                        company_context = f"{company_context}\n\n{researched_context}".strip() if company_context else researched_context
                        self._log(
                            f"[{i+1}/{total}] Context found ({len(researched_context)} chars)",
                            "info", email,
                        )
                    else:
                        self._log(f"[{i+1}/{total}] No context found for '{company}'", "warn", email)

            self.sig_status.emit(f"Generating email for {name}…")
            self._log(f"[{i+1}/{total}] Generating → {name} <{email}>", "info", email)
            try:
                lang = "fr"
                retries = int(cfg.get("llm_retries", 2))
                subject = body = ""
                for attempt in range(1, retries + 1):
                    prompt_text = build_prompt(
                        self.cv_text, contact, cfg["goal"], company_context,
                        sender_name=cfg.get("sender_name", ""),
                        sender_email=cfg.get("sender_email", ""),
                        sender_phone=cfg.get("sender_phone", ""),
                        sender_extra=cfg.get("sender_extra", ""),
                        language=lang,
                    )
                    try:
                        raw = llm.generate(prompt_text)
                    except LLMProviderError as ple:
                        self._log(f"[{i+1}/{total}] Provider {ple.provider} returned {ple.status_code} — attempting fallbacks", "warn", email)
                        fallbacks = cfg.get("provider_fallbacks", [])
                        fallback_succeeded = False
                        for fb in fallbacks:
                            try:
                                p_name = fb.get("provider")
                                p_key = fb.get("api_key")
                                p_model = fb.get("model", "")
                                if not p_name or not p_key:
                                    continue
                                self._log(f"[{i+1}/{total}] Trying fallback provider {p_name}", "info", email)
                                fb_llm = LLMClient(p_name, p_key, p_model, llm_debug_provider=cfg.get("llm_debug_provider", False))
                                raw = fb_llm.generate(prompt_text)
                                fallback_succeeded = True
                                llm = fb_llm
                                break
                            except Exception as e:
                                self._log(f"[{i+1}/{total}] Fallback {fb.get('provider')} failed: {e}", "warn", email)
                                continue
                        if not fallback_succeeded:
                            raise
                    try:
                        subject, body = parse_llm_json(raw)
                    except Exception as e:
                        self._log(f"[{i+1}/{total}] Parse error from LLM (attempt {attempt}): {e}", "warn", email)
                        body = ""
                    if is_complete_french_email(body):
                        self._log(f"[{i+1}/{total}] LLM produced complete French body on attempt {attempt}", "info", email)
                        break
                    else:
                        self._log(f"[{i+1}/{total}] LLM output failed completeness check (attempt {attempt})", "warn", email)
                        try:
                            self._log(f"[{i+1}/{total}] Raw LLM output (truncated): {raw[:800]}", "debug", email)
                        except Exception:
                            pass
                        try:
                            if cfg.get("llm_debug"):
                                p = Path(f"llm_debug_campaign_{i+1}_attempt{attempt}.txt")
                                p.write_text(raw or "", encoding="utf-8")
                                self._log(f"[{i+1}/{total}] Full LLM output written to {p}", "info", email)
                        except Exception:
                            pass
                        if attempt < retries:
                            time.sleep(1)
                        # If subject present but body incomplete, try a focused repair prompt
                        try:
                            if subject and not is_complete_french_email(body):
                                repair_tries = 2
                                for r_i in range(1, repair_tries + 1):
                                    repair_prompt = (
                                        f"Le modèle a renvoyé ceci:\n{raw}\n\n"
                                        f"Le sujet détecté est : {subject}\n"
                                        "Le champ 'body' est manquant ou incomplet. En te basant uniquement sur le CV et le contexte fournis, retourne STRICTEMENT un objet JSON valide avec exactement les deux clés 'subject' et 'body' (aucun texte supplémentaire). Ne modifie pas le 'subject'. Le 'body' doit être en français et contenir explicitement les tokens : Portfolio, LinkedIn, GitHub, stage d'été, INSAT.\n"
                                        f"CV:\n\"\"\"\n{self.cv_text}\n\"\"\"\n"
                                        f"Contexte entreprise:\n\"\"\"\n{company_context}\n\"\"\"\n"
                                    )
                                    self._log(f"[{i+1}/{total}] Repair attempt {r_i} for body", "info", email)
                                    repair_raw = llm.generate(repair_prompt)
                                    try:
                                        subject_r, body_r = parse_llm_json(repair_raw)
                                    except Exception as e:
                                        self._log(f"[{i+1}/{total}] Parse error from repair LLM attempt {r_i}: {e}", "warn", email)
                                        body_r = ""
                                    if is_complete_french_email(body_r):
                                        body = body_r
                                        subject = subject or subject_r or subject
                                        self._log(f"[{i+1}/{total}] Repair succeeded on attempt {r_i}", "info", email)
                                        break
                                    time.sleep(1)
                        except Exception:
                            pass
                        continue
                if not is_complete_french_email(body):
                    subject, body = build_french_fallback_email(
                        self.cv_text, contact, cfg["goal"], company_context,
                        sender_name=cfg.get("sender_name", ""),
                        sender_email=cfg.get("sender_email", ""),
                        sender_phone=cfg.get("sender_phone", ""),
                        sender_extra=cfg.get("sender_extra", ""),
                    )
                    self._log(
                        f"[{i+1}/{total}] Incomplete French body after {retries} attempts — using fallback French draft",
                        "warn", email,
                    )
                
                subject = FIXED_EMAIL_SUBJECT
                sender_name = resolve_sender_name(
                    self.cv_text,
                    cfg.get("sender_name", ""),
                    cfg.get("sender_email", ""),
                )
                # Ensure portfolio / links are present before appending signature
                body = ensure_portfolio_links(body, cfg.get("sender_extra", ""))
                signature = build_signature_block(
                    sender_name,
                    cfg.get("sender_email", ""),
                    cfg.get("sender_phone", ""),
                )
                body = enforce_signature(body, signature)
            except Exception as exc:
                self._log(f"[{i+1}/{total}] LLM error: {exc}", "error", email, False)
                failed += 1
                self.sig_progress.emit(i + 1, total)
                continue

            time.sleep(3)

            if cfg.get("preview"):
                self._preview_event.clear()
                self.sig_preview.emit(name, email, subject, body)
                self._preview_event.wait(timeout=300)
                send, subject, body = self._preview_result
                if not send:
                    self._log(f"[{i+1}/{total}] Skipped by user: {email}", "warn", email)
                    self.sig_progress.emit(i + 1, total)
                    continue

            self.sig_status.emit(f"Sending to {name} <{email}>…")
            try:
                mailer.send(
                    to=email, subject=subject, body=body,
                    cv_path=self.cv_path if cfg["attach_cv"] else None,
                )
                self._log(
                    f"[{i+1}/{total}] ✓ Sent → {name} <{email}> | {subject}",
                    "success", email, True,
                )
                sent += 1
            except smtplib.SMTPAuthenticationError:
                self._log("SMTP auth failed — check credentials.", "error", email, False)
                failed += 1
                break
            except (smtplib.SMTPException, OSError) as exc:
                self._log(f"[{i+1}/{total}] SMTP error: {exc}", "error", email, False)
                failed += 1
            except Exception as exc:
                self._log(f"[{i+1}/{total}] Send error: {exc}", "error", email, False)
                failed += 1

            self.sig_progress.emit(i + 1, total)

            if i < total - 1 and not self._stopped:
                remaining = delay_s
                while remaining > 0 and not self._stopped:
                    if self._check_paused():
                        break
                    tick = min(remaining, 5)
                    time.sleep(tick)
                    remaining -= tick
                    m, s = divmod(int(remaining), 60)
                    self.sig_status.emit(f"Next email in {m}m {s:02d}s…")

        self._flush_log()
        summary = f"Done: {sent} sent, {failed} failed / {total} total."
        self._log(summary)
        self.sig_status.emit("Idle")
        self.sig_done.emit(failed == 0, summary)


# ── Main Window ────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} — Email Outreach")
        self.setMinimumSize(1020, 700)
        self.resize(1200, 800)

        self._cv_path   = ""
        self._csv_path  = ""
        self._contacts: List[Dict] = []
        self._cv_text   = ""
        self._thread: Optional[CampaignThread] = None
        self._gen_thread: Optional[GenerationThread] = None
        self._cv_info_thread: Optional[CVInfoThread] = None
        self._is_paused = False
        self._sender_name  = ""
        self._sender_phone = ""
        self._sender_extra = ""
        self._contacts_language = ""
        self._cv_llm_extracted = False

        self._build_ui()
        self._load_local_config()

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        vbox = QVBoxLayout(root)
        vbox.setSpacing(0)
        vbox.setContentsMargins(0, 0, 0, 0)

        vbox.addWidget(self._ui_header())

        body = QWidget()
        body_lay = QVBoxLayout(body)
        body_lay.setSpacing(10)
        body_lay.setContentsMargins(14, 12, 14, 10)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(1)
        splitter.addWidget(self._ui_left())
        splitter.addWidget(self._ui_right())
        splitter.setSizes([420, 680])
        body_lay.addWidget(splitter, 1)
        body_lay.addWidget(self._ui_progress_bar())

        vbox.addWidget(body, 1)

    def _ui_header(self) -> QWidget:
        w = QWidget()
        w.setFixedHeight(68)
        w.setStyleSheet(
            f"background-color: {C['mantle']};"
            f"border-bottom: 1px solid {C['surface0']};"
        )
        h = QHBoxLayout(w)
        h.setContentsMargins(20, 0, 20, 0)

        # Title group
        col = QVBoxLayout()
        col.setSpacing(2)
        title = QLabel("Outreach Console")
        title.setObjectName("lblTitle")
        title.setFont(QFont("Segoe UI Variable", 18, QFont.Weight.Bold))
        title.setStyleSheet(f"font-size: 20px; font-weight: 700; color: {C['text']}; background: transparent;")
        sub = QLabel("Personalized email generation, review, and sending")
        sub.setObjectName("lblSubtitle")
        sub.setStyleSheet(f"font-size: 11px; color: {C['overlay0']}; background: transparent;")
        col.addWidget(title)
        col.addWidget(sub)
        h.addLayout(col)
        h.addStretch()

        # Status pill
        self.lbl_status_pill = QLabel("Ready")
        self.lbl_status_pill.setObjectName("lblBadge")
        self.lbl_status_pill.setStyleSheet(
            f"background-color: {C['surface0']}; border: 1px solid {C['surface1']};"
            f"border-radius: 12px; padding: 4px 14px; color: {C['overlay2']}; font-size: 12px;"
        )
        h.addWidget(self.lbl_status_pill)
        h.addSpacing(12)

        self.btn_test = QPushButton("Test SMTP")
        self.btn_test.setObjectName("btnTest")
        self.btn_test.setFixedHeight(36)
        self.btn_test.clicked.connect(self._send_test)
        h.addWidget(self.btn_test)
        return w

    # ── Left panel ────────────────────────────────────────────────────────────
    def _ui_left(self) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        inner = QWidget()
        lay = QVBoxLayout(inner)
        lay.setSpacing(12)
        lay.setContentsMargins(2, 2, 12, 4)
        lay.addWidget(self._grp_files())
        lay.addWidget(self._grp_config())
        lay.addWidget(self._grp_settings())
        lay.addWidget(self._grp_controls())
        lay.addStretch()

        scroll.setWidget(inner)
        return scroll

    def _file_row(self, label: str, lbl_attr: str, slot) -> QWidget:
        row = QWidget()
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(8)
        lbl_label = QLabel(label)
        lbl_label.setStyleSheet(f"color: {C['overlay2']}; font-size: 12px; font-weight: 600;")
        lbl_label.setFixedWidth(88)
        h.addWidget(lbl_label)
        lbl = QLabel("No file selected")
        lbl.setStyleSheet(f"color: {C['overlay0']}; font-size: 12px;")
        lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        setattr(self, lbl_attr, lbl)
        h.addWidget(lbl, 1)
        btn = QPushButton("Browse")
        btn.setObjectName("btnBrowse")
        btn.setFixedWidth(80)
        btn.setFixedHeight(30)
        btn.clicked.connect(slot)
        h.addWidget(btn)
        return row

    def _grp_files(self) -> QGroupBox:
        grp = QGroupBox("Files")
        lay = QVBoxLayout(grp)
        lay.setSpacing(10)
        lay.setContentsMargins(14, 10, 14, 14)
        lay.addWidget(self._file_row("CV / Resume:", "lbl_cv",  self._browse_cv))

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"background: {C['surface0']}; max-height: 1px;")
        lay.addWidget(sep)

        lay.addWidget(self._file_row("Contacts CSV:", "lbl_csv", self._browse_csv))

        self.lbl_count = QLabel("")
        self.lbl_count.setStyleSheet(f"color: {C['green']}; font-size: 11px; padding-left: 96px;")
        lay.addWidget(self.lbl_count)
        return grp

    def _grp_config(self) -> QGroupBox:
        grp = QGroupBox("Configuration")
        form = QFormLayout(grp)
        form.setSpacing(10)
        form.setContentsMargins(14, 10, 14, 14)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.inp_goal = QLineEdit()
        self.inp_goal.setPlaceholderText("e.g. summer AI / Edge AI engineering internship")
        form.addRow("Goal:", self.inp_goal)

        self.cmb_provider = QComboBox()
        self.cmb_provider.addItems([
            "OpenAI", "Gemini", "Groq", "OpenRouter", "Anthropic",
            "Mistral", "Together AI", "Cohere",
        ])
        self.cmb_provider.currentTextChanged.connect(self._on_provider)
        form.addRow("LLM Provider:", self.cmb_provider)

        self.inp_model = QLineEdit()
        self.inp_model.setPlaceholderText("openai/gpt-oss-120b")
        self.inp_model.setText("openai/gpt-oss-120b")
        form.addRow("Model:", self.inp_model)

        self.inp_api_key = QLineEdit()
        self.inp_api_key.setEchoMode(QLineEdit.EchoMode.Password)
        self.inp_api_key.setPlaceholderText("Paste your API key…")
        form.addRow("API Key:", self.inp_api_key)

        sep_lbl = QLabel("SMTP")
        sep_lbl.setStyleSheet(f"color: {C['surface2']}; font-size: 11px;")
        form.addRow("", sep_lbl)

        self.inp_sender = QLineEdit()
        self.inp_sender.setPlaceholderText("your@gmail.com")
        form.addRow("Gmail:", self.inp_sender)

        self.inp_password = QLineEdit()
        self.inp_password.setEchoMode(QLineEdit.EchoMode.Password)
        self.inp_password.setPlaceholderText("16-char App Password")
        form.addRow("App Password:", self.inp_password)

        return grp

    def _grp_settings(self) -> QGroupBox:
        grp = QGroupBox("Settings")
        form = QFormLayout(grp)
        form.setSpacing(10)
        form.setContentsMargins(14, 10, 14, 14)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        delay_w = QWidget()
        dh = QHBoxLayout(delay_w)
        dh.setContentsMargins(0, 0, 0, 0)
        self.spn_delay = QSpinBox()
        self.spn_delay.setRange(1, 120)
        self.spn_delay.setValue(10)
        self.spn_delay.setSuffix(" min")
        self.spn_delay.setFixedWidth(100)
        dh.addWidget(self.spn_delay)
        dh.addStretch()
        form.addRow("Send delay:", delay_w)

        self.chk_attach   = QCheckBox("Attach CV to every email")
        self.chk_preview  = QCheckBox("Preview each email before sending")
        self.chk_research = QCheckBox("Research each company online before writing")
        form.addRow("", self.chk_attach)
        form.addRow("", self.chk_preview)
        form.addRow("", self.chk_research)
        return grp

    def _grp_controls(self) -> QGroupBox:
        grp = QGroupBox("Campaign")
        outer = QVBoxLayout(grp)
        outer.setContentsMargins(14, 10, 14, 14)
        outer.setSpacing(8)

        top_row = QHBoxLayout()
        self.btn_preview_all = QPushButton("Generate Drafts")
        self.btn_preview_all.setObjectName("btnGenerate")
        self.btn_preview_all.setMinimumHeight(44)
        self.btn_preview_all.setToolTip(
            "Generate all emails with the LLM first, review them, then send."
        )
        self.btn_preview_all.clicked.connect(self._generate_drafts)

        self.btn_start = QPushButton("Start Campaign")
        self.btn_start.setObjectName("btnStart")
        self.btn_start.setMinimumHeight(44)
        self.btn_start.setToolTip("Generate and send emails one by one (no preview).")
        self.btn_start.clicked.connect(self._start)
        top_row.addWidget(self.btn_preview_all)
        top_row.addWidget(self.btn_start)
        outer.addLayout(top_row)

        bot_row = QHBoxLayout()
        self.btn_pause = QPushButton("Pause")
        self.btn_pause.setObjectName("btnPause")
        self.btn_pause.setMinimumHeight(36)
        self.btn_pause.setEnabled(False)
        self.btn_pause.clicked.connect(self._toggle_pause)

        self.btn_stop = QPushButton("Stop")
        self.btn_stop.setObjectName("btnStop")
        self.btn_stop.setMinimumHeight(36)
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self._stop)
        bot_row.addWidget(self.btn_pause)
        bot_row.addWidget(self.btn_stop)
        outer.addLayout(bot_row)

        return grp

    # ── Right panel ───────────────────────────────────────────────────────────
    def _ui_right(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setSpacing(12)
        v.setContentsMargins(10, 2, 2, 4)

        # Status bar
        sr = QWidget()
        sr.setStyleSheet(
            f"background-color: {C['mantle']}; border: 1px solid {C['surface0']}; border-radius: 8px;"
        )
        sh = QHBoxLayout(sr)
        sh.setContentsMargins(14, 8, 14, 8)
        lbl = QLabel("Status")
        lbl.setStyleSheet(f"font-weight: 700; color: {C['overlay2']}; font-size: 12px; background: transparent;")
        sh.addWidget(lbl)
        self.lbl_status = QLabel("Ready. Review inputs, generate drafts, or start a campaign.")
        self.lbl_status.setObjectName("lblStatus")
        self.lbl_status.setWordWrap(True)
        self.lbl_status.setStyleSheet(f"color: {C['text']}; background: transparent;")
        sh.addWidget(self.lbl_status, 1)
        v.addWidget(sr)

        # Log
        log_grp = QGroupBox("Activity Log")
        lv = QVBoxLayout(log_grp)
        lv.setContentsMargins(10, 10, 10, 10)
        lv.setSpacing(8)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        lv.addWidget(self.log_view)

        footer = QHBoxLayout()
        self.lbl_log_count = QLabel("0 entries")
        self.lbl_log_count.setStyleSheet(f"color: {C['overlay0']}; font-size: 11px;")
        footer.addWidget(self.lbl_log_count)
        footer.addStretch()
        btn_clear = QPushButton("Clear")
        btn_clear.setFixedWidth(72)
        btn_clear.setFixedHeight(28)
        btn_clear.clicked.connect(self._clear_log)
        footer.addWidget(btn_clear)
        lv.addLayout(footer)
        v.addWidget(log_grp, 1)
        return w

    def _ui_progress_bar(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 2, 0, 0)
        v.setSpacing(0)
        self.progress = QProgressBar()
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.progress.setFormat("Ready")
        self.progress.setFixedHeight(26)
        v.addWidget(self.progress)
        return w

    # ── Slots ─────────────────────────────────────────────────────────────────
    def _load_local_config(self):
        if not CONFIG_FILE.exists():
            return

        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception as e:
            self._append_log(f"Could not load {CONFIG_FILE.name}: {e}", "warn")
            return

        provider = str(cfg.get("provider", "")).strip()
        if provider:
            provider_map = {
                self.cmb_provider.itemText(i).lower(): self.cmb_provider.itemText(i)
                for i in range(self.cmb_provider.count())
            }
            normalized = provider_map.get(provider.lower())
            if normalized:
                self.cmb_provider.setCurrentText(normalized)
            else:
                self._append_log(f"Unknown provider in {CONFIG_FILE.name}: {provider}", "warn")

        self._set_text_from_config(self.inp_goal, cfg, "goal")
        self._set_text_from_config(self.inp_model, cfg, "model")
        self._set_text_from_config(self.inp_api_key, cfg, "api_key")
        self._set_text_from_config(self.inp_sender, cfg, "gmail_address", "email")
        self._set_text_from_config(self.inp_password, cfg, "gmail_app_password", "password")

        if "delay_minutes" in cfg:
            try:
                self.spn_delay.setValue(int(cfg["delay_minutes"]))
            except (TypeError, ValueError):
                self._append_log(f"Invalid delay_minutes in {CONFIG_FILE.name}", "warn")

        self._set_checkbox_from_config(self.chk_attach, cfg, "attach_cv")
        self._set_checkbox_from_config(self.chk_preview, cfg, "preview")
        self._set_checkbox_from_config(self.chk_research, cfg, "research")

        cv_path = str(cfg.get("cv_path", "")).strip()
        if cv_path:
            self._load_cv_file(cv_path)

        csv_path = str(cfg.get("contacts_csv", cfg.get("csv_path", ""))).strip()
        if csv_path:
            self._load_contacts_file(csv_path)

        self._append_log(f"Loaded defaults from {CONFIG_FILE.name}", "success")

    def _set_text_from_config(self, widget: QLineEdit, cfg: Dict[str, Any],
                              *keys: str):
        for key in keys:
            value = cfg.get(key)
            if value is not None and str(value).strip():
                widget.setText(str(value).strip())
                return

    def _set_checkbox_from_config(self, widget: QCheckBox, cfg: Dict[str, Any],
                                  key: str):
        if key in cfg:
            widget.setChecked(bool(cfg[key]))

    def _load_cv_file(self, path: str):
        if not os.path.isfile(path):
            self._append_log(f"CV file not found: {path}", "warn")
            return
        self._cv_path = path
        name = Path(path).name
        self.lbl_cv.setText(name)
        self.lbl_cv.setStyleSheet(f"color: {C['text']}; font-size: 12px;")
        self._cv_text = parse_cv(path)
        if self._cv_text.startswith("["):
            self._append_log(f"CV parse warning: {self._cv_text}", "warn")
        else:
            self._append_log(f"CV loaded: {name} ({len(self._cv_text)} chars)", "success")
            self._cv_llm_extracted = False
            self._start_cv_extraction()

    def _load_contacts_file(self, path: str):
        if not os.path.isfile(path):
            self._append_log(f"Contacts CSV not found: {path}", "warn")
            return False
        try:
            self._csv_path = path
            self._contacts = load_contacts(path)
            self._contacts_language = _contacts_language_for_path(path)
            name = Path(path).name
            self.lbl_csv.setText(name)
            self.lbl_csv.setStyleSheet(f"color: {C['text']}; font-size: 12px;")
            n = len(self._contacts)
            self.lbl_count.setText(f"✓  {n} contact{'s' if n != 1 else ''} loaded")
            self.progress.setRange(0, n)
            self.progress.setValue(0)
            self.progress.setFormat(f"0 / {n}")
            self._append_log(f"CSV loaded: {name} — {n} contacts", "success")
            return True
        except Exception as e:
            self._append_log(f"Failed to load CSV: {e}", "error")
            return False

    def _browse_cv(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select CV / Resume", "",
            "Documents (*.pdf *.docx *.doc);;All Files (*)",
        )
        if not path:
            return
        self._load_cv_file(path)

    def _browse_csv(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Contacts CSV", "",
            "CSV Files (*.csv);;All Files (*)",
        )
        if not path:
            return
        try:
            self._csv_path = path
            self._contacts = load_contacts(path)
            self._contacts_language = _contacts_language_for_path(path)
            name = Path(path).name
            self.lbl_csv.setText(name)
            self.lbl_csv.setStyleSheet(f"color: {C['text']}; font-size: 12px;")
            n = len(self._contacts)
            self.lbl_count.setText(f"✓  {n} contact{'s' if n != 1 else ''} loaded")
            self.progress.setRange(0, n)
            self.progress.setValue(0)
            self.progress.setFormat(f"0 / {n}")
            self._append_log(f"CSV loaded: {name} — {n} contacts", "success")
        except Exception as e:
            QMessageBox.critical(self, "CSV Error", f"Failed to load CSV:\n{e}")

    def _start_cv_extraction(self):
        if self._cv_info_thread and self._cv_info_thread.isRunning():
            return
        # Determine whether provider debug dumps are enabled in config.local.json
        debug_flag = False
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    debug_flag = bool(cfg.get("llm_debug_provider", False))
            except Exception:
                debug_flag = False

        self._cv_info_thread = CVInfoThread(
            self._cv_text,
            provider=self.cmb_provider.currentText(),
            api_key=self.inp_api_key.text().strip(),
            model=self.inp_model.text().strip(),
            llm_debug_provider=debug_flag,
        )
        self._cv_info_thread.sig_done.connect(self._on_cv_info)
        self._cv_info_thread.start()

    def _on_cv_info(self, info: dict):
        self._cv_llm_extracted = True
        self._sender_name  = info.get("name",  "").strip()
        self._sender_phone = info.get("phone", "").strip()
        self._sender_extra = info.get("extra", "").strip()

        extracted_email = info.get("email", "").strip()
        if extracted_email and not self.inp_sender.text().strip():
            self.inp_sender.setText(extracted_email)

        found = [k for k in ("name", "phone", "email", "extra") if info.get(k)]
        if found:
            self._append_log(
                f"Sender info extracted from CV: {', '.join(found)}", "success"
            )
        else:
            self._append_log("Could not extract sender info from CV (LLM will derive from CV text).", "warn")

    _PROVIDER_DEFAULTS = {
        "OpenAI":      "openai/gpt-oss-120b",
        "Gemini":      "gemini-2.0-flash",
        "Groq":        "llama-3.3-70b-versatile",
        "OpenRouter":  "google/gemini-2.0-flash-exp:free",
        "Anthropic":   "claude-3-5-haiku-20241022",
        "Mistral":     "mistral-small-latest",
        "Together AI": "meta-llama/Llama-3.3-70B-Instruct-Turbo",
        "Cohere":      "command-r-plus-08-2024",
    }

    def _on_provider(self, text: str):
        default = self._PROVIDER_DEFAULTS.get(text, "")
        self.inp_model.setPlaceholderText(default)
        if not self.inp_model.text().strip():
            self.inp_model.setText(default)

    def _validate(self) -> Optional[str]:
        if not self._cv_path:
            return "Please select a CV / Resume file."
        if not self._contacts:
            return "Please load a contacts CSV file."
        if not self.inp_goal.text().strip():
            return "Please enter the email goal."
        if not self.inp_api_key.text().strip():
            return "Please enter the LLM API key."
        if not self.inp_sender.text().strip():
            return "Please enter your Gmail address."
        if not valid_email(self.inp_sender.text().strip()):
            return "The Gmail address is not valid."
        if not self.inp_password.text().strip():
            return "Please enter the Gmail App Password."
        return None

    def _build_cfg(self) -> Dict[str, Any]:
        cfg = {
            "goal":         self.inp_goal.text().strip(),
            "provider":     self.cmb_provider.currentText(),
            "api_key":      self.inp_api_key.text().strip(),
            "model":        self.inp_model.text().strip(),
            "email":        self.inp_sender.text().strip(),
            "password":     self.inp_password.text().strip(),
            "delay_min":    self.spn_delay.value(),
            "attach_cv":    self.chk_attach.isChecked(),
            "preview":      self.chk_preview.isChecked(),
            "research":     self.chk_research.isChecked(),
            "sender_name":  self._sender_name,
            "sender_email": self.inp_sender.text().strip(),
            "sender_phone": self._sender_phone,
            "sender_extra": self._sender_extra,
            "language":     self._contacts_language,
        }
        # If local config file contains debug flags, include them
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    file_cfg = json.load(f)
                    cfg["llm_debug"] = bool(file_cfg.get("llm_debug", False))
                    cfg["llm_debug_provider"] = bool(file_cfg.get("llm_debug_provider", False))
                    # allow configuring number of LLM attempts per contact
                    try:
                        cfg["llm_retries"] = int(file_cfg.get("llm_retries", 4))
                    except Exception:
                        cfg["llm_retries"] = 4
                    # optional provider fallback list: [{"provider":"Groq","api_key":"...","model":"..."}, ...]
                    cfg["provider_fallbacks"] = file_cfg.get("provider_fallbacks", [])
            except Exception:
                cfg["llm_debug"] = False
                cfg["llm_debug_provider"] = False
        else:
            cfg["llm_debug"] = False
            cfg["llm_debug_provider"] = False
            cfg["llm_retries"] = 4
            cfg["provider_fallbacks"] = []

        return cfg

    def _ensure_cv_extracted(self):
        if self._cv_text and not self._cv_llm_extracted and self.inp_api_key.text().strip():
            self._append_log("Extracting sender info from CV via LLM…", "info")
            self._start_cv_extraction()
            self._append_log("CV extraction is running in the background; campaign will start now.", "info")

    def _start(self):
        err = self._validate()
        if err:
            QMessageBox.warning(self, "Missing Input", err)
            return

        self._ensure_cv_extracted()
        cfg = self._build_cfg()
        self.progress.setValue(0)
        self._is_paused = False
        self._set_running(True)

        self._thread = CampaignThread(
            cfg, self._contacts, self._cv_text, self._cv_path, self
        )
        self._thread.sig_log.connect(self._on_log)
        self._thread.sig_progress.connect(self._on_progress)
        self._thread.sig_status.connect(self._on_status)
        self._thread.sig_preview.connect(self._on_preview)
        self._thread.sig_done.connect(self._on_done)
        self._thread.start()

    def _toggle_pause(self):
        if not self._thread:
            return
        if self._is_paused:
            self._thread.resume()
            self._is_paused = False
            self.btn_pause.setText("Pause")
            self._on_status("Resumed…")
        else:
            self._thread.pause()
            self._is_paused = True
            self.btn_pause.setText("Resume")
            self._on_status("Paused — click Resume to continue")

    def _generate_drafts(self):
        err = self._validate()
        if err:
            QMessageBox.warning(self, "Missing Input", err)
            return

        self._ensure_cv_extracted()
        cfg = self._build_cfg()
        cfg["preview"] = False
        self.progress.setValue(0)
        self._set_running(True)
        self._on_status("Generating drafts — please wait…")

        self._gen_thread = GenerationThread(cfg, self._contacts, self._cv_text, self)
        self._gen_thread.sig_log.connect(self._on_log)
        self._gen_thread.sig_progress.connect(self._on_progress)
        self._gen_thread.sig_status.connect(self._on_status)
        self._gen_thread.sig_done.connect(self._on_gen_done)
        self._gen_thread.start()

    def _on_gen_done(self, drafts: list):
        self._set_running(False)
        self._on_status("Idle")
        if not drafts:
            QMessageBox.warning(self, "No Drafts", "No valid drafts were generated.")
            return

        dlg = PreviewAllDialog(drafts, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            selected = dlg.selected_drafts()
            if selected:
                self._send_drafts(selected)

    def _send_drafts(self, drafts: List[DraftEmail]):
        cfg = self._build_cfg()
        cfg["preview"] = False
        cfg["research"] = False
        self.progress.setRange(0, len(drafts))
        self.progress.setValue(0)
        self._set_running(True)

        self._thread = CampaignThread(
            cfg, [], self._cv_text, self._cv_path, self, drafts=drafts
        )
        self._thread.sig_log.connect(self._on_log)
        self._thread.sig_progress.connect(self._on_progress)
        self._thread.sig_status.connect(self._on_status)
        self._thread.sig_done.connect(self._on_done)
        self._thread.start()

    def _stop(self):
        if self._thread:
            self._thread.stop()
            self._thread.set_preview_result(False, "", "")
        if self._gen_thread:
            self._gen_thread.stop()
        self._on_status("Stopping…")

    # ── Thread callbacks ──────────────────────────────────────────────────────
    def _on_log(self, msg: str, level: str):
        self._append_log(msg, level)

    def _on_progress(self, current: int, total: int):
        self.progress.setRange(0, total)
        self.progress.setValue(current)
        self.progress.setFormat(f"{current} / {total}")

    def _on_status(self, msg: str):
        self.lbl_status.setText(msg)
        self.lbl_status_pill.setText(msg[:48])

    def _on_preview(self, name: str, email: str, subject: str, body: str):
        dlg = PreviewDialog(name, email, subject, body, self)
        result = dlg.exec()
        if self._thread:
            if result == QDialog.DialogCode.Accepted:
                self._thread.set_preview_result(True, dlg.subject, dlg.body)
            else:
                self._thread.set_preview_result(False, "", "")

    def _on_done(self, success: bool, summary: str):
        self._set_running(False)
        self._append_log(summary, "success" if success else "warn")
        self._on_status("Idle")
        QMessageBox.information(self, "Campaign Complete", summary)

    # ── Test email ────────────────────────────────────────────────────────────
    def _send_test(self):
        sender = self.inp_sender.text().strip()
        pwd    = self.inp_password.text().strip()
        if not sender or not pwd:
            QMessageBox.warning(
                self, "Missing Credentials",
                "Enter your Gmail address and App Password first.",
            )
            return
        try:
            EmailSender(sender, pwd).send(
                to=sender,
                subject="[Test] Email Outreach App",
                body=(
                    "This is a test email from your Email Outreach Automation app.\n\n"
                    "SMTP credentials are working correctly!"
                ),
            )
            self._append_log(f"Test email sent to {sender}", "success")
            QMessageBox.information(self, "Test Email Sent",
                                    f"Test email delivered to {sender}")
        except Exception as e:
            self._append_log(f"Test email failed: {e}", "error")
            QMessageBox.critical(self, "Test Failed", str(e))

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _set_running(self, running: bool):
        self.btn_preview_all.setEnabled(not running)
        self.btn_start.setEnabled(not running)
        self.btn_pause.setEnabled(running)
        self.btn_stop.setEnabled(running)
        if not running:
            self.btn_pause.setText("Pause")
            self._is_paused = False
            self.lbl_status_pill.setStyleSheet(
                f"background-color: {C['surface0']}; border: 1px solid {C['surface1']};"
                f"border-radius: 12px; padding: 4px 14px; color: {C['overlay2']}; font-size: 12px;"
            )
        else:
            self.lbl_status_pill.setStyleSheet(
                f"background-color: rgba(103,211,145,0.12); border: 1px solid {C['accent']};"
                f"border-radius: 12px; padding: 4px 14px; color: {C['accent']}; font-size: 12px;"
            )

    _log_entry_count = 0

    def _clear_log(self):
        self.log_view.clear()
        self._log_entry_count = 0
        self.lbl_log_count.setText("0 entries")

    def _append_log(self, msg: str, level: str = "info"):
        colors = {
            "info":    C['text'],
            "success": C['green'],
            "warn":    C['yellow'],
            "error":   C['red'],
        }
        icons = {
            "info":    "·",
            "success": "✓",
            "warn":    "⚠",
            "error":   "✗",
        }
        color = colors.get(level, C['text'])
        icon  = icons.get(level, "·")
        ts = datetime.now().strftime("%H:%M:%S")
        safe = msg.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        html = (
            f'<span style="color:{C["surface2"]}">[{ts}]</span> '
            f'<span style="color:{color}">{icon} {safe}</span>'
        )
        self.log_view.append(html)
        cur = self.log_view.textCursor()
        cur.movePosition(QTextCursor.MoveOperation.End)
        self.log_view.setTextCursor(cur)
        self._log_entry_count += 1
        self.lbl_log_count.setText(f"{self._log_entry_count} entr{'y' if self._log_entry_count == 1 else 'ies'}")

    def closeEvent(self, event):
        if self._cv_info_thread and self._cv_info_thread.isRunning():
            self._cv_info_thread.wait(2000)
        active = (
            (self._thread and self._thread.isRunning()) or
            (self._gen_thread and self._gen_thread.isRunning())
        )
        if active:
            reply = QMessageBox.question(
                self, "Task Running",
                "A generation or campaign is in progress. Stop it and close?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.No:
                event.ignore()
                return
            if self._gen_thread:
                self._gen_thread.stop()
                self._gen_thread.wait(3000)
            if self._thread:
                self._thread.stop()
                self._thread.set_preview_result(False, "", "")
                self._thread.wait(5000)
        event.accept()


# ── Entry Point ────────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setStyleSheet(DARK_STYLE)
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
