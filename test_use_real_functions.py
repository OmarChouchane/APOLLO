import json
import sys
import types
from pathlib import Path

# Create lightweight dummy PyQt6 modules to allow importing the app in headless test
if 'PyQt6' not in sys.modules:
    pyqt6 = types.ModuleType('PyQt6')
    qtwidgets = types.ModuleType('PyQt6.QtWidgets')
    qtcore = types.ModuleType('PyQt6.QtCore')
    qtgui = types.ModuleType('PyQt6.QtGui')

    # Provide minimal placeholders used at import-time
    class Dummy:
        def __init__(self, *a, **k):
            pass

    for mod in (qtwidgets, qtcore, qtgui):
        for name in ('QApplication', 'QMainWindow', 'QWidget', 'QVBoxLayout', 'QHBoxLayout',
                     'QPushButton', 'QLabel', 'QLineEdit', 'QTextEdit', 'QProgressBar',
                     'QFileDialog', 'QFrame', 'QCheckBox', 'QSpinBox', 'QGroupBox',
                     'QMessageBox', 'QDialog', 'QDialogButtonBox', 'QSplitter', 'QComboBox',
                     'QScrollArea', 'QFormLayout', 'QSizePolicy', 'QTableWidget', 'QTableWidgetItem',
                     'QHeaderView', 'Qt', 'QThread', 'pyqtSignal', 'QMutex', 'QWaitCondition',
                     'QFont', 'QTextCursor', 'QPalette', 'QColor'):
            setattr(mod, name, Dummy)

    sys.modules['PyQt6'] = pyqt6
    sys.modules['PyQt6.QtWidgets'] = qtwidgets
    sys.modules['PyQt6.QtCore'] = qtcore
    sys.modules['PyQt6.QtGui'] = qtgui

from email_outreach_app import (
    load_contacts, parse_cv, build_prompt, detect_language_from_country,
    build_signature_block, parse_llm_json
)

CSV_PATH = 'dummy_contacts.csv'
CONFIG_PATH = Path("config.local.json")

def load_config():
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding='utf-8'))
    return {}


def mock_llm_generate(prompt: str, sender_label: str, contact: dict, cv_text: str, cfg: dict):
    # Build a top-1% mocked email following the rules in build_prompt and your example format
    company = contact.get('company', contact.get('Company', ''))
    sector = contact.get('sector', contact.get('Sector', ''))
    description = contact.get('description', contact.get('Description', ''))

    # Parse strengths from tech keywords
    techs_raw = contact.get('tech_keywords') or contact.get('Tech Keywords') or contact.get('tech keywords','')
    techs = [t.strip() for t in (techs_raw or '').replace(';', ',').split(',') if t.strip()]
    strengths = techs[:3]

    subject = cfg.get('goal') or 'Summer Internship Application – DevOps / Cloud / Software Development'

    # Build a realistic top-tier email following your example format
    opening = f"Dear {company} team,"

    p1 = "I hope you are doing well.\n\nI'm reaching out to express my strong interest in a summer internship opportunity in DevOps / Cloud / Software Development."

    # Personalized paragraph about why the company matters
    company_interest = ""
    if sector:
        sector_lower = sector.lower()
        if 'saas' in sector_lower or 'platform' in sector_lower:
            company_interest = f"What particularly attracts me to {company} is your focus on building scalable SaaS platforms combining engineering excellence with cloud infrastructure. Your mission to solve complex business problems through technology resonates strongly with how I approach systems design and DevOps."
        elif 'fintech' in sector_lower:
            company_interest = f"What attracts me to {company} is your engineering-driven approach to fintech: building secure, reliable systems that handle real-world complexity. The intersection of high-performance infrastructure and financial systems is precisely where I want to deepen my expertise."
        elif 'e-commerce' in sector_lower or 'marketplace' in sector_lower:
            company_interest = f"I'm impressed by {company}'s commitment to scaling technology for millions of users. Your approach to solving logistics and data challenges through engineering fascinates me — that's exactly the kind of impact-driven environment where I want to contribute."
        else:
            company_interest = f"What attracts me to {company} in the {sector} space is your clear focus on engineering excellence combined with real business impact. That's the kind of environment where I thrive."
    else:
        company_interest = f"I'm impressed by {company}'s engineering culture and how you combine cloud infrastructure with practical DevOps practices. That's exactly where I want to make my mark."

    p2_intro = "Through my academic projects and self-directed learning, I've focused on the intersection of software engineering and infrastructure. I've built hands-on experience with " + ", ".join(strengths) + " and have designed systems emphasizing reliability, automation, and scalability. I'm drawn to the challenge of making complex systems efficient and understandable—from containerization to CI/CD pipelines to cloud deployments."

    p3 = f"I'm currently seeking a summer internship in DevOps / Cloud / Software Development, and I'm also very open to an extended internship or part-time arrangement if that aligns with your team's needs. Joining {company} would give me the opportunity to grow within your engineering culture and contribute meaningfully from day one."

    closing = "Please find my CV attached. I'd welcome the chance to discuss how I could contribute to your team.\n\n"

    # Portfolio/LinkedIn/GitHub info
    social_info = "Portfolio: portfolio-omarchouchanes-projects.vercel.app\nLinkedIn: linkedin.com/in/omar-chouchane\nGitHub: github.com/OmarChouchane\n\n"

    signature = build_signature_block(sender_label, cfg.get('gmail_address',''), cfg.get('sender_phone',''))
    body = f"{opening}\n\n{p1}\n\n{company_interest}\n\n{p2_intro}\n\n{p3}\n\n{closing}{social_info}Best regards,\n{signature}"

    # Return JSON string as the real LLM would
    out = json.dumps({"subject": subject, "body": body}, ensure_ascii=False)
    return out


if __name__ == '__main__':
    cfg = load_config()
    if not Path(CSV_PATH).exists():
        print(f"CSV not found: {CSV_PATH}")
        raise SystemExit(1)

    contacts = load_contacts(CSV_PATH)
    if not contacts:
        print("No contacts parsed from CSV")
        raise SystemExit(1)

    # Take first two rows
    sample = contacts[:2]

    # Load CV text via parse_cv if possible
    cv_text = ""
    cv_path = cfg.get('cv_path')
    if cv_path:
        try:
            cv_text = parse_cv(cv_path)
        except Exception as e:
            print(f"Failed to parse CV: {e}")
            cv_text = ""

    sender_label = cfg.get('gmail_address') or cfg.get('sender_name') or 'Omar Chouchane'

    for i, contact in enumerate(sample, start=1):
        lang = detect_language_from_country(contact)
        prompt = build_prompt(cv_text, contact, cfg.get('goal','Summer Internship Application – DevOps / Cloud / Software Development'), company_context=contact.get('description',''), sender_name=cfg.get('sender_name',''), sender_email=cfg.get('gmail_address',''), sender_phone=cfg.get('sender_phone',''), language=lang)
        print('\n' + '='*40)
        print(f'Contact #{i}: {contact.get("company") or contact.get("Company") or contact.get("name")} (detected lang: {lang})')
        print('\n--- PROMPT (truncated to 1500 chars) ---')
        print(prompt[:1500])
        print('\n--- MOCKED LLM OUTPUT ---')
        raw = mock_llm_generate(prompt, sender_label, contact, cv_text, cfg)
        subj, body = parse_llm_json(raw)
        print(json.dumps({"subject": subj, "body": body}, ensure_ascii=False, indent=2))

    print('\nDone.')
