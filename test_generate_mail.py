from email_outreach_app import build_prompt, detect_language_from_country, build_signature_block

# Sample contact (from your provided CSV rows)
contact = {
    "company": "Personio",
    "country": "Germany",
    "sector": "HRtech",
    "description": "HR platform automating recruiting, payroll and personnel management for SMEs.",
    "linkedin": "https://www.linkedin.com/company/personio/",
    "tech_keywords": "Java; Spring Boot; React; AWS; SaaS; DevOps",
    "name": "Kamka team",
}

cv_text = (
    "Omar Chouchane\n"
    "Networks & Telecommunications Engineering student at INSAT.\n"
    "Experience building projects with Kubernetes, Docker, Terraform, CI/CD pipelines, and backend development.\n"
    "Worked on automating deployments and observability for small clusters."
)

# Decide language
lang = detect_language_from_country(contact)

# Build prompt (for inspection) --- not used by mock LLM but shown for debugging
prompt = build_prompt(cv_text, contact, "Summer Internship Application – DevOps / Cloud / Software Development",
                      company_context=contact.get('description',''), sender_name='Omar Chouchane',
                      sender_email='omar.chouchane@insat.ucar.tn', sender_phone='+216 52 834 833', language=lang)

# Mock LLM generation: construct a top-1% email following rules
def mocked_generate_email(cv_text, contact):
    name = contact.get('name') or contact.get('company') or 'There'
    company = contact.get('company', '')
    techs = [t.strip() for t in contact.get('tech_keywords','').replace(';',',').split(',') if t.strip()]
    strengths = techs[:3]
    subject = 'Summer Internship Application – DevOps / Cloud / Software Development'
    # Compose body: 3 short paragraphs
    p1 = f"Dear {company} team,\n\nI am writing to express my interest in a summer internship in DevOps / Cloud / Software Development. I am available this summer and am also open to an extended internship or a part-time arrangement if that fits your needs."
    p2 = f"Through my projects, I have focused on {', '.join(strengths)} and practical automation: I built Kubernetes-based deployments, automated infrastructure with Terraform, and implemented CI/CD pipelines to speed release cycles. I enjoy designing reliable systems and reducing manual toil."
    p3 = "Please find my CV attached — I would welcome the chance to discuss how I could contribute to your engineering efforts.\n\n"
    signature = build_signature_block('Omar Chouchane', 'omar.chouchane@insat.ucar.tn', '+216 52 834 833')
    body = p1 + "\n\n" + p2 + "\n\n" + p3 + signature
    return subject, body

if __name__ == '__main__':
    print('--- PROMPT (truncated) ---')
    print(prompt[:1000])
    print('\n--- GENERATED EMAIL ---')
    subj, body = mocked_generate_email(cv_text, contact)
    import json
    print(json.dumps({"subject": subj, "body": body}, ensure_ascii=False, indent=2))
