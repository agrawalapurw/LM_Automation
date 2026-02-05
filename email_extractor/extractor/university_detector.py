"""
University Detector
Detects if a lead is from a university or educational institution.
"""
import re
import unicodedata
from typing import Dict
import functools

import tldextract

# Optional web check
try:
    import requests
    from bs4 import BeautifulSoup
    HAS_WEB = True
except Exception:
    HAS_WEB = False

# Expanded academic keywords (EMEA-wide, normalized)
CORE_ACADEMIC_WORDS = {
    "university", "universitat", "universitaet", "universite", "universita", "universidad",
    "universidade", "universiteit", "uniwersytet", "universitatea", "universitet", "universitetet",
    "hochschule", "fachhochschule", "schule", "school", "college", "lycee", "lycée", "gymnasium",
    "polytechnic", "polytechnique", "politecnico", "politechnika", "institut", "institute",
    "academy", "akademia", "academia", "faculdade", "facultad", "faculte", "fakultaet", "fakultät",
}

COMMERCIAL_INDICATORS = {
    "gmbh", "ag", "ltd", "limited", "inc", "corp", "corporation", "llc",
    "consulting", "consultancy", "solutions", "services", "systems",
    "technologies", "tech", "group", "holding", "holdings", "partner",
    "partners", "international", "global", "worldwide",
}

ACADEMIC_TLD_HINTS = {".edu"}  # direct .edu
ACADEMIC_2LD_HINTS = {"ac", "edu", "uni"}  # e.g., ac.uk, edu.eg, uni.rostock.de (heuristic)


def normalize_text(text: str) -> str:
    if not text:
        return ""
    text = str(text)
    text = (text
            .replace("ä", "a").replace("ö", "o").replace("ü", "u")
            .replace("Ä", "A").replace("Ö", "O").replace("Ü", "U")
            .replace("ß", "ss"))
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def extract_domain(email: str) -> str:
    if not email:
        return ""
    m = re.search(r'@([A-Za-z0-9.-]+)$', str(email).strip())
    return m.group(1).lower() if m else ""


def contains_core_academic_word(text: str) -> bool:
    if not text:
        return False
    words = set(normalize_text(text).split())
    return any(w in words for w in CORE_ACADEMIC_WORDS)


def has_academic_tld(domain: str) -> bool:
    """Heuristic: .edu; or second-level like ac.uk / edu.xx / uni.xx."""
    if not domain:
        return False
    d = domain.lower()
    # direct .edu
    for t in ACADEMIC_TLD_HINTS:
        if d.endswith(t):
            return True
    # second-level patterns like ac.uk / edu.eg
    parts = d.split(".")
    if len(parts) >= 3:
        if parts[-2] in ACADEMIC_2LD_HINTS:
            return True
    return False


@functools.lru_cache(maxsize=2048)
def _web_check_is_academic(base_domain: str) -> bool:
    """Fetch homepage and look for academic keywords (very conservative)."""
    if not HAS_WEB or not base_domain:
        return False
    try:
        url = f"https://{base_domain}"
        resp = requests.get(url, timeout=2)
        if resp.status_code >= 400:
            return False
        soup = BeautifulSoup(resp.text, "html.parser")
        text = " ".join([
            soup.title.get_text(" ", strip=True) if soup.title else "",
            " ".join(h.get_text(" ", strip=True) for h in soup.find_all(["h1", "h2"])),
            " ".join(p.get_text(" ", strip=True) for p in soup.find_all("p")[:2]),
        ])
        return contains_core_academic_word(text)
    except Exception:
        return False


class UniversityDetector:
    """Detect if a lead is from a university or educational institution."""

    def __init__(self, validation_loader=None, enable_web_check: bool = False):
        self.validation_loader = validation_loader
        self.enable_web_check = enable_web_check

    def is_university(self, company: str, country: str, email: str) -> Dict[str, str]:
        domain = extract_domain(email)

        # Get registered domain (example: mail.cs.ox.ac.uk -> ox.ac.uk)
        reg_domain = ""
        if domain:
            ext = tldextract.extract(domain)
            if ext.registered_domain:
                reg_domain = ext.registered_domain.lower()

        # 1) Strong: CSV academic domains
        if self.validation_loader:
            if self.validation_loader.is_academic_domain(domain) or \
               (reg_domain and self.validation_loader.is_academic_domain(reg_domain)):
                name = self.validation_loader.get_academic_name(domain) or \
                       self.validation_loader.get_academic_name(reg_domain)
                reason = f"Academic domain in database: {reg_domain or domain}"
                if name:
                    reason += f" ({name})"
                return {"is_university": True, "reason": reason, "confidence": "high"}

            # Direct accounts override
            if self.validation_loader.is_direct_account(company):
                return {"is_university": False, "reason": f"Known commercial company: {company}", "confidence": "high"}

        # 2) Heuristic: academic TLDs and 2LDs
        if has_academic_tld(domain) or (reg_domain and has_academic_tld(reg_domain)):
            return {"is_university": True, "reason": f"Academic TLD pattern: {reg_domain or domain}", "confidence": "medium"}

        # 3) Company name contains academic keyword
        if contains_core_academic_word(company):
            return {"is_university": True, "reason": f"Company name contains academic keyword: {company}", "confidence": "medium"}

        # 4) Optional web check (slow; cached; conservative)
        if self.enable_web_check and reg_domain and _web_check_is_academic(reg_domain):
            return {"is_university": True, "reason": f"Website indicates academic institution: {reg_domain}", "confidence": "low"}

        # 5) Commercial indicators reduce likelihood
        norm_company = normalize_text(company)
        if any(token in norm_company for token in {"gmbh", "ltd", "inc", "llc", "ag"}):
            return {"is_university": False, "reason": f"Contains commercial indicators: {company}", "confidence": "medium"}

        # Default: not a university
        return {"is_university": False, "reason": "No clear university indicators", "confidence": "high"}