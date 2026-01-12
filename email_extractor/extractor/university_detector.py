import re
import unicodedata
from typing import Dict


# Free email domains
FREE_MAILERS = {
    "gmail.com", "yahoo.com", "outlook.com", "hotmail.com", "live.com",
    "aol.com", "icloud.com", "mail.com", "proton.me", "protonmail.com",
    "gmx.com", "gmx.de", "web.de", "yandex.com", "zoho.com", "qq.com"
}

# Academic TLD patterns - STRICT
ACADEMIC_TLDS = [
    ".edu", ".ac.uk", ".ac.at", ".ac.be", ".ac.il", ".ac.jp",
    ".edu.au", ".edu.cn", ".edu.sg", ".edu.hk", ".edu.tw"
]

# Single-word core academic keywords
CORE_ACADEMIC_WORDS = {
    "university", "universitat", "universitaet", "universite", "universita", "universidad", "universidade",
    "hochschule", "fachhochschule",
    "college",
    "polytechnic", "polytechnique", "politecnico",
}

# Domain keywords that strongly indicate academic institutions
STRICT_ACADEMIC_DOMAIN_TOKENS = {
    "university", "universitat", "college", "academic", "scholar", "campus", "edu"
}

# Known commercial/industrial companies (EXPANDED)
COMMERCIAL_COMPANIES = {
    # Automotive
    "airbus", "bosch", "siemens", "bmw", "mercedes", "volkswagen", "vw", "audi",
    "continental", "zf", "schaeffler", "mahle", "porsche", "daimler", "opel", "ford",
    "renault", "peugeot", "citroen", "fiat", "alfa romeo", "ferrari", "lamborghini",
    
    # Technology
    "intel", "google", "amazon", "microsoft", "sap", "oracle", "ibm", "apple",
    "samsung", "huawei", "cisco", "hp", "dell", "lenovo", "asus", "acer",
    "infineon", "nxp", "stmicroelectronics", "texas instruments", "analog devices",
    "qualcomm", "nvidia", "amd", "arm", "broadcom",
    
    # Industrial/Manufacturing
    "abb", "schneider", "philips", "basf", "bayer", "thyssenkrupp", "linde",
    "henkel", "covestro", "wacker", "evonik", "lanxess", "merck", "fresenius",
    "kuka", "trumpf", "festo", "sick", "pepperl fuchs", "balluff", "turck",
    
    # Telecom
    "deutsche telekom", "orange", "vodafone", "telefonica", "bt", "swisscom",
    "proximus", "kpn", "telecom italia", "telekom austria",
    
    # Energy
    "siemens energy", "vestas", "eon", "rwe", "engie", "edf", "iberdrola",
    "total", "shell", "bp", "eni", "equinor",
    
    # Aerospace/Defense
    "airbus defence", "thales", "leonardo", "safran", "rolls royce", "mtu",
    "diehl", "hensoldt", "rheinmetall",
    
    # Consulting/Services
    "accenture", "deloitte", "pwc", "kpmg", "ey", "mckinsey", "bcg", "bain",
    "capgemini", "atos", "sopra steria", "gft", "msg", "adesso",
    
    # Electronics/Components
    "robert bosch", "continental automotive", "valeo", "hella", "osram", "ledvance",
    "weidmuller", "phoenix contact", "wago", "murrelektronik", "harting", "lapp",
    
    # Chemicals
    "clariant", "solvay", "akzonobel", "arkema", "brenntag", "altana",
    
    # Medical/Pharma
    "roche", "novartis", "sanofi", "bayer healthcare", "fresenius medical",
    "siemens healthineers", "philips healthcare",
    
    # Distribution/Retail
    "arrow", "avnet", "future electronics", "rutronik", "distrelec", "farnell",
    "mouser", "digikey", "rs components", "wurth elektronik",
    
    # Engineering/Tech companies
    "nexus quantum", "nexus", "quantum engineering", "engineering solutions",
    "tech solutions", "innovative engineering",
    
    # Other
    "training", "consulting", "gmbh", "ag", "se", "ltd", "limited", "inc",
    "corporation", "corp", "solutions", "services", "systems", "technologies",
    "tech", "group", "holding", "partner", "international", "quantum"
}

# Commercial indicators - words that suggest it's NOT a university
# NOTE: "technical" alone is NOT here because "Technical University" is valid
COMMERCIAL_INDICATORS = {
    "gmbh", "ag", "ltd", "limited", "inc", "corp", "corporation", "llc",
    "consulting", "consultancy", "solutions", "services", "systems", 
    "technologies", "tech", "group", "holding", "holdings", "partner", 
    "partners", "international", "global", "worldwide",
    "quantum", "nexus", "innovative", "advanced", "smart", "digital",
    "engineering solutions", "engineering services", "technical services"
}

# Country-specific KNOWN universities
KNOWN_UNIVERSITIES = {
    "germany": {
        "rwth aachen", "rheinisch westfalische technische hochschule",
        "technische universitat munchen", "tum",
        "technische universitat darmstadt", "tu darmstadt",  # ADDED
        "ludwig maximilians universitat", "lmu munchen", "lmu",
        "freie universitat berlin", "humboldt universitat", 
        "karlsruher institut fur technologie", "kit",
        "tu berlin", "tu munchen", "tu dresden", "tu braunschweig",
        "universitat heidelberg", "universitat freiburg", "universitat gottingen",
        "hochschule munchen", "hochschule darmstadt", "fachhochschule"
    },
    "de": {
        "rwth aachen", "rwth", "tum", "lmu", "kit", 
        "tu berlin", "tu munchen", "tu dresden", "tu darmstadt",  # ADDED
        "hochschule munchen", "hochschule darmstadt", "fh aachen"
    },
    "uk": {
        "university of cambridge", "cambridge university", "university of oxford", "oxford university",
        "imperial college london", "imperial college", "university college london", "ucl", 
        "kings college london", "kcl",
        "university of edinburgh", "university of manchester", "university of warwick",
        "lse", "london school of economics"
    },
    "united kingdom": {
        "cambridge", "oxford", "imperial", "ucl", "edinburgh", "manchester", "warwick"
    },
    "france": {
        "sorbonne", "universite de la sorbonne",
        "ecole polytechnique", "ecole normale superieure", 
        "universite paris", "universite de lyon", "universite de toulouse"
    },
    "italy": {
        "sapienza", "universita di roma la sapienza",
        "politecnico di milano", "politecnico di torino", 
        "universita di bologna", "universita di padova", "universita di roma"
    },
    "spain": {
        "universidad autonoma de madrid", "universidad complutense", 
        "universitat de barcelona", "universidad politecnica de madrid", 
        "universidad politecnica de valencia"
    },
    "netherlands": {
        "universiteit van amsterdam", "vrije universiteit", "vrije universiteit amsterdam",
        "leiden university", "universiteit leiden", "utrecht university", "universiteit utrecht",
        "tu delft", "technische universiteit delft"
    },
    "switzerland": {
        "eth zurich", "eidgenossische technische hochschule",
        "epfl", "ecole polytechnique federale",
        "universite de geneve", "universitat zurich"
    },
    "austria": {
        "technische universitat wien", "universitat wien", "tu wien", "tu graz"
    },
    "turkey": {  # ADDED
        "bursa technical university", "bursa teknik universitesi",
        "middle east technical university", "metu", "odtu",
        "istanbul technical university", "itu",
        "bogazici university", "bogazici universitesi",
        "hacettepe university", "ankara university",
        "bilkent university", "koc university"
    },
    "tr": {  # ADDED
        "bursa technical", "bursa teknik",
        "metu", "odtu", "itu", "bogazici", "hacettepe", "bilkent", "koc"
    },
}


def normalize_text(text: str) -> str:
    """Normalize text to ASCII lowercase for comparison."""
    if not text:
        return ""
    text = str(text)
    # Handle German umlauts specifically before ASCII conversion
    text = text.replace('ä', 'a').replace('ö', 'o').replace('ü', 'u')
    text = text.replace('Ä', 'A').replace('Ö', 'O').replace('Ü', 'U')
    text = text.replace('ß', 'ss')
    
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def extract_domain(email: str) -> str:
    """Extract domain from email address."""
    if not email:
        return ""
    match = re.search(r'@([A-Za-z0-9.-]+)$', str(email).strip())
    return match.group(1).lower() if match else ""


def is_free_mailer(domain: str) -> bool:
    """Check if domain is a free email provider."""
    return domain.lower() in FREE_MAILERS


def has_commercial_indicators(text: str) -> bool:
    """Check if text contains commercial indicators.
    IMPORTANT: Does NOT include 'technical' alone since 'Technical University' is valid.
    """
    if not text:
        return False
    
    normalized = normalize_text(text)
    
    for indicator in COMMERCIAL_INDICATORS:
        # Check as whole phrase
        if indicator in normalized:
            return True
    
    return False


def is_commercial_company(text: str) -> bool:
    """Check if text matches known commercial company."""
    if not text:
        return False
    
    normalized = normalize_text(text)
    
    # Check against known commercial companies
    for company in COMMERCIAL_COMPANIES:
        company_norm = normalize_text(company)
        if company_norm in normalized:
            return True
    
    return False


def has_strict_academic_tld(domain: str) -> bool:
    """Check if domain has strict academic TLD."""
    if not domain:
        return False
    domain_lower = domain.lower()
    return any(domain_lower.endswith(tld) for tld in ACADEMIC_TLDS)


def domain_contains_university_keyword(domain: str) -> bool:
    """Check if domain explicitly contains university-related keywords."""
    if not domain:
        return False
    
    normalized = normalize_text(domain)
    
    # Must contain these exact tokens
    for token in STRICT_ACADEMIC_DOMAIN_TOKENS:
        if token in normalized:
            return True
    
    return False


def matches_known_university(company: str, country: str) -> bool:
    """Check if company matches a known university for the country."""
    if not company or not country:
        return False
    
    company_norm = normalize_text(company)
    country_norm = normalize_text(country)
    
    # First check: is it a commercial company? (but allow it if it's in known universities)
    # This is important for edge cases
    
    # Get known universities for this country
    universities = set()
    for key, vals in KNOWN_UNIVERSITIES.items():
        if key == country_norm or key in country_norm or country_norm in key:
            universities |= vals
    
    if not universities:
        return False
    
    # Check for match
    for uni in universities:
        uni_norm = normalize_text(uni)
        
        # Exact match
        if company_norm == uni_norm:
            return True
        
        # University name is contained in company name
        if uni_norm in company_norm:
            # Make sure it's not a substring of a larger word
            # e.g., "bursa" shouldn't match "bursa something"
            # but "bursa technical" should match
            words = company_norm.split()
            uni_words = uni_norm.split()
            
            # Check if all university words appear in order
            if all(word in words for word in uni_words):
                return True
    
    return False


def contains_core_academic_word(text: str) -> bool:
    """Check if text contains a core academic word (university, college, etc.)
    Must appear as a whole word, not just a substring.
    """
    if not text:
        return False
    
    normalized = normalize_text(text)
    words = normalized.split()
    
    for word in words:
        if word in CORE_ACADEMIC_WORDS:
            return True
    
    return False


class UniversityDetector:
    """Detect if a lead is from a university or educational institution."""
    
    def is_university(self, company: str, country: str, email: str) -> Dict[str, str]:
        """
        CONSERVATIVE but ACCURATE check if the lead is from a university.
        
        Priority order:
        1. Known university list (highest confidence)
        2. Academic domain TLD
        3. Domain contains university keyword
        4. Company name contains core academic word (university, college, etc.)
        
        Returns:
            dict with keys:
                - is_university: bool
                - reason: str (explanation)
                - confidence: str (high/medium/low)
        """
        domain = extract_domain(email)
        
        # STEP 1: Check against known universities FIRST (most reliable)
        if matches_known_university(company, country):
            return {
                "is_university": True,
                "reason": f"Known university in {country}: {company}",
                "confidence": "high"
            }
        
        # STEP 2: Check if it's a commercial company (but not if it matched known uni above)
        if is_commercial_company(company):
            return {
                "is_university": False,
                "reason": f"Known commercial company: {company}",
                "confidence": "high"
            }
        
        # STEP 3: Check domain (only if not free mailer)
        if not is_free_mailer(domain):
            # HIGH CONFIDENCE: Strict academic TLD (.edu, .ac.uk, etc.)
            if has_strict_academic_tld(domain):
                return {
                    "is_university": True,
                    "reason": f"Academic domain TLD: {domain}",
                    "confidence": "high"
                }
            
            # HIGH CONFIDENCE: Domain contains "university" or similar
            if domain_contains_university_keyword(domain):
                return {
                    "is_university": True,
                    "reason": f"Domain contains university keyword: {domain}",
                    "confidence": "high"
                }
        
        # STEP 4: Check company name for core academic words
        # Must contain university, college, hochschule, etc.
        if contains_core_academic_word(company):
            # Additional check: make sure it doesn't have strong commercial indicators
            if has_commercial_indicators(company):
                return {
                    "is_university": False,
                    "reason": f"Contains academic word but has commercial indicators: {company}",
                    "confidence": "medium"
                }
            
            return {
                "is_university": True,
                "reason": f"Company name contains academic keyword: {company}",
                "confidence": "medium"
            }
        
        # DEFAULT: NOT A UNIVERSITY
        return {
            "is_university": False,
            "reason": "No clear university indicators",
            "confidence": "high"
        }