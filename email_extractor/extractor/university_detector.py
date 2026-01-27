"""
University Detector
Detects if a lead is from a university or educational institution.
"""

import re
import unicodedata
from typing import Dict

# Core academic keywords
CORE_ACADEMIC_WORDS = {
    "university", "universitat", "universitaet", "universite", "universita", 
    "universidad", "universidade", "hochschule", "fachhochschule",
    "college", "polytechnic", "polytechnique", "politecnico",
}

# Commercial indicators
COMMERCIAL_INDICATORS = {
    "gmbh", "ag", "ltd", "limited", "inc", "corp", "corporation", "llc",
    "consulting", "consultancy", "solutions", "services", "systems", 
    "technologies", "tech", "group", "holding", "holdings", "partner", 
    "partners", "international", "global", "worldwide",
}

def normalize_text(text: str) -> str:
    """Normalize text to ASCII lowercase."""
    if not text:
        return ""
    
    text = str(text)
    # Handle German umlauts
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

def has_commercial_indicators(text: str) -> bool:
    """Check if text contains commercial indicators."""
    if not text:
        return False
    
    normalized = normalize_text(text)
    return any(indicator in normalized for indicator in COMMERCIAL_INDICATORS)

def contains_core_academic_word(text: str) -> bool:
    """Check if text contains a core academic word."""
    if not text:
        return False
    
    normalized = normalize_text(text)
    words = normalized.split()
    
    return any(word in CORE_ACADEMIC_WORDS for word in words)

class UniversityDetector:
    """Detect if a lead is from a university or educational institution."""
    
    def __init__(self, validation_loader=None):
        """Initialize detector.
        
        Args:
            validation_loader: Optional ValidationDataLoader instance
        """
        self.validation_loader = validation_loader
    
    def is_university(self, company: str, country: str, email: str) -> Dict[str, str]:
        """Check if the lead is from a university.
        
        Returns:
            dict with keys:
                - is_university: bool
                - reason: str
                - confidence: str
        """
        domain = extract_domain(email)
        
        # Check validation data first
        if self.validation_loader:
            if self.validation_loader.is_academic_domain(domain):
                return {
                    "is_university": True,
                    "reason": f"Academic domain in database: {domain}",
                    "confidence": "high"
                }
        
        # Check if commercial company (from validation data)
        if self.validation_loader and self.validation_loader.is_direct_account(company):
            return {
                "is_university": False,
                "reason": f"Known commercial company: {company}",
                "confidence": "high"
            }
        
        # Check for commercial indicators
        if has_commercial_indicators(company):
            return {
                "is_university": False,
                "reason": f"Contains commercial indicators: {company}",
                "confidence": "medium"
            }
        
        # Check company name for academic words
        if contains_core_academic_word(company):
            return {
                "is_university": True,
                "reason": f"Company name contains academic keyword: {company}",
                "confidence": "medium"
            }
        
        # Default: not a university
        return {
            "is_university": False,
            "reason": "No clear university indicators",
            "confidence": "high"
        }