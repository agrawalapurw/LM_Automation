"""
Domain Validator
Validates if email domain matches company name.
"""

import re
import socket
from typing import Dict

# Free email domains (fallback if validation data not loaded)
FREE_MAILERS = {
    "gmail.com", "yahoo.com", "outlook.com", "hotmail.com", "live.com",
    "aol.com", "icloud.com", "mail.com", "proton.me", "protonmail.com",
    "gmx.com", "gmx.de", "web.de", "yandex.com", "zoho.com", "qq.com",
    "mail.ru", "163.com", "126.com", "sina.com", "sohu.com"
}

class DomainValidator:
    """Validate if email domain matches company name."""
    
    def __init__(self, validation_loader=None):
        """Initialize validator.
        
        Args:
            validation_loader: Optional ValidationDataLoader instance
        """
        self.validation_loader = validation_loader
    
    def extract_domain(self, email: str) -> str:
        """Extract domain from email address."""
        if not email:
            return ""
        match = re.search(r'@([A-Za-z0-9.-]+)$', str(email).strip())
        return match.group(1).lower() if match else ""
    
    def normalize_name(self, text: str) -> str:
        """Normalize company/domain name for comparison."""
        if not text:
            return ""
        
        text = str(text).lower()
        
        # Remove common company suffixes
        suffixes = [
            "gmbh", "ag", "se", "ltd", "limited", "inc", "corp", "corporation",
            "llc", "plc", "sa", "srl", "bv", "nv", "kg", "ohg", "gbr",
            "co", "company", "group", "holding", "holdings", "international"
        ]
        
        for suffix in suffixes:
            text = re.sub(rf'\b{suffix}\b', '', text)
        
        # Remove special characters, keep only alphanumeric
        text = re.sub(r'[^a-z0-9]+', '', text)
        
        return text.strip()
    
    def extract_main_domain(self, domain: str) -> str:
        """Extract main domain without subdomains and TLD."""
        if not domain:
            return ""
        
        parts = domain.split('.')
        
        # Skip common subdomains
        if len(parts) >= 2:
            if parts[0] in ['mail', 'email', 'webmail', 'smtp', 'pop', 'imap', 'www']:
                parts = parts[1:]
        
        # Get the domain name
        if len(parts) >= 2:
            if len(parts) >= 3 and len(parts[-2]) <= 3:
                return parts[-3]  # For .co.uk, .com.au, etc.
            return parts[-2]  # For .com, .de, .org, etc.
        
        return parts[0] if parts else ""
    
    def is_free_mailer(self, domain: str) -> bool:
        """Check if domain is a free email provider."""
        # Check validation loader first
        if self.validation_loader and self.validation_loader.is_freemail_domain(domain):
            return True
        
        # Fallback to hardcoded list
        return domain.lower() in FREE_MAILERS
    
    def calculate_similarity(self, str1: str, str2: str) -> float:
        """Calculate similarity score between two strings (0.0 to 1.0)."""
        if not str1 or not str2:
            return 0.0
        
        if str1 == str2:
            return 1.0
        
        if str1 in str2 or str2 in str1:
            return 0.8
        
        # Count matching characters
        matches = sum(1 for i in range(min(len(str1), len(str2))) if str1[i] == str2[i])
        max_len = max(len(str1), len(str2))
        
        return matches / max_len if max_len > 0 else 0.0
    
    def validate_domain(self, company: str, email: str) -> Dict[str, str]:
        """Validate if email domain matches company name.
        
        Returns:
            dict with keys:
                - status: str
                - details: str
                - confidence: str
        """
        domain = self.extract_domain(email)
        
        if not domain:
            return {
                "status": "No Email Domain",
                "details": "Email address is empty or invalid",
                "confidence": "high"
            }
        
        if not company:
            return {
                "status": "No Company Name",
                "details": "Company name is empty",
                "confidence": "high"
            }
        
        # Check if free mailer
        if self.is_free_mailer(domain):
            return {
                "status": "Free Mailer",
                "details": f"Using free email provider: {domain}",
                "confidence": "high"
            }
        
        # Check if excluded domain
        if self.validation_loader and self.validation_loader.is_excluded_domain(domain):
            return {
                "status": "Excluded Domain",
                "details": f"Domain is in excluded list: {domain}",
                "confidence": "high"
            }
        
        # Normalize and compare
        company_normalized = self.normalize_name(company)
        domain_main = self.extract_main_domain(domain)
        domain_normalized = self.normalize_name(domain_main)
        
        similarity = self.calculate_similarity(company_normalized, domain_normalized)
        
        # High similarity
        if similarity >= 0.8:
            return {
                "status": "Valid Company Domain",
                "details": f"Domain matches company: {company} → {domain}",
                "confidence": "high"
            }
        
        # Medium similarity
        elif similarity >= 0.5:
            return {
                "status": "Possible Domain Match",
                "details": f"Partial match: {company} ≈ {domain}",
                "confidence": "medium"
            }
        
        # Low similarity
        else:
            return {
                "status": "Domain Mismatch",
                "details": f"Company and domain don't match: {company} ≠ {domain}",
                "confidence": "high"
            }