import re
from typing import Dict
import socket


# Free email domains
FREE_MAILERS = {
    "gmail.com", "yahoo.com", "outlook.com", "hotmail.com", "live.com",
    "aol.com", "icloud.com", "mail.com", "proton.me", "protonmail.com",
    "gmx.com", "gmx.de", "web.de", "yandex.com", "zoho.com", "qq.com",
    "mail.ru", "163.com", "126.com", "sina.com", "sohu.com"
}


class DomainValidator:
    """Validate if email domain matches company name."""
    
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
        
        # Convert to lowercase
        text = str(text).lower()
        
        # Remove common company suffixes
        suffixes = [
            "gmbh", "ag", "se", "ltd", "limited", "inc", "corp", "corporation",
            "llc", "plc", "sa", "srl", "bv", "nv", "kg", "ohg", "gbr",
            "co", "company", "group", "holding", "holdings", "international"
        ]
        
        for suffix in suffixes:
            # Remove as whole word
            text = re.sub(rf'\b{suffix}\b', '', text)
        
        # Remove special characters, keep only alphanumeric
        text = re.sub(r'[^a-z0-9]+', '', text)
        
        return text.strip()
    
    def extract_main_domain(self, domain: str) -> str:
        """Extract main domain without subdomains and TLD.
        
        Example: mail.company.co.uk -> company
        """
        if not domain:
            return ""
        
        # Remove common subdomains
        parts = domain.split('.')
        
        # If we have subdomain.domain.tld or domain.tld
        if len(parts) >= 2:
            # Skip common subdomains
            if parts[0] in ['mail', 'email', 'webmail', 'smtp', 'pop', 'imap', 'www']:
                parts = parts[1:]
        
        # Get the domain name (second to last part usually)
        # For domain.tld -> take 'domain'
        # For domain.co.uk -> take 'domain'
        if len(parts) >= 2:
            # Check if second-to-last is a short TLD part (co, com, ac, etc.)
            if len(parts) >= 3 and len(parts[-2]) <= 3:
                return parts[-3]  # For .co.uk, .com.au, etc.
            return parts[-2]  # For .com, .de, .org, etc.
        
        return parts[0] if parts else ""
    
    def is_free_mailer(self, domain: str) -> bool:
        """Check if domain is a free email provider."""
        return domain.lower() in FREE_MAILERS
    
    def check_domain_accessibility(self, domain: str) -> bool:
        """Check if domain is accessible (has DNS records).
        
        Returns True if domain appears to be valid/accessible.
        """
        if not domain:
            return False
        
        try:
            # Try to resolve domain
            socket.gethostbyname(domain)
            return True
        except socket.gaierror:
            # Domain doesn't resolve
            return False
        except Exception:
            # Other errors, assume accessible to be safe
            return True
    
    def calculate_similarity(self, str1: str, str2: str) -> float:
        """Calculate simple similarity score between two strings.
        
        Returns value between 0.0 and 1.0
        """
        if not str1 or not str2:
            return 0.0
        
        # Exact match
        if str1 == str2:
            return 1.0
        
        # One contains the other
        if str1 in str2 or str2 in str1:
            return 0.8
        
        # Check for common substring
        min_len = min(len(str1), len(str2))
        max_len = max(len(str1), len(str2))
        
        # Count matching characters in order
        matches = 0
        for i in range(min(len(str1), len(str2))):
            if str1[i] == str2[i]:
                matches += 1
        
        return matches / max_len if max_len > 0 else 0.0
    
    def validate_domain(self, company: str, email: str) -> Dict[str, str]:
        """Validate if email domain matches company name.
        
        Returns:
            dict with keys:
                - status: str (Valid Company Domain / Mismatch / Free Mailer / etc.)
                - details: str (explanation)
                - confidence: str (high/medium/low)
        """
        domain = self.extract_domain(email)
        
        # Check if empty
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
        
        # Normalize company name and domain
        company_normalized = self.normalize_name(company)
        domain_main = self.extract_main_domain(domain)
        domain_normalized = self.normalize_name(domain_main)
        
        # Check similarity
        similarity = self.calculate_similarity(company_normalized, domain_normalized)
        
        # High similarity - likely a match
        if similarity >= 0.8:
            # Check if domain is accessible
            is_accessible = self.check_domain_accessibility(domain)
            
            if is_accessible:
                return {
                    "status": "Valid Company Domain",
                    "details": f"Domain matches company: {company} → {domain}",
                    "confidence": "high"
                }
            else:
                return {
                    "status": "Matching Domain - Not Accessible",
                    "details": f"Domain matches but not accessible: {domain}",
                    "confidence": "medium"
                }
        
        # Medium similarity - possible match
        elif similarity >= 0.5:
            is_accessible = self.check_domain_accessibility(domain)
            
            return {
                "status": "Possible Domain Match",
                "details": f"Partial match: {company} ≈ {domain} (accessibility: {'yes' if is_accessible else 'no'})",
                "confidence": "medium"
            }
        
        # Low similarity - mismatch
        else:
            is_accessible = self.check_domain_accessibility(domain)
            
            return {
                "status": "Domain Mismatch",
                "details": f"Company and domain don't match: {company} ≠ {domain} (domain accessible: {'yes' if is_accessible else 'no'})",
                "confidence": "high"
            }