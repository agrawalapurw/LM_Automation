"""
Email Parser
Parses Outlook email items into structured data.
"""

import re
from typing import Dict
from urllib.parse import urlparse, parse_qs, unquote

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

EMAIL_PATTERN = re.compile(r'([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})', re.I)
URL_PATTERN = re.compile(r'https?://[^\s"\'>]+', re.I)

# Field labels to extract
FIELDS = [
    "Subject", "Sender", "ReceivedTime", "All Emails Found",
    "First Name", "Last Name", "Email Address", "Company", "Pages Viewed",
    "Submit Time", "Form Name", "URL Of Form", "Salutation", "Business Phone",
    "Country", "City", "State/Province", "Researched State or Province",
    "Job Role", "Industry", "Lead Triggering Activities", "Project yes/no",
    "Start of Production", "Project Volume", "Project Timeframe", "Rejection reason",
    "Rejection reason free text", "Lead Source - Most Recent", "Lead Source - Original",
    "Lead Source Name - Most Recent", "Lead Source Name - Original", "Lead Trigger",
    "Lead Lifecycle Count", "Account Type", "Lead Lifecycle ID", "Lead editor",
    "Subject Line", "Notification", "PreMQL review/validation link",
    "Company Matching Status",
    "Potential Distribution Partner (matching in beta testing)", "Digital activity",
    "Eloqua Profiler", "Initial Call Notes",
    "Has Contact Sales Form",
    "Status",
    "Action Taken",
    "Company Domain Validation",
    "Validation Status",
    "Validation Reason",
    "Take Action",
    "Valid Company → Reject Reason",
    "Invalid Company Reason",
    "Reject Reason",
    "Additional Scoring Information",
    "Send to",
    "Move to Folder",
    "Form Submission Status",
    "Email Move Status"
]

class EmailParser:
    """Parse Outlook email items into structured data."""
    
    def __init__(self, university_detector=None, validation_loader=None):
        """Initialize parser.
        
        Args:
            university_detector: UniversityDetector instance
            validation_loader: ValidationDataLoader instance
        """
        self.university_detector = university_detector
        self.validation_loader = validation_loader
    
    def parse_email(self, email_item) -> Dict[str, str]:
        """Parse Outlook email item into structured data."""
        subject = getattr(email_item, "Subject", "") or ""
        sender = getattr(email_item, "SenderEmailAddress", "") or ""
        received = getattr(email_item, "ReceivedTime", None)
        body = getattr(email_item, "Body", "") or ""
        html = getattr(email_item, "HTMLBody", "") or ""
        
        # Parse fields
        data = {}
        
        if HAS_BS4 and html:
            data.update(self._parse_html(html))
        
        data.update(self._parse_text(body))
        
        # Build row
        row = {
            "Subject": subject,
            "Sender": sender,
            "ReceivedTime": self._format_datetime(received),
            "All Emails Found": "; ".join(sorted(set(self._find_emails(body)))),
        }
        
        for field in FIELDS:
            if field not in row:
                row[field] = data.get(field, "")
        
        # Check for contact sales form
        row["Has Contact Sales Form"] = self._check_contact_sales_form(
            row.get("Lead Triggering Activities", "")
        )
        
        # Get key fields for validation
        company = row.get("Company", "")
        country = row.get("Country", "")
        email = row.get("Email Address", "")
        
        # Run validation checks
        if self.validation_loader:
            validation_result = self.validation_loader.validate_lead(company, country, email)
            
            if not validation_result["is_valid"]:
                row["Status"] = validation_result["validation_type"]
                row["Action Taken"] = validation_result["reason"]
                row["Validation Status"] = "Invalid"
                row["Validation Reason"] = validation_result["reason"]
            else:
                row["Validation Status"] = "Valid"
                row["Validation Reason"] = validation_result.get("reason", "")
        
        # Check for university (only if not already marked invalid)
        if row.get("Status", "") == "" and self.university_detector:
            university_result = self.university_detector.is_university(company, country, email)
            
            if university_result["is_university"]:
                row["Status"] = "University Contact"
                row["Action Taken"] = f"Identified as university - {university_result['reason']}"
        
        # Default status if still empty
        if not row.get("Status"):
            row["Status"] = "Not Started"
        
        if not row.get("Action Taken"):
            row["Action Taken"] = ""
        
        return row
    
    def _check_contact_sales_form(self, triggering_activities: str) -> str:
        """Check if Lead Triggering Activities contains contact_sales_forms."""
        if not triggering_activities:
            return "No"
        
        lower_activities = triggering_activities.lower()
        patterns = [
            "contact_sales_forms", "contact sales forms",
            "contact_sales_form", "contact sales form",
            "contactsalesforms", "contactsalesform"
        ]
        
        return "Yes" if any(p in lower_activities for p in patterns) else "No"
    
    def _parse_html(self, html: str) -> Dict[str, str]:
        """Extract fields from HTML body."""
        soup = BeautifulSoup(html, "html.parser")
        data = {}
        
        # Extract table-based fields
        for table in soup.find_all("table"):
            for row in table.find_all("tr"):
                cells = row.find_all(["td", "th"])
                if len(cells) >= 2:
                    label = self._normalize_label(cells[0].get_text(strip=True))
                    value = cells[1].get_text(separator="\n", strip=True)
                    link = cells[1].find("a")
                    
                    if link and link.get("href"):
                        href = self._unwrap_url(link.get("href"))
                        if "link" in label.lower() or "url" in label.lower():
                            data[label] = href
                        else:
                            data[label] = value
                    elif value:
                        data[label] = self._clean_value(value)
        
        # Extract links
        for anchor in soup.find_all("a"):
            text = anchor.get_text(strip=True).lower()
            href = anchor.get("href", "").strip()
            
            if not href:
                continue
            
            if any(kw in text for kw in ["click here", "qualify", "qualification"]):
                data["PreMQL review/validation link"] = self._unwrap_url(href)
            elif "profiler" in text:
                data["Eloqua Profiler"] = self._unwrap_url(href)
        
        return data
    
    def _parse_text(self, body: str) -> Dict[str, str]:
        """Extract fields from plain text body."""
        data = {}
        lines = [l.strip() for l in body.splitlines()]
        
        current_field = None
        buffer = []
        
        for line in lines:
            if line.startswith("Copyright"):
                break
            
            normalized = self._normalize_label(line)
            if normalized in FIELDS:
                if current_field and buffer:
                    data[current_field] = self._clean_value("\n".join(buffer).strip())
                current_field = normalized
                buffer = []
            elif current_field:
                buffer.append(line)
        
        if current_field and buffer:
            data[current_field] = self._clean_value("\n".join(buffer).strip())
        
        # Split PreMQL link and Company Matching Status
        if "PreMQL review/validation link" in data:
            value = data["PreMQL review/validation link"]
            if "Company Matching Status" in value:
                parts = value.split("\n")
                data["PreMQL review/validation link"] = parts[0].strip()
                if len(parts) > 1:
                    status_parts = [p.strip() for p in parts[1:] 
                                   if p.strip() and "Company Matching Status" not in p]
                    if status_parts:
                        data["Company Matching Status"] = status_parts[0]
        
        # Find links in text
        if "PreMQL review/validation link" not in data:
            for i, line in enumerate(lines):
                if re.search(r"click\s+(here|on\s+this\s+link)", line, re.I):
                    window = "\n".join(lines[i:i+5])
                    match = URL_PATTERN.search(window)
                    if match:
                        data["PreMQL review/validation link"] = self._unwrap_url(match.group(0))
                        break
        
        return data
    
    def _clean_value(self, value: str) -> str:
        """Clean extracted value."""
        if not value:
            return ""
        
        if "Copyright" in value:
            value = value.split("Copyright")[0].strip()
        
        value = re.sub(r'Copyright.*?All rights reserved\.?', '', value, flags=re.DOTALL | re.IGNORECASE)
        value = re.sub(r'Oracle and/or its affiliates\.?', '', value, flags=re.IGNORECASE)
        value = re.sub(r'<https?://[^>]+>', '', value)
        value = re.sub(r'https?://[^\s]*tinydot\.gif[^\s]*', '', value)
        value = re.sub(r'https?://img\d+\.en25\.com[^\s]*', '', value)
        
        if value.strip() == "Company Matching Status":
            return ""
        
        lines = [l.strip() for l in value.splitlines() if l.strip()]
        return "\n".join(lines).strip()
    
    def _normalize_label(self, text: str) -> str:
        """Normalize field label."""
        text = text.replace(":", "").strip()
        
        aliases = {
            "lead qualification link": "PreMQL review/validation link",
            "qualification link": "PreMQL review/validation link",
            "click here": "PreMQL review/validation link",
        }
        
        lower = text.lower()
        if lower in aliases:
            return aliases[lower]
        
        for field in FIELDS:
            if lower == field.lower():
                return field
        
        return text
    
    def _unwrap_url(self, url: str) -> str:
        """Unwrap tracking/redirect URLs."""
        try:
            parsed = urlparse(url)
            qs = parse_qs(parsed.query)
            
            if "safelinks.protection.outlook.com" in parsed.netloc:
                if "url" in qs:
                    return unquote(qs["url"][0])
            
            for key in ["url", "u", "redirect", "target"]:
                if key in qs and qs[key][0].startswith("http"):
                    return unquote(qs[key][0])
        except:
            pass
        
        return url
    
    def _find_emails(self, text: str) -> list:
        """Extract all email addresses from text."""
        text = text.replace("mailto:", " ").replace("<", " ").replace(">", " ")
        return [m.lower() for m in EMAIL_PATTERN.findall(text)]
    
    def _format_datetime(self, dt) -> str:
        """Format datetime to ISO string."""
        if not dt:
            return ""
        try:
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except:
            return str(dt)