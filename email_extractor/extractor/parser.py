import re
from typing import Dict
from urllib.parse import urlparse, parse_qs, unquote
from extractor.university_detector import UniversityDetector

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
    "Company Domain Validation"  # NEW
]


class EmailParser:
    def __init__(self):
        self.university_detector = UniversityDetector()
    
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
        
        # Check for contact_sales_forms
        row["Has Contact Sales Form"] = self._check_contact_sales_form(row.get("Lead Triggering Activities", ""))
        
        # NEW: Check for university
        company = row.get("Company", "")
        country = row.get("Country", "")
        email = row.get("Email Address", "")
        
        university_result = self.university_detector.is_university(company, country, email)
        
        if university_result["is_university"]:
            row["Status"] = "University Contact"
            row["Action Taken"] = f"Identified as university - {university_result['reason']} (Confidence: {university_result['confidence']})"
        else:
            # Will be updated by email mover if applicable
            row["Status"] = "Not Started"
            row["Action Taken"] = ""
        
        return row
    
    def _check_contact_sales_form(self, triggering_activities: str) -> str:
        """Check if Lead Triggering Activities contains contact_sales_forms."""
        if not triggering_activities:
            return "No"
        
        # Check for various patterns
        lower_activities = triggering_activities.lower()
        
        patterns = [
            "contact_sales_forms",
            "contact sales forms",
            "contact_sales_form",
            "contact sales form",
            "contactsalesforms",
            "contactsalesform"
        ]
        
        for pattern in patterns:
            if pattern in lower_activities:
                return "Yes"
        
        return "No"
    
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
                        # Clean the value
                        cleaned_value = self._clean_value(value)
                        data[label] = cleaned_value
        
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
            # Stop at Copyright
            if line.startswith("Copyright"):
                break
            
            # Check if line is a field label
            normalized = self._normalize_label(line)
            if normalized in FIELDS:
                # Save previous field
                if current_field and buffer:
                    value = "\n".join(buffer).strip()
                    cleaned_value = self._clean_value(value)
                    data[current_field] = cleaned_value
                current_field = normalized
                buffer = []
            elif current_field:
                buffer.append(line)
        
        # Save last field
        if current_field and buffer:
            value = "\n".join(buffer).strip()
            cleaned_value = self._clean_value(value)
            data[current_field] = cleaned_value
        
        # Split PreMQL link and Company Matching Status
        if "PreMQL review/validation link" in data:
            value = data["PreMQL review/validation link"]
            # Check if it contains "Company Matching Status"
            if "Company Matching Status" in value:
                parts = value.split("\n")
                # First part is the URL
                url_part = parts[0].strip()
                data["PreMQL review/validation link"] = url_part
                # Everything after is Company Matching Status
                if len(parts) > 1:
                    status_parts = [p.strip() for p in parts[1:] if p.strip() and "Company Matching Status" not in p]
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
        """Clean extracted value from unwanted content."""
        if not value:
            return ""
        
        # Remove everything after Copyright
        if "Copyright" in value:
            value = value.split("Copyright")[0].strip()
        
        # Remove Oracle footer patterns
        value = re.sub(r'Copyright.*?All rights reserved\.?', '', value, flags=re.DOTALL | re.IGNORECASE)
        value = re.sub(r'Oracle and/or its affiliates\.?', '', value, flags=re.IGNORECASE)
        
        # Remove image URLs and tracking pixels
        value = re.sub(r'<https?://[^>]+>', '', value)
        value = re.sub(r'https?://[^\s]*tinydot\.gif[^\s]*', '', value)
        value = re.sub(r'https?://img\d+\.en25\.com[^\s]*', '', value)
        
        # Remove "Company Matching Status" label if it's standalone
        if value.strip() == "Company Matching Status":
            return ""
        
        # Clean up whitespace
        lines = [l.strip() for l in value.splitlines() if l.strip()]
        value = "\n".join(lines)
        
        return value.strip()
    
    def _normalize_label(self, text: str) -> str:
        """Normalize field label."""
        text = text.replace(":", "").strip()
        
        # Map aliases
        aliases = {
            "lead qualification link": "PreMQL review/validation link",
            "qualification link": "PreMQL review/validation link",
            "click here": "PreMQL review/validation link",
        }
        
        lower = text.lower()
        if lower in aliases:
            return aliases[lower]
        
        # Match canonical label
        for field in FIELDS:
            if lower == field.lower():
                return field
        
        return text
    
    def _unwrap_url(self, url: str) -> str:
        """Unwrap tracking/redirect URLs."""
        try:
            parsed = urlparse(url)
            qs = parse_qs(parsed.query)
            
            # Outlook SafeLinks
            if "safelinks.protection.outlook.com" in parsed.netloc:
                if "url" in qs:
                    return unquote(qs["url"][0])
            
            # Generic unwrap
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