"""
Validation Data Loader
Loads files containing validation rules for lead processing.
"""
import os
from typing import Set, Dict, Optional
import pandas as pd


class ValidationDataLoader:
    """Load and manage validation data from CSV/XLSX files."""

    def __init__(self, data_folder: Optional[str] = None):
        """
        Args:
            data_folder: Path to folder containing validation files.
                         If None, resolves to the repo-level 'validation_data' folder.
        """
        # Resolve default to the project root's validation_data
        if data_folder is None:
            # extractor/ -> Email_extractor/ -> repo root
            extractor_dir = os.path.dirname(__file__)
            email_extractor_dir = os.path.dirname(extractor_dir)
            repo_root = os.path.dirname(email_extractor_dir)
            data_folder = os.path.join(repo_root, "validation_data")

        self.data_folder = data_folder

        self.academic_domains: Set[str] = set()
        self.academic_domain_names: Dict[str, str] = {}  # domain -> institution name (if provided)
        self.excluded_domains: Set[str] = set()
        self.direct_accounts: Set[str] = set()
        self.blacklisted_countries: Set[str] = set()
        self.freemail_domains: Set[str] = set()

        self.load_all()

    def load_all(self):
        print(f"Loading validation data from: {self.data_folder}")
        try:
            self._load_domains("academic_domains", self.academic_domains, self.academic_domain_names)
            self._load_simple("excluded_domains", self.excluded_domains)
            self._load_simple("direct_accounts", self.direct_accounts)
            self._load_simple("blacklisted_countries", self.blacklisted_countries)
            self._load_simple("freemail_domains", self.freemail_domains)

            print(f"✓ Loaded validation data:")
            print(f"  - Academic domains: {len(self.academic_domains)}")
            print(f"  - Excluded domains: {len(self.excluded_domains)}")
            print(f"  - Direct accounts: {len(self.direct_accounts)}")
            print(f"  - Blacklisted countries: {len(self.blacklisted_countries)}")
            print(f"  - Freemail domains: {len(self.freemail_domains)}")
        except Exception as e:
            print(f"⚠ Warning: Could not load all validation data: {e}")
            print("  Continuing with available data...")

    def _load_simple(self, base_name: str, target_set: Set[str]):
        """Load a CSV/XLSX with 'Option Values' column into a set."""
        df = self._read_file(base_name)
        if df is None:
            print(f"  ⊘ {base_name}: file not found")
            return
        if "Option Values" not in df.columns:
            print(f"  ⊘ {base_name}: 'Option Values' column not found")
            return
        values = (
            df["Option Values"]
            .dropna()
            .astype(str)
            .str.strip()
            .str.lower()
            .unique()
        )
        target_set.update(values)

    def _load_domains(self, base_name: str, target_set: Set[str], name_map: Dict[str, str]):
        """Load academic domains with optional institution names (Option Name)."""
        df = self._read_file(base_name)
        if df is None:
            print(f"  ⊘ {base_name}: file not found")
            return
        if "Option Values" not in df.columns:
            print(f"  ⊘ {base_name}: 'Option Values' column not found")
            return
        domains = (
            df["Option Values"]
            .dropna()
            .astype(str)
            .str.strip()
            .str.lower()
            .unique()
        )
        target_set.update(domains)
        if "Option Name" in df.columns:
            for _, row in df.dropna(subset=["Option Values"]).iterrows():
                dom = str(row["Option Values"]).strip().lower()
                name = str(row.get("Option Name", "")).strip()
                if dom and name:
                    name_map[dom] = name

    def _read_file(self, base_name: str) -> Optional[pd.DataFrame]:
        """Read <base_name>.csv or <base_name>.xlsx from data_folder."""
        csv_path = os.path.join(self.data_folder, f"{base_name}.csv")
        xlsx_path = os.path.join(self.data_folder, f"{base_name}.xlsx")
        try:
            if os.path.exists(csv_path):
                return pd.read_csv(csv_path)
            if os.path.exists(xlsx_path):
                return pd.read_excel(xlsx_path)
        except Exception as e:
            print(f"  ⊘ Error reading {base_name}: {e}")
        return None

    # Lookups
    def is_academic_domain(self, domain: str) -> bool:
        """Check if domain belongs to an academic domain list (exact or suffix match)."""
        d = (domain or "").strip().lower()
        if not d:
            return False
        if d in self.academic_domains:
            return True
        # Suffix match for subdomains: check if any academic domain is a suffix of d
        return any(d == ad or d.endswith("." + ad) for ad in self.academic_domains)

    def get_academic_name(self, domain: str) -> str:
        """Return mapped institution name for a domain if known."""
        d = (domain or "").strip().lower()
        if d in self.academic_domain_names:
            return self.academic_domain_names[d]
        # Try suffix matches
        for ad, name in self.academic_domain_names.items():
            if d.endswith("." + ad):
                return name
        return ""

    def is_excluded_domain(self, domain: str) -> bool:
        d = (domain or "").strip().lower()
        return d in self.excluded_domains or any(d.endswith("." + ed) for ed in self.excluded_domains)

    def is_direct_account(self, company: str) -> bool:
        return (company or "").strip().lower() in self.direct_accounts

    def is_blacklisted_country(self, country: str) -> bool:
        return (country or "").strip().lower() in self.blacklisted_countries

    def is_freemail_domain(self, domain: str) -> bool:
        d = (domain or "").strip().lower()
        return d in self.freemail_domains or any(d.endswith("." + fd) for fd in self.freemail_domains)

    def validate_lead(self, company: str, country: str, email: str):
        """Maintains original interface if called from other modules."""
        domain = ""
        if email and "@" in email:
            domain = email.split("@")[-1].lower().strip()
        if self.is_blacklisted_country(country):
            return {"is_valid": False, "reason": f"Blacklisted country: {country}", "validation_type": "Country"}
        if self.is_direct_account(company):
            return {"is_valid": False, "reason": f"Direct account: {company}", "validation_type": "Direct Account"}
        if self.is_excluded_domain(domain):
            return {"is_valid": False, "reason": f"Excluded domain: {domain}", "validation_type": "Excluded Domain"}
        if self.is_academic_domain(domain):
            return {"is_valid": False, "reason": f"Academic domain: {domain}", "validation_type": "Academic"}
        if self.is_freemail_domain(domain):
            return {"is_valid": True, "reason": f"Freemail provider: {domain}", "validation_type": "Freemail"}
        return {"is_valid": True, "reason": "", "validation_type": "Valid"}