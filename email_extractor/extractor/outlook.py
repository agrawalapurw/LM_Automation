"""
Outlook Client
Handles Outlook connection and email fetching.
"""

import sys
from datetime import datetime, timedelta
from typing import List, Tuple
import win32com.client

class OutlookClient:
    """Client for interacting with Outlook."""
    
    def __init__(self):
        """Initialize Outlook connection."""
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
    
    def list_stores(self) -> List[str]:
        """Get list of available Outlook stores."""
        stores = []
        try:
            for i in range(self.namespace.Stores.Count):
                stores.append(self.namespace.Stores.Item(i + 1).DisplayName)
        except:
            for i in range(self.namespace.Folders.Count):
                stores.append(self.namespace.Folders.Item(i + 1).Name)
        return stores
    
    def get_store(self, store_name: str):
        """Get store by name."""
        try:
            for i in range(self.namespace.Stores.Count):
                store = self.namespace.Stores.Item(i + 1)
                if store.DisplayName == store_name:
                    return store.GetRootFolder()
        except:
            pass
        
        for i in range(self.namespace.Folders.Count):
            folder = self.namespace.Folders.Item(i + 1)
            if folder.Name == store_name:
                return folder
        
        raise ValueError(f"Store not found: {store_name}")
    
    def select_store(self):
        """Interactive store selection."""
        stores = self.list_stores()
        if not stores:
            raise RuntimeError("No Outlook stores found.")
        
        print("\nAvailable stores:")
        for i, store in enumerate(stores, 1):
            print(f"  [{i}] {store}")
        
        while True:
            try:
                choice = int(input("Select store number: ").strip())
                if 1 <= choice <= len(stores):
                    return self.get_store(stores[choice - 1])
            except:
                pass
            print("Invalid selection. Try again.")
    
    def select_folder(self, root):
        """Interactive folder navigation."""
        path = [root]
        
        while True:
            current = path[-1]
            print(f"\nCurrent: {current.Name}")
            
            subfolders = list(current.Folders)
            if subfolders:
                print("Subfolders:")
                for i, folder in enumerate(subfolders, 1):
                    print(f"  [{i}] {folder.Name}")
            
            cmd = input("Enter number to open, 's' to select, 'u' for up, 'q' to quit: ").strip().lower()
            
            if cmd == "s":
                return current
            elif cmd == "u":
                if len(path) > 1:
                    path.pop()
            elif cmd == "q":
                sys.exit(0)
            else:
                try:
                    idx = int(cmd) - 1
                    if 0 <= idx < len(subfolders):
                        path.append(subfolders[idx])
                except:
                    print("Invalid input.")
    
    def parse_date_input(self, date_str: str) -> List[Tuple[datetime, datetime]]:
        """Parse date string into list of (start, end) tuples.
        
        Supports:
            - Single date: 2024-01-15
            - Multiple dates: 2024-01-15,2024-01-16
            - Range: 2024-01-15 to 2024-01-20
        """
        ranges = []
        
        def to_range(date):
            start = datetime(date.year, date.month, date.day, 0, 0, 0)
            end = start + timedelta(days=1)
            return (start, end)
        
        if " to " in date_str:
            parts = date_str.split(" to ")
            start = datetime.fromisoformat(parts[0].strip())
            end = datetime.fromisoformat(parts[1].strip())
            
            current = datetime(start.year, start.month, start.day)
            end_date = datetime(end.year, end.month, end.day)
            
            while current <= end_date:
                ranges.append(to_range(current))
                current += timedelta(days=1)
        elif "," in date_str:
            for part in date_str.split(","):
                date = datetime.fromisoformat(part.strip())
                ranges.append(to_range(date))
        else:
            date = datetime.fromisoformat(date_str.strip())
            ranges.append(to_range(date))
        
        return ranges
    
    def fetch_emails(self, folder, date_ranges: List[Tuple[datetime, datetime]], subject_filters: List[str]):
        """Fetch emails from folder matching criteria."""
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        
        collected = []
        seen_ids = set()
        
        for start, end in date_ranges:
            query = self._build_query(start, end)
            
            try:
                filtered = items.Restrict(query)
            except:
                items.IncludeRecurrences = False
                filtered = items.Restrict(query)
            
            for i in range(1, filtered.Count + 1):
                try:
                    item = filtered.Item(i)
                    
                    # Skip non-mail items
                    if getattr(item, "Class", None) != 43:
                        continue
                    
                    # Check subject filter
                    subject = getattr(item, "Subject", "") or ""
                    if subject_filters and not any(f.lower() in subject.lower() for f in subject_filters):
                        continue
                    
                    # Deduplicate
                    entry_id = getattr(item, "EntryID", None)
                    if entry_id and entry_id not in seen_ids:
                        collected.append(item)
                        seen_ids.add(entry_id)
                except:
                    continue
        
        return collected
    
    def _build_query(self, start: datetime, end: datetime) -> str:
        """Build Outlook filter query."""
        fmt = "%m/%d/%Y %I:%M %p"
        return f"[ReceivedTime] >= '{start.strftime(fmt)}' AND [ReceivedTime] < '{end.strftime(fmt)}'"