import os
import sys
import pandas as pd
import win32com.client
from typing import Dict, Optional


class SmartEmailMover:
    """Smart email mover that finds folders automatically."""
    
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.folder_cache = {}  # Cache found folders
        self.stats = {
            "moved": 0,
            "failed": 0,
            "skipped": 0,
            "folder_not_found": 0
        }
    
    def find_folder_recursive(self, root_folder, target_name: str, max_depth: int = 10, current_depth: int = 0):
        """Recursively search for a folder by name (case-insensitive).
        
        Args:
            root_folder: Starting folder to search from
            target_name: Name of folder to find
            max_depth: Maximum depth to search
            current_depth: Current recursion depth
            
        Returns:
            Found folder object or None
        """
        if current_depth > max_depth:
            return None
        
        try:
            for folder in root_folder.Folders:
                folder_name = folder.Name.strip()
                
                # Case-insensitive match
                if folder_name.lower() == target_name.lower():
                    return folder
                
                # Partial match (for flexibility)
                if target_name.lower() in folder_name.lower():
                    return folder
                
                # Search in subfolders
                result = self.find_folder_recursive(folder, target_name, max_depth, current_depth + 1)
                if result:
                    return result
        except Exception as e:
            # Silently continue if folder access fails
            pass
        
        return None
    
    def find_folder_in_all_stores(self, target_name: str):
        """Search for a folder across all Outlook stores.
        
        Args:
            target_name: Name of folder to find
            
        Returns:
            Found folder object or None
        """
        # Check cache first
        cache_key = target_name.lower()
        if cache_key in self.folder_cache:
            return self.folder_cache[cache_key]
        
        print(f"  Searching for folder: {target_name}...", end=" ")
        
        # Search in all stores
        try:
            for i in range(self.namespace.Stores.Count):
                store = self.namespace.Stores.Item(i + 1)
                root = store.GetRootFolder()
                
                # Search recursively
                found = self.find_folder_recursive(root, target_name)
                if found:
                    # Cache the result
                    self.folder_cache[cache_key] = found
                    print(f"✓ Found at: {found.FolderPath}")
                    return found
        except:
            pass
        
        # Also try through Folders collection
        try:
            for i in range(self.namespace.Folders.Count):
                root = self.namespace.Folders.Item(i + 1)
                
                found = self.find_folder_recursive(root, target_name)
                if found:
                    self.folder_cache[cache_key] = found
                    print(f"✓ Found at: {found.FolderPath}")
                    return found
        except:
            pass
        
        print("✗ Not found")
        return None
    
    def find_email_by_subject_and_date(self, folder, subject: str, received_time: str):
        """Find an email in a folder by subject and received time.
        
        Args:
            folder: Outlook folder to search in
            subject: Email subject
            received_time: Received time string (format: "YYYY-MM-DD HH:MM:SS")
            
        Returns:
            Email item or None
        """
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            
            # Parse received time - handle multiple formats
            from datetime import datetime
            target_time = None
            
            for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"]:
                try:
                    target_time = datetime.strptime(received_time.strip(), fmt)
                    break
                except:
                    continue
            
            if not target_time:
                print(f"    ⊘ Could not parse time: {received_time}")
                return None
            
            # Normalize subject for comparison
            subject_normalized = subject.strip().lower()
            
            # Search for matching email (increased limit and tolerance)
            for i in range(1, min(items.Count + 1, 1000)):  # Search up to 1000 recent emails
                try:
                    item = items.Item(i)
                    
                    if getattr(item, "Class", None) != 43:  # Not a mail item
                        continue
                    
                    item_subject = (getattr(item, "Subject", "") or "").strip().lower()
                    item_received = getattr(item, "ReceivedTime", None)
                    
                    # Match subject (exact match)
                    if item_subject != subject_normalized:
                        continue
                    
                    # Match time (within 5 minute tolerance for clock differences)
                    if item_received:
                        try:
                            time_diff = abs((item_received - target_time).total_seconds())
                            if time_diff <= 300:  # Within 5 minutes
                                return item
                        except:
                            # If datetime comparison fails, try date-only comparison
                            if (item_received.year == target_time.year and
                                item_received.month == target_time.month and
                                item_received.day == target_time.day):
                                return item
                except:
                    continue
            
            # If not found, try just by subject (last resort)
            print(f"    ⊘ Trying subject-only match...")
            for i in range(1, min(items.Count + 1, 500)):
                try:
                    item = items.Item(i)
                    if getattr(item, "Class", None) != 43:
                        continue
                    
                    item_subject = (getattr(item, "Subject", "") or "").strip().lower()
                    if item_subject == subject_normalized:
                        print(f"    ✓ Found by subject match")
                        return item
                except:
                    continue
            
        except Exception as e:
            print(f"    ✗ Error searching for email: {e}")
        
        return None
    
    def find_email_in_all_folders(self, subject: str, received_time: str):
        """Search for email across all folders if not found in source.
        
        Args:
            subject: Email subject
            received_time: Received time
            
        Returns:
            Tuple of (email_item, folder) or (None, None)
        """
        print(f"    Searching in all folders...")
        
        try:
            # Search all stores
            for i in range(self.namespace.Stores.Count):
                store = self.namespace.Stores.Item(i + 1)
                root = store.GetRootFolder()
                
                result = self._search_folder_tree(root, subject, received_time)
                if result:
                    return result
        except:
            pass
        
        try:
            # Also try Folders collection
            for i in range(self.namespace.Folders.Count):
                root = self.namespace.Folders.Item(i + 1)
                
                result = self._search_folder_tree(root, subject, received_time)
                if result:
                    return result
        except:
            pass
        
        return None, None
    
    def _search_folder_tree(self, folder, subject: str, received_time: str, max_depth: int = 5, current_depth: int = 0):
        """Recursively search folder tree for email."""
        if current_depth > max_depth:
            return None, None
        
        # Search in current folder
        email = self.find_email_by_subject_and_date(folder, subject, received_time)
        if email:
            print(f"    ✓ Found in: {folder.FolderPath}")
            return email, folder
        
        # Search in subfolders
        try:
            for subfolder in folder.Folders:
                result = self._search_folder_tree(subfolder, subject, received_time, max_depth, current_depth + 1)
                if result[0]:
                    return result
        except:
            pass
        
        return None, None
    
    def move_email(self, email_item, target_folder):
        """Move an email to target folder.
        
        Args:
            email_item: Outlook email item
            target_folder: Destination folder
            
        Returns:
            True if successful, False otherwise
        """
        try:
            email_item.Move(target_folder)
            return True
        except Exception as e:
            print(f"  ✗ Error moving email: {e}")
            return False
    
    def process_excel_file(self, excel_path: str, source_folder):
        """Process Excel file and move emails based on 'Move to Folder' column.
        
        Args:
            excel_path: Path to Excel file
            source_folder: Outlook folder where emails currently are
        """
        print(f"\n{'='*60}")
        print("EMAIL MOVING AUTOMATION")
        print(f"{'='*60}\n")
        
        # Process both sheets
        for sheet_name in ["Validation", "Review"]:
            try:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
            except:
                print(f"⊘ Sheet '{sheet_name}' not found, skipping...")
                continue
            
            print(f"\nProcessing {sheet_name} Sheet ({len(df)} rows)")
            print("-" * 60)
            
            self._process_sheet(df, sheet_name, source_folder, excel_path)
        
        # Print final summary
        self._print_summary()
    
    def _process_sheet(self, df: pd.DataFrame, sheet_name: str, source_folder, excel_path: str):
        """Process a single sheet."""
        status_updates = {}
        
        for index, row in df.iterrows():
            row_num = index + 2
            
            # Get move to folder value
            move_to = str(row.get("Move to Folder", "")).strip()
            
            if not move_to or move_to == "nan":
                status_updates[index] = "Not processed - No folder selected"
                self.stats["skipped"] += 1
                continue
            
            # Get email identifiers
            subject = str(row.get("Subject", "")).strip()
            received_time = str(row.get("ReceivedTime", "")).strip()
            
            if not subject or not received_time:
                print(f"Row {row_num}: ⊘ Skipped - Missing subject or time")
                status_updates[index] = "Skipped - Missing identifiers"
                self.stats["skipped"] += 1
                continue
            
            print(f"Row {row_num}: {subject[:40]}... → {move_to}")
            
            # Find target folder
            target_folder = self.find_folder_in_all_stores(move_to)
            
            if not target_folder:
                print(f"  ✗ Target folder not found: {move_to}")
                status_updates[index] = f"Failed - Folder '{move_to}' not found"
                self.stats["folder_not_found"] += 1
                continue
            
            # Find the email in source folder first
            email_item = self.find_email_by_subject_and_date(source_folder, subject, received_time)
            
            # If not found in source, search everywhere
            if not email_item:
                print(f"  ⊘ Not in source folder, searching all folders...")
                email_item, found_folder = self.find_email_in_all_folders(subject, received_time)
                
                if email_item:
                    print(f"    ✓ Found in: {found_folder.Name}")
                else:
                    print(f"  ✗ Email not found anywhere")
                    status_updates[index] = "Failed - Email not found"
                    self.stats["failed"] += 1
                    continue
            
            # Move the email
            if self.move_email(email_item, target_folder):
                print(f"  ✓ Moved successfully")
                status_updates[index] = f"✓ Moved to {move_to}"
                self.stats["moved"] += 1
            else:
                status_updates[index] = f"Failed - Could not move to {move_to}"
                self.stats["failed"] += 1
        
        # Update Excel with move status
        self._update_excel_status(excel_path, sheet_name, status_updates)
    
    def _update_excel_status(self, excel_path: str, sheet_name: str, status_updates: Dict):
        """Update Excel with email move status."""
        if not status_updates:
            return
        
        print(f"\n  Updating Excel with move status...")
        
        try:
            from openpyxl import load_workbook
            
            wb = load_workbook(excel_path)
            ws = wb[sheet_name]
            
            # Find or create "Email Move Status" column
            header_row = 1
            status_col = None
            
            for col in range(1, ws.max_column + 1):
                cell_value = ws.cell(row=header_row, column=col).value
                if cell_value == "Email Move Status":
                    status_col = col
                    break
            
            # If column doesn't exist, create it
            if status_col is None:
                status_col = ws.max_column + 1
                ws.cell(row=header_row, column=status_col, value="Email Move Status")
            
            # Update status for each row
            for excel_row, status in status_updates.items():
                excel_write_row = excel_row + 2
                ws.cell(row=excel_write_row, column=status_col, value=status)
            
            wb.save(excel_path)
            print(f"  ✓ Updated {len(status_updates)} rows in Excel")
            
        except Exception as e:
            print(f"  ✗ Error updating Excel: {e}")
    
    def _print_summary(self):
        """Print final statistics."""
        print(f"\n{'='*60}")
        print("EMAIL MOVING SUMMARY")
        print(f"{'='*60}")
        print(f"Successfully moved:    {self.stats['moved']}")
        print(f"Failed to move:        {self.stats['failed']}")
        print(f"Folder not found:      {self.stats['folder_not_found']}")
        print(f"Skipped (no folder):   {self.stats['skipped']}")
        print(f"{'='*60}\n")


def select_source_folder(namespace):
    """Interactive folder selection."""
    stores = []
    try:
        for i in range(namespace.Stores.Count):
            stores.append(namespace.Stores.Item(i + 1).DisplayName)
    except:
        for i in range(namespace.Folders.Count):
            stores.append(namespace.Folders.Item(i + 1).Name)
    
    if not stores:
        raise RuntimeError("No Outlook stores found")
    
    print("Available stores:")
    for i, store in enumerate(stores, 1):
        print(f"  [{i}] {store}")
    
    while True:
        try:
            choice = int(input("Select store: ").strip())
            if 1 <= choice <= len(stores):
                store_name = stores[choice - 1]
                break
        except:
            pass
        print("Invalid selection.")
    
    # Get root folder
    try:
        for i in range(namespace.Stores.Count):
            store = namespace.Stores.Item(i + 1)
            if store.DisplayName == store_name:
                root = store.GetRootFolder()
                break
    except:
        for i in range(namespace.Folders.Count):
            root = namespace.Folders.Item(i + 1)
            if root.Name == store_name:
                break
    
    # Navigate to folder
    path = [root]
    while True:
        current = path[-1]
        print(f"\nCurrent: {current.Name}")
        
        try:
            subfolders = list(current.Folders)
        except:
            subfolders = []
        
        if subfolders:
            print("Subfolders:")
            for i, folder in enumerate(subfolders, 1):
                print(f"  [{i}] {folder.Name}")
        
        cmd = input("Enter number, 's' to select, 'u' for up, 'q' to quit: ").strip().lower()
        
        if cmd == 's':
            return current
        elif cmd == 'u':
            if len(path) > 1:
                path.pop()
        elif cmd == 'q':
            sys.exit(0)
        else:
            try:
                idx = int(cmd) - 1
                if 0 <= idx < len(subfolders):
                    path.append(subfolders[idx])
            except:
                print("Invalid input.")


def main():
    print("=" * 60)
    print("EMAIL MOVING TOOL")
    print("=" * 60)
    print("\nThis tool will:")
    print("1. Read 'Move to Folder' column from Excel")
    print("2. Find target folders automatically (any depth)")
    print("3. Move emails to respective folders")
    print("4. Update Excel with move status")
    print("=" * 60)
    
    # Get Excel file
    output_dir = os.path.join(os.path.dirname(__file__), "output")
    
    if not os.path.exists(output_dir):
        print(f"\n✗ Output directory not found: {output_dir}")
        return
    
    excel_files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx') and not f.startswith('~$')]
    
    if not excel_files:
        print(f"\n✗ No Excel files found in {output_dir}")
        return
    
    print("\nAvailable Excel files:")
    for i, file in enumerate(excel_files, 1):
        print(f"  [{i}] {file}")
    
    while True:
        try:
            choice = int(input("\nSelect file number: ").strip())
            if 1 <= choice <= len(excel_files):
                excel_file = excel_files[choice - 1]
                break
        except:
            pass
        print("Invalid selection.")
    
    excel_path = os.path.join(output_dir, excel_file)
    print(f"\n✓ Selected: {excel_file}")
    
    # Select source folder (where emails currently are)
    print("\nSelect the folder where Pre-MQL emails are currently located:")
    print("(This is the folder you selected during extraction)")
    
    mover = SmartEmailMover()
    source_folder = select_source_folder(mover.namespace)
    
    print(f"\n✓ Source folder: {source_folder.FolderPath}")
    
    # Confirmation
    print("\n" + "=" * 60)
    print("READY TO MOVE EMAILS")
    print("=" * 60)
    print(f"Excel file: {excel_file}")
    print(f"Source folder: {source_folder.Name}")
    print(f"Target folders will be found automatically")
    print("=" * 60)
    
    confirm = input("\nProceed? [Y/n]: ").strip().lower()
    if confirm and confirm not in ("y", "yes"):
        print("Cancelled.")
        return
    
    # Process
    print("\nStarting email moving process...\n")
    mover.process_excel_file(excel_path, source_folder)
    
    print("\n✓ Email moving completed!")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user.")
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()