import re
from typing import List, Dict


class EmailMover:
    """Move emails to appropriate folders based on rules."""
    
    def __init__(self, outlook_client):
        self.outlook = outlook_client
        self.move_log = []
    
    def find_folder_recursive(self, root_folder, target_name, max_depth=5, current_depth=0):
        """Recursively search for a folder by name."""
        if current_depth > max_depth:
            return None
        
        try:
            for folder in root_folder.Folders:
                folder_name = folder.Name.strip()
                # Check if this is the MQL folder
                if target_name.upper() in folder_name.upper():
                    return folder
                
                # Search in subfolders
                result = self.find_folder_recursive(folder, target_name, max_depth, current_depth + 1)
                if result:
                    return result
        except Exception as e:
            print(f"Error searching folder: {e}")
        
        return None
    
    def list_all_folders(self, root_folder, indent=0, max_depth=3, current_depth=0):
        """List all folders for debugging."""
        if current_depth > max_depth:
            return
        
        try:
            for folder in root_folder.Folders:
                print("  " * indent + f"└─ {folder.Name}")
                self.list_all_folders(folder, indent + 1, max_depth, current_depth + 1)
        except Exception as e:
            print(f"Error listing folders: {e}")
    
    def get_mql_subfolders(self, root_folder):
        """Navigate to MQL folder and get its subfolders."""
        print("\nSearching for MQL folder...")
        print("Available folders:")
        self.list_all_folders(root_folder, indent=1)
        
        # Try to find MQL folder
        mql_folder = self.find_folder_recursive(root_folder, "1. MQL")
        
        if not mql_folder:
            print("\n✗ Could not find 'MQL' folder automatically.")
            return self.manual_folder_selection(root_folder)
        
        print(f"\n✓ Found MQL folder: {mql_folder.Name}")
        print("\nSubfolders in MQL:")
        
        # Get subfolders
        subfolders = {}
        try:
            for subfolder in mql_folder.Folders:
                name = subfolder.Name.strip()
                name_upper = name.upper()
                print(f"  - {name}")
                
                # Map folder names
                if "ARROW" in name_upper:
                    subfolders["ARROW"] = subfolder
                elif "FUTURE" in name_upper:
                    subfolders["FUTURE"] = subfolder
                elif "RUTRONIK" in name_upper:
                    subfolders["RUTRONIK"] = subfolder
                elif "OTHER" in name_upper and "DISTRIBUTION" in name_upper:
                    subfolders["OTHER DISTRIBUTION PARTNERS"] = subfolder
                elif "EBV" in name_upper or "AVNET" in name_upper:
                    subfolders["EBV/AVNET"] = subfolder
                elif "NON" in name_upper and "EBV" in name_upper:
                    subfolders["NON EBV LEADS"] = subfolder
        except Exception as e:
            print(f"Error reading subfolders: {e}")
            return {}
        
        return subfolders
    
    def manual_folder_selection(self, root_folder):
        """Allow user to manually navigate to MQL folder."""
        print("\nLet's navigate to the MQL folder manually.")
        current = root_folder
        path = [current]
        
        while True:
            print(f"\nCurrent: {current.Name}")
            
            try:
                subfolders = list(current.Folders)
            except:
                subfolders = []
            
            if not subfolders:
                print("No subfolders found.")
                if len(path) > 1:
                    cmd = input("Press 'u' to go up, 'q' to quit: ").strip().lower()
                    if cmd == 'u':
                        path.pop()
                        current = path[-1]
                    else:
                        return {}
                else:
                    return {}
                continue
            
            print("Subfolders:")
            for i, folder in enumerate(subfolders, 1):
                print(f"  [{i}] {folder.Name}")
            
            cmd = input("\nEnter number to open, 's' if this is MQL folder, 'u' to go up, 'q' to quit: ").strip().lower()
            
            if cmd == 's':
                # This is the MQL folder, get its subfolders
                subfolders_dict = {}
                for subfolder in subfolders:
                    name = subfolder.Name.strip()
                    name_upper = name.upper()
                    
                    if "ARROW" in name_upper:
                        subfolders_dict["ARROW"] = subfolder
                    elif "FUTURE" in name_upper:
                        subfolders_dict["FUTURE"] = subfolder
                    elif "RUTRONIK" in name_upper:
                        subfolders_dict["RUTRONIK"] = subfolder
                    elif "OTHER" in name_upper and "DISTRIBUTION" in name_upper:
                        subfolders_dict["OTHER DISTRIBUTION PARTNERS"] = subfolder
                
                return subfolders_dict
            elif cmd == 'u':
                if len(path) > 1:
                    path.pop()
                    current = path[-1]
            elif cmd == 'q':
                return {}
            else:
                try:
                    idx = int(cmd) - 1
                    if 0 <= idx < len(subfolders):
                        current = subfolders[idx]
                        path.append(current)
                except:
                    print("Invalid input.")
    
    def determine_target_folder(self, row: Dict[str, str]) -> tuple:
        """Determine which folder the email should be moved to.
        
        Returns:
            tuple: (folder_name, action_description) or (None, reason_for_skip)
        """
        has_contact_form = row.get("Has Contact Sales Form", "").strip()
        potential_partner = row.get("Potential Distribution Partner (matching in beta testing)", "").strip()
        
        # Only move if Has Contact Sales Form = No
        if has_contact_form != "No":
            return (None, f"Skipped - Has Contact Sales Form: {has_contact_form}")
        
        # Normalize potential partner value
        partner_upper = potential_partner.upper()
        
        # Skip if empty or just "?"
        if not potential_partner or potential_partner == "?":
            return (None, f"Skipped - No distribution partner identified")
        
        # Check for specific partners
        if "ARROW" in partner_upper:
            return ("ARROW", f"Moved to Arrow folder - Partner: {potential_partner}")
        elif "FUTURE" in partner_upper:
            return ("FUTURE", f"Moved to Future folder - Partner: {potential_partner}")
        elif "RUTRONIK" in partner_upper:
            return ("RUTRONIK", f"Moved to Rutronik folder - Partner: {potential_partner}")
        elif "AVNET" in partner_upper:
            return (None, f"Skipped - AVNET partner")
        else:
            # Any other distribution partner
            return ("OTHER DISTRIBUTION PARTNERS", f"Moved to Other Distribution Partners folder - Partner: {potential_partner}")
        
        return (None, "Skipped - No matching criteria")
    
    def move_email(self, email_item, target_folder):
        """Move an email to the target folder."""
        try:
            email_item.Move(target_folder)
            return True
        except Exception as e:
            print(f"Error moving email: {e}")
            return False
    
    def process_emails(self, emails: List, parsed_data: List[Dict], subfolders: Dict):
        """Process and move emails based on rules.
        
        Returns:
            dict: Mapping of email index to status information
        """
        status_map = {}  # Maps index to (status, action)
        
        if not subfolders:
            print("No subfolders available for moving emails")
            # Mark all as not processed
            for i in range(len(emails)):
                status_map[i] = ("Not Started", "Email moving disabled - No MQL subfolders found")
            return status_map
        
        # Create folder mapping
        folder_map = {
            "ARROW": subfolders.get("ARROW"),
            "FUTURE": subfolders.get("FUTURE"),
            "RUTRONIK": subfolders.get("RUTRONIK"),
            "OTHER DISTRIBUTION PARTNERS": subfolders.get("OTHER DISTRIBUTION PARTNERS"),
        }
        
        print("\nFolder mapping:")
        for name, folder in folder_map.items():
            status = "✓" if folder else "✗"
            print(f"  {status} {name}: {folder.Name if folder else 'NOT FOUND'}")
        
        # Track statistics
        stats = {
            "ARROW": 0,
            "FUTURE": 0,
            "RUTRONIK": 0,
            "OTHER DISTRIBUTION PARTNERS": 0,
            "SKIPPED": 0,
            "FAILED": 0
        }
        
        print("\nProcessing emails for moving...")
        
        for i, (email_item, row_data) in enumerate(zip(emails, parsed_data)):
            target_name, action_desc = self.determine_target_folder(row_data)
            
            if not target_name:
                # Email was skipped
                stats["SKIPPED"] += 1
                status_map[i] = ("Not Started", action_desc)
                continue
            
            target_folder = folder_map.get(target_name)
            
            if not target_folder:
                # Folder not found
                action = f"Failed - Target folder '{target_name}' not found"
                print(f"Warning: {action}")
                stats["FAILED"] += 1
                status_map[i] = ("Failed", action)
                continue
            
            # Move the email
            subject = row_data.get("Subject", "Unknown")
            
            if self.move_email(email_item, target_folder):
                stats[target_name] += 1
                status_map[i] = ("Completed", action_desc)
                self.move_log.append({
                    "index": i,
                    "subject": subject,
                    "to_folder": target_name,
                    "status": "success"
                })
                print(f"  ✓ {action_desc[:70]}...")
            else:
                action = f"Failed - Could not move to {target_name}"
                stats["FAILED"] += 1
                status_map[i] = ("Failed", action)
                self.move_log.append({
                    "index": i,
                    "subject": subject,
                    "to_folder": target_name,
                    "status": "failed"
                })
        
        # Print summary
        print("\n" + "=" * 60)
        print("EMAIL MOVING SUMMARY")
        print("=" * 60)
        print(f"  Arrow:                       {stats['ARROW']}")
        print(f"  Future:                      {stats['FUTURE']}")
        print(f"  Rutronik:                    {stats['RUTRONIK']}")
        print(f"  Other Distribution Partners: {stats['OTHER DISTRIBUTION PARTNERS']}")
        print(f"  Skipped:                     {stats['SKIPPED']}")
        print(f"  Failed:                      {stats['FAILED']}")
        print("=" * 60)
        
        return status_map