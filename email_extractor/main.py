"""
Main Entry Point
Email extraction and processing workflow.
"""

import os
from datetime import timedelta
import pandas as pd
from extractor.outlook import OutlookClient
from extractor.parser import EmailParser
from extractor.excel_writer import ExcelWriter
from extractor.email_mover import EmailMover
from extractor.domain_validator import DomainValidator
from extractor.university_detector import UniversityDetector
from extractor.validation_data import ValidationDataLoader

DEFAULT_FILTERS = ["Pre-MQL ready for review", "Pre-MQL ready for validation"]

def get_date_label(ranges):
    """Generate filename label from date ranges."""
    if not ranges:
        return "Extraction_selection"
    
    sorted_ranges = sorted(ranges, key=lambda x: x[0])
    
    if len(sorted_ranges) == 1:
        start, end = sorted_ranges[0]
        if (end - start) == timedelta(days=1):
            return f"Extraction_{start.strftime('%d%b%y')}"
        
        end_inclusive = end - timedelta(days=1)
        
        if start.month == end_inclusive.month and start.year == end_inclusive.year:
            return f"Extraction_{start.strftime('%d')}to{end_inclusive.strftime('%d%b%y')}"
        else:
            return f"Extraction_{start.strftime('%d%b')}to{end_inclusive.strftime('%d%b%y')}"
    
    first_date = sorted_ranges[0][0]
    last_date = sorted_ranges[-1][0]
    
    if first_date.month == last_date.month and first_date.year == last_date.year:
        return f"Extraction_{first_date.strftime('%d')}to{last_date.strftime('%d%b%y')}"
    else:
        return f"Extraction_{first_date.strftime('%d%b')}to{last_date.strftime('%d%b%y')}"

def ensure_output_dir():
    """Create output directory if it doesn't exist."""
    output_dir = os.path.join(os.path.dirname(__file__), "output")
    os.makedirs(output_dir, exist_ok=True)
    return output_dir

def get_unique_path(directory, base_name):
    """Generate unique filename if file exists."""
    path = os.path.join(directory, f"{base_name}.xlsx")
    if not os.path.exists(path):
        return path
    
    counter = 2
    while True:
        path = os.path.join(directory, f"{base_name} ({counter}).xlsx")
        if not os.path.exists(path):
            return path
        counter += 1

def main():
    print("=" * 60)
    print("EMAIL EXTRACTOR - Pre-MQL Tool")
    print("=" * 60)
    
    # Load validation data
    print("\nLoading validation data...")
    validation_data_dir = os.path.join(os.path.dirname(__file__), "validation_data")
    validation_loader = ValidationDataLoader(validation_data_dir)
    
    print("\nConnecting to Outlook...")
    
    # Initialize components
    outlook = OutlookClient()
    domain_validator = DomainValidator(validation_loader)
    university_detector = UniversityDetector(validation_loader)
    parser = EmailParser(university_detector, validation_loader)
    
    # Select store and folder
    store = outlook.select_store()
    folder = outlook.select_folder(store)
    
    # Get date range
    print("\nEnter date(s):")
    print("  Single: 2024-01-15")
    print("  Multiple: 2024-01-15,2024-01-16")
    print("  Range: 2024-01-15 to 2024-01-20")
    date_input = input("Date(s): ").strip()
    date_ranges = outlook.parse_date_input(date_input)
    
    # Get subject filters
    use_default = input("\nUse default subject filters? [Y/n]: ").strip().lower()
    if use_default in ("", "y", "yes"):
        filters = DEFAULT_FILTERS
    else:
        custom = input("Enter keywords (comma-separated, blank for no filter): ").strip()
        filters = [x.strip() for x in custom.split(",") if x.strip()] if custom else []
    
    # Ask about moving emails
    should_move = input("\nMove emails to distribution partner folders? [Y/n]: ").strip().lower()
    move_emails = should_move in ("", "y", "yes")
    
    # Fetch and parse emails
    print("\nFetching emails...")
    emails = outlook.fetch_emails(folder, date_ranges, filters)
    
    if not emails:
        print("No matching emails found.")
        return
    
    print(f"Found {len(emails)} emails. Parsing...")
    rows = [parser.parse_email(email) for email in emails]
    
    # Move emails if requested
    status_map = {}
    if move_emails:
        mover = EmailMover(outlook)
        subfolders = mover.get_mql_subfolders(store)
        
        if subfolders:
            print(f"\nFound {len(subfolders)} subfolders in MQL")
            status_map = mover.process_emails(emails, rows, subfolders)
        else:
            print("\nCould not find MQL subfolders. Skipping email moving.")
            for i in range(len(rows)):
                status_map[i] = ("Not Started", "Email moving was skipped")
    else:
        for i in range(len(rows)):
            status_map[i] = ("Not Started", "Email moving was not requested")
    
    # Update rows with status (preserve protected statuses)
    protected_statuses = ["University Contact", "Completed", "Academic", "Excluded Domain", 
                          "Direct Account", "Country"]
    
    for i, row in enumerate(rows):
        current_status = row.get("Status", "")
        
        if current_status not in protected_statuses:
            if i in status_map:
                status, action = status_map[i]
                row["Status"] = status
                row["Action Taken"] = action
            elif not current_status:
                row["Status"] = "Not Started"
                row["Action Taken"] = "No action taken"
    
    # Create DataFrame
    df = pd.DataFrame(rows)
    
    # Validate company domains
    print("\nValidating company domains...")
    validation_results = []
    for _, row in df.iterrows():
        company = row.get("Company", "")
        email = row.get("Email Address", "")
        result = domain_validator.validate_domain(company, email)
        validation_results.append(result["status"])
    
    df["Company Domain Validation"] = validation_results
    
    # Split by subject type
    df_validation = df[df["Subject"].str.contains("validation", case=False, na=False)].copy()
    df_review = df[df["Subject"].str.contains("review", case=False, na=False)].copy()
    
    if df_validation.empty and df_review.empty:
        print("No emails matched 'validation' or 'review' subjects.")
        return
    
    # Mark Mass Market in Review sheet
    if not df_review.empty:
        mass_market_mask = df_review["Account Type"].str.contains("mass market", case=False, na=False)
        
        mass_market_updated = 0
        for idx in df_review[mass_market_mask].index:
            current_status = df_review.at[idx, "Status"]
            if current_status not in protected_statuses:
                df_review.at[idx, "Status"] = "Mass Market"
                df_review.at[idx, "Action Taken"] = "Identified as Mass Market account"
                mass_market_updated += 1
        
        if mass_market_updated > 0:
            print(f"\n✓ Identified {mass_market_updated} Mass Market accounts in Review sheet")
    
    # Save to Excel
    output_dir = ensure_output_dir()
    date_label = get_date_label(date_ranges)
    filename = get_unique_path(output_dir, f"{date_label}_PreMQL")
    
    writer = ExcelWriter()
    writer.write_workbook(df_validation, df_review, filename)
    
    print(f"\n✓ Saved to: {filename}")
    print(f"  - Validation sheet: {len(df_validation)} rows")
    print(f"  - Review sheet: {len(df_review)} rows")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user.")
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()