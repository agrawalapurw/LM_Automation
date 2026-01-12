import os
from datetime import timedelta
import pandas as pd
from extractor.outlook import OutlookClient
from extractor.parser import EmailParser
from extractor.excel_writer import ExcelWriter
from extractor.email_mover import EmailMover


DEFAULT_FILTERS = ["Pre-MQL ready for review", "Pre-MQL ready for validation"]


def get_date_label(ranges):
    """Generate filename label from date ranges."""
    if not ranges:
        return "selection"
    
    sorted_ranges = sorted(ranges, key=lambda x: x[0])
    
    if len(sorted_ranges) == 1:
        start, end = sorted_ranges[0]
        if (end - start) == timedelta(days=1):
            return start.strftime("%Y_%m_%d")
        return f"{start.strftime('%Y_%m_%d')}_to_{(end - timedelta(days=1)).strftime('%Y_%m_%d')}"
    
    dates = [start.strftime("%Y_%m_%d") for start, _ in sorted_ranges]
    return "+".join(dates)


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
    print("\nConnecting to Outlook...")
    
    # Initialize components
    outlook = OutlookClient()
    parser = EmailParser()
    
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
    
    # Create DataFrame
    df = pd.DataFrame(rows)
    
    # Split by subject type
    df_validation = df[df["Subject"].str.contains("validation", case=False, na=False)].copy()
    df_review = df[df["Subject"].str.contains("review", case=False, na=False)].copy()
    
    if df_validation.empty and df_review.empty:
        print("No emails matched 'validation' or 'review' subjects.")
        return
    
    # Save to Excel
    output_dir = ensure_output_dir()
    date_label = get_date_label(date_ranges)
    filename = get_unique_path(output_dir, f"{date_label}_PreMQL")
    
    writer = ExcelWriter()
    writer.write_workbook(df_validation, df_review, filename)
    
    print(f"\n✓ Saved to: {filename}")
    print(f"  - Validation sheet: {len(df_validation)} rows")
    print(f"  - Review sheet: {len(df_review)} rows")
    
    # Move emails if requested
    if move_emails:
        mover = EmailMover(outlook)
        
        # Get MQL subfolders
        subfolders = mover.get_mql_subfolders(store)
        
        if subfolders:
            print(f"\nFound {len(subfolders)} subfolders in MQL")
            mover.process_emails(emails, rows, subfolders)
        else:
            print("\nCould not find MQL subfolders. Skipping email moving.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user.")
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()