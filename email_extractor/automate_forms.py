import os
import sys
from extractor.web_automation import WebFormAutomation


def main():
    print("=" * 60)
    print("WEB FORM AUTOMATION TOOL")
    print("=" * 60)
    print("\nThis tool will:")
    print("1. Read your Excel file")
    print("2. Navigate to each PreMQL link")
    print("3. Fill in the form with data from Excel columns")
    print("4. Submit each form")
    print("\nIMPORTANT: Make sure you've manually filled in all required columns first!")
    print("=" * 60)
    
    # Get Excel file path
    output_dir = os.path.join(os.path.dirname(__file__), "output")
    
    if not os.path.exists(output_dir):
        print(f"\n✗ Output directory not found: {output_dir}")
        return
    
    # List Excel files in output directory
    excel_files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx')]
    
    if not excel_files:
        print(f"\n✗ No Excel files found in {output_dir}")
        return
    
    print(f"\nAvailable Excel files:")
    for i, file in enumerate(excel_files, 1):
        print(f"  [{i}] {file}")
    
    # Select file
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
    
    # Select sheet
    print("\nWhich sheet to process?")
    print("  [1] Validation")
    print("  [2] Review")
    print("  [3] Both")
    
    while True:
        sheet_choice = input("\nSelect option: ").strip()
        if sheet_choice in ("1", "2", "3"):
            break
        print("Invalid selection.")
    
    # Headless mode option
    headless_input = input("\nRun in headless mode (no browser UI)? [y/N]: ").strip().lower()
    headless = headless_input in ("y", "yes")
    
    # Confirmation
    print("\n" + "=" * 60)
    print("READY TO START")
    print("=" * 60)
    print(f"File: {excel_file}")
    if sheet_choice == "1":
        print("Sheet: Validation")
    elif sheet_choice == "2":
        print("Sheet: Review")
    else:
        print("Sheets: Validation + Review")
    print(f"Headless: {headless}")
    print("=" * 60)
    
    confirm = input("\nProceed? [Y/n]: ").strip().lower()
    if confirm and confirm not in ("y", "yes"):
        print("Cancelled.")
        return
    
    # Start automation
    automation = WebFormAutomation(headless=headless)
    
    try:
        automation.start_browser()
        
        if sheet_choice in ("1", "3"):
            automation.process_excel_file(excel_path, "Validation")
        
        if sheet_choice in ("2", "3"):
            automation.process_excel_file(excel_path, "Review")
        
    except KeyboardInterrupt:
        print("\n\n⚠ Interrupted by user")
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        automation.stop_browser()
    
    print("\n✓ Automation completed!")


if __name__ == "__main__":
    main()