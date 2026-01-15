import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from typing import Dict, Optional
from openpyxl import load_workbook
from datetime import datetime
import os


class WebFormAutomation:
    """Automate filling web forms from Excel data."""
    
    def __init__(self, headless: bool = False):
        """Initialize the web automation with Chrome driver.
        
        Args:
            headless: If True, runs browser in headless mode (no UI)
        """
        self.driver = None
        self.headless = headless
        self.stats = {
            "processed": 0,
            "skipped": 0,
            "success": 0,
            "failed": 0
        }
        self.excel_path = None
        self.status_updates = {}  # Store status updates to write back
    
    def start_browser(self):
        """Start Chrome browser with local driver."""
        print("Starting browser (Chrome)...")
        
        try:
            options = webdriver.ChromeOptions()
            
            if self.headless:
                options.add_argument('--headless')
            
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            
            # Look for chromedriver.exe in project directory
            project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            chromedriver_path = os.path.join(project_dir, "chromedriver.exe")
            
            if not os.path.exists(chromedriver_path):
                print(f"\n✗ chromedriver.exe not found at: {chromedriver_path}")
                print("\nPlease download ChromeDriver:")
                print("1. Check your Chrome version at chrome://settings/help")
                print("2. Download matching driver from https://googlechromelabs.github.io/chrome-for-testing/")
                print("3. Extract chromedriver.exe to the project folder")
                raise FileNotFoundError("chromedriver.exe not found")
            
            service = Service(executable_path=chromedriver_path)
            
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.maximize_window()
            print("✓ Browser started")
            
        except Exception as e:
            print(f"✗ Error starting browser: {e}")
            raise
    
    def stop_browser(self):
        """Stop browser."""
        if self.driver:
            self.driver.quit()
            print("✓ Browser closed")
    
    def process_excel_file(self, excel_path: str, sheet_name: str):
        """Process an Excel sheet and fill forms for each row.
        
        Args:
            excel_path: Path to the Excel file
            sheet_name: Name of the sheet ("Validation" or "Review")
        """
        self.excel_path = excel_path
        self.status_updates[sheet_name] = {}
        
        print(f"\n{'='*60}")
        print(f"Processing {sheet_name} Sheet")
        print(f"{'='*60}\n")
        
        # Read Excel
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        
        total_rows = len(df)
        print(f"Total rows: {total_rows}\n")
        
        # Process each row
        for index, row in df.iterrows():
            row_num = index + 2  # Excel row number (accounting for header)
            excel_row = index  # DataFrame index
            
            print(f"\n--- Processing Row {row_num} ---")
            
            # Check if Take Action is empty
            take_action = str(row.get("Take Action", "")).strip()
            if not take_action or take_action == "nan":
                print(f"⊘ Skipping - Take Action is empty")
                self.status_updates[sheet_name][excel_row] = "Skipped - No Take Action"
                self.stats["skipped"] += 1
                continue
            
            # Get the link
            link = str(row.get("PreMQL review/validation link", "")).strip()
            if not link or link == "nan" or not link.startswith("http"):
                print(f"⊘ Skipping - Invalid or missing link")
                self.status_updates[sheet_name][excel_row] = "Skipped - No Link"
                self.stats["skipped"] += 1
                continue
            
            print(f"Link: {link[:50]}...")
            print(f"Take Action: {take_action}")
            
            # Prepare form data based on sheet type
            if sheet_name == "Validation":
                form_data = self._prepare_validation_data(row)
            else:  # Review
                form_data = self._prepare_review_data(row)
            
            # Fill the form
            success, error_msg = self._fill_form(link, form_data, sheet_name)
            
            if success:
                self.stats["success"] += 1
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.status_updates[sheet_name][excel_row] = f"✓ Success - {timestamp}"
                print(f"✓ Successfully processed row {row_num}")
            else:
                self.stats["failed"] += 1
                self.status_updates[sheet_name][excel_row] = f"✗ Failed - {error_msg}"
                print(f"✗ Failed to process row {row_num}: {error_msg}")
            
            self.stats["processed"] += 1
            
            # Small delay between rows
            time.sleep(1)
        
        # Write status updates back to Excel for this sheet
        self._update_excel_sheet(sheet_name)
        
        # Print summary
        self._print_stats()
    
    def _update_excel_sheet(self, sheet_name: str):
        """Write Form Submission Status back to Excel for a specific sheet."""
        if sheet_name not in self.status_updates or not self.status_updates[sheet_name]:
            return
        
        print(f"\nUpdating {sheet_name} sheet in Excel...")
        
        try:
            # Load workbook
            wb = load_workbook(self.excel_path)
            ws = wb[sheet_name]
            
            # Find the "Form Submission Status" column
            header_row = 1
            status_col = None
            
            for col in range(1, ws.max_column + 1):
                cell_value = ws.cell(row=header_row, column=col).value
                if cell_value == "Form Submission Status":
                    status_col = col
                    break
            
            if status_col is None:
                print("  ⊘ Could not find 'Form Submission Status' column")
                return
            
            # Update status for each row
            for excel_row, status in self.status_updates[sheet_name].items():
                excel_write_row = excel_row + 2  # +2 because header is row 1, and index starts at 0
                ws.cell(row=excel_write_row, column=status_col, value=status)
            
            # Save workbook
            wb.save(self.excel_path)
            print(f"  ✓ Updated {len(self.status_updates[sheet_name])} rows in Excel")
            
        except Exception as e:
            print(f"  ✗ Error updating Excel: {e}")
    
    def _prepare_validation_data(self, row: pd.Series) -> Dict[str, str]:
        """Prepare form data from Validation sheet row."""
        data = {
            "take_action": str(row.get("Take Action", "")).strip(),
            "reject_reason": str(row.get("Valid Company → Reject Reason", "")).strip(),
            "invalid_reason": str(row.get("Invalid Company Reason", "")).strip(),
            "additional_info": str(row.get("Additional Scoring Information", "")).strip(),
            "send_to": str(row.get("Send to", "")).strip(),
        }
        
        # Clean up 'nan' values
        for key in data:
            if data[key] == "nan":
                data[key] = ""
        
        return data
    
    def _prepare_review_data(self, row: pd.Series) -> Dict[str, str]:
        """Prepare form data from Review sheet row."""
        data = {
            "take_action": str(row.get("Take Action", "")).strip(),
            "reject_reason": str(row.get("Reject Reason", "")).strip(),
            "additional_info": str(row.get("Additional Scoring Information", "")).strip(),
            "send_to": str(row.get("Send to", "")).strip(),
        }
        
        # Clean up 'nan' values
        for key in data:
            if data[key] == "nan":
                data[key] = ""
        
        return data
    
    def _fill_form(self, url: str, form_data: Dict[str, str], sheet_type: str) -> tuple:
        """Navigate to URL and fill the form.
        
        Args:
            url: URL of the form
            form_data: Dictionary with form field values
            sheet_type: "Validation" or "Review"
        
        Returns:
            Tuple of (success: bool, error_message: str)
        """
        try:
            # Navigate to the page
            self.driver.get(url)
            
            # Wait for page to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            time.sleep(2)  # Additional wait for dynamic content
            
            # Fill form based on sheet type
            if sheet_type == "Validation":
                success, error = self._fill_validation_form(form_data)
            else:
                success, error = self._fill_review_form(form_data)
            
            return success, error
            
        except TimeoutException:
            return False, "Timeout loading page"
        except Exception as e:
            return False, str(e)[:50]
    
    def _fill_validation_form(self, data: Dict[str, str]) -> tuple:
        """Fill validation form fields.
        
        Note: This is a TEMPLATE. You need to update the selectors based on
        the actual HTML structure of your form.
        
        Returns:
            Tuple of (success: bool, error_message: str)
        """
        try:
            take_action = data["take_action"]
            
            # Example: Fill Take Action dropdown
            # REPLACE THESE SELECTORS WITH ACTUAL ONES FROM YOUR FORM
            self._fill_dropdown("select[name='take_action']", take_action)
            
            # Based on Take Action, fill corresponding fields
            if "MQL" in take_action:
                # Fill Additional Scoring Information
                self._fill_text_field("textarea[name='additional_info']", data["additional_info"])
                # Fill Send to
                self._fill_text_field("input[name='send_to']", data["send_to"])
            
            elif "Reject" in take_action:
                # Fill Reject Reason
                self._fill_dropdown("select[name='reject_reason']", data["reject_reason"])
            
            elif "Invalid" in take_action:
                # Fill Invalid Reason
                self._fill_dropdown("select[name='invalid_reason']", data["invalid_reason"])
            
            # Submit the form
            self._submit_form()
            
            return True, ""
            
        except Exception as e:
            return False, str(e)[:50]
    
    def _fill_review_form(self, data: Dict[str, str]) -> tuple:
        """Fill review form fields.
        
        Note: This is a TEMPLATE. You need to update the selectors based on
        the actual HTML structure of your form.
        
        Returns:
            Tuple of (success: bool, error_message: str)
        """
        try:
            take_action = data["take_action"]
            
            # Example: Fill Take Action dropdown
            # REPLACE THESE SELECTORS WITH ACTUAL ONES FROM YOUR FORM
            self._fill_dropdown("select[name='take_action']", take_action)
            
            # Based on Take Action, fill corresponding fields
            if "MQL" in take_action:
                # Fill Additional Scoring Information
                self._fill_text_field("textarea[name='additional_info']", data["additional_info"])
                # Fill Send to
                self._fill_text_field("input[name='send_to']", data["send_to"])
            
            elif "Reject" in take_action:
                # Fill Reject Reason
                self._fill_dropdown("select[name='reject_reason']", data["reject_reason"])
            
            # Submit the form
            self._submit_form()
            
            return True, ""
            
        except Exception as e:
            return False, str(e)[:50]
    
    def _fill_dropdown(self, selector: str, value: str):
        """Fill a dropdown field."""
        if not value:
            return
        
        try:
            element = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, selector))
            )
            select = Select(element)
            select.select_by_visible_text(value)
            print(f"  ✓ Filled dropdown: {value}")
        except Exception as e:
            print(f"  ⊘ Could not fill dropdown '{selector}': {e}")
            raise
    
    def _fill_text_field(self, selector: str, value: str):
        """Fill a text/textarea field."""
        if not value:
            return
        
        try:
            element = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, selector))
            )
            element.clear()
            element.send_keys(value)
            print(f"  ✓ Filled text field: {value[:30]}...")
        except Exception as e:
            print(f"  ⊘ Could not fill text field '{selector}': {e}")
            raise
    
    def _submit_form(self):
        """Submit the form."""
        try:
            # Option 1: Find submit button by type
            submit_button = self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
            submit_button.click()
            
            # Wait for submission to complete
            time.sleep(2)
            
            print(f"  ✓ Form submitted")
        except Exception as e:
            print(f"  ⊘ Could not submit form: {e}")
            raise
    
    def _print_stats(self):
        """Print processing statistics."""
        print(f"\n{'='*60}")
        print("PROCESSING SUMMARY")
        print(f"{'='*60}")
        print(f"Total processed:  {self.stats['processed']}")
        print(f"Successful:       {self.stats['success']}")
        print(f"Failed:           {self.stats['failed']}")
        print(f"Skipped:          {self.stats['skipped']}")
        print(f"{'='*60}\n")