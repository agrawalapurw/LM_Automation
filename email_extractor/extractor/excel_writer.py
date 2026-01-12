import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting.rule import FormulaRule


# Dropdown options for Validation sheet
TAKE_ACTION_VALIDATION = [
    "Valid Company → MQL",
    "Valid Company → Reject",
    "Invalid Company"
]

VALID_REJECT_REASONS_VALIDATION = [
    "Not able to contact",
    "Not a Disti lead",
    "Contacted - no general potential",
    "Contacted - no current potential",
    "Not contacted - no general potential",
    "Insufficient lead profile information",
    "Lead already known",
    "University Contact",
    "Distribution Partner"
]

INVALID_COMPANY_REASONS = [
    "University Contact",
    "Distribution Partner",
    "Company - with no potential",
    "Agency",
    "Free-mailer with no potential",
    "Competitor"
]

# Dropdown options for Review sheet
TAKE_ACTION_REVIEW = [
    "MQL - Send to Sales",
    "Reject"
]

REJECT_REASONS_REVIEW = [
    "Not able to contact",
    "Not a Disti lead",
    "Contacted - no general potential",
    "Contacted - no current potential",
    "Not contacted - no general potential",
    "Insufficient lead profile information",
    "Lead already known",
    "University Contact",
    "Distribution Partner"
]

# Workflow columns for each sheet
VALIDATION_WORKFLOW_COLS = [
    "Take Action",
    "Valid Company → Reject Reason",
    "Invalid Company Reason",
    "Additional Scoring Information",
    "Send to",
    "Status",
    "Action Taken"
]

REVIEW_WORKFLOW_COLS = [
    "Take Action",
    "Reject Reason",
    "Additional Scoring Information",
    "Send to",
    "Status",
    "Action Taken"
]


class ExcelWriter:
    """Write two DataFrames to Excel with separate sheets and dropdowns."""
    
    def write_workbook(self, df_validation: pd.DataFrame, df_review: pd.DataFrame, filepath: str):
        """Write two DataFrames to Excel with separate sheets."""
        wb = Workbook()
        
        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
        
        # Write Validation sheet
        if not df_validation.empty:
            ws_val = wb.create_sheet("Validation")
            self._write_validation_sheet(ws_val, df_validation)
        
        # Write Review sheet
        if not df_review.empty:
            ws_rev = wb.create_sheet("Review")
            self._write_review_sheet(ws_rev, df_review)
        
        wb.save(filepath)
    
    def _write_validation_sheet(self, worksheet, df: pd.DataFrame):
        """Write Validation sheet with special columns and dropdowns."""
        
        # Add validation-specific workflow columns if they don't exist
        for col in VALIDATION_WORKFLOW_COLS:
            if col not in df.columns:
                df[col] = ""
        
        # Remove review-specific columns if they exist
        if "Reject Reason" in df.columns:
            df = df.drop(columns=["Reject Reason"])
        
        # Reorder columns: data columns first, then workflow columns at the end
        data_columns = [col for col in df.columns if col not in VALIDATION_WORKFLOW_COLS]
        ordered_columns = data_columns + VALIDATION_WORKFLOW_COLS
        df = df[ordered_columns]
        
        # Write headers
        headers = list(df.columns)
        worksheet.append(headers)
        
        # Write data rows
        for _, row in df.iterrows():
            worksheet.append([row[col] for col in headers])
        
        # Get column indices for workflow columns
        take_action_idx = headers.index("Take Action") + 1
        valid_reject_idx = headers.index("Valid Company → Reject Reason") + 1
        invalid_idx = headers.index("Invalid Company Reason") + 1
        additional_info_idx = headers.index("Additional Scoring Information") + 1
        send_to_idx = headers.index("Send to") + 1
        
        last_row = worksheet.max_row
        
        # Add dropdowns
        print(f"Adding dropdowns to Validation sheet (rows 2-{last_row})...")
        
        # Take Action dropdown
        take_action_letter = get_column_letter(take_action_idx)
        dv1 = DataValidation(
            type="list",
            formula1=f'"' + ','.join(TAKE_ACTION_VALIDATION) + '"',
            allow_blank=True
        )
        worksheet.add_data_validation(dv1)
        dv1.add(f"{take_action_letter}2:{take_action_letter}{last_row}")
        
        # Valid Company → Reject Reason dropdown
        valid_reject_letter = get_column_letter(valid_reject_idx)
        dv2 = DataValidation(
            type="list",
            formula1=f'"' + ','.join(VALID_REJECT_REASONS_VALIDATION) + '"',
            allow_blank=True
        )
        worksheet.add_data_validation(dv2)
        dv2.add(f"{valid_reject_letter}2:{valid_reject_letter}{last_row}")
        
        # Invalid Company Reason dropdown
        invalid_letter = get_column_letter(invalid_idx)
        dv3 = DataValidation(
            type="list",
            formula1=f'"' + ','.join(INVALID_COMPANY_REASONS) + '"',
            allow_blank=True
        )
        worksheet.add_data_validation(dv3)
        dv3.add(f"{invalid_letter}2:{invalid_letter}{last_row}")
        
        # Add conditional formatting
        self._add_conditional_formatting_validation(
            worksheet,
            take_action_idx,
            valid_reject_idx,
            invalid_idx,
            additional_info_idx,
            send_to_idx,
            last_row
        )
        
        # Format headers
        self._format_headers(worksheet)
        
        # Freeze header row
        worksheet.freeze_panes = "A2"
        
        # Auto-adjust column widths
        self._adjust_column_widths(worksheet)
    
    def _write_review_sheet(self, worksheet, df: pd.DataFrame):
        """Write Review sheet with Take Action and Reject Reason dropdowns."""
        
        # Add review-specific workflow columns if they don't exist
        for col in REVIEW_WORKFLOW_COLS:
            if col not in df.columns:
                df[col] = ""
        
        # Remove validation-specific columns if they exist
        validation_specific = [
            "Valid Company → Reject Reason",
            "Invalid Company Reason"
        ]
        df = df.drop(columns=[col for col in validation_specific if col in df.columns])
        
        # Reorder columns: data columns first, then workflow columns at the end
        data_columns = [col for col in df.columns if col not in REVIEW_WORKFLOW_COLS]
        ordered_columns = data_columns + REVIEW_WORKFLOW_COLS
        df = df[ordered_columns]
        
        # Write headers
        headers = list(df.columns)
        worksheet.append(headers)
        
        # Write data rows
        for _, row in df.iterrows():
            worksheet.append([row[col] for col in headers])
        
        # Get column indices for workflow columns
        take_action_idx = headers.index("Take Action") + 1
        reject_reason_idx = headers.index("Reject Reason") + 1
        additional_info_idx = headers.index("Additional Scoring Information") + 1
        send_to_idx = headers.index("Send to") + 1
        
        last_row = worksheet.max_row
        
        # Add dropdowns
        print(f"Adding dropdowns to Review sheet (rows 2-{last_row})...")
        
        # Take Action dropdown
        take_action_letter = get_column_letter(take_action_idx)
        dv1 = DataValidation(
            type="list",
            formula1=f'"' + ','.join(TAKE_ACTION_REVIEW) + '"',
            allow_blank=True
        )
        worksheet.add_data_validation(dv1)
        dv1.add(f"{take_action_letter}2:{take_action_letter}{last_row}")
        
        # Reject Reason dropdown
        reject_reason_letter = get_column_letter(reject_reason_idx)
        dv2 = DataValidation(
            type="list",
            formula1=f'"' + ','.join(REJECT_REASONS_REVIEW) + '"',
            allow_blank=True
        )
        worksheet.add_data_validation(dv2)
        dv2.add(f"{reject_reason_letter}2:{reject_reason_letter}{last_row}")
        
        # Add conditional formatting
        self._add_conditional_formatting_review(
            worksheet,
            take_action_idx,
            reject_reason_idx,
            additional_info_idx,
            send_to_idx,
            last_row
        )
        
        # Format headers
        self._format_headers(worksheet)
        
        # Freeze header row
        worksheet.freeze_panes = "A2"
        
        # Auto-adjust column widths
        self._adjust_column_widths(worksheet)
    
    def _add_conditional_formatting_validation(self, worksheet, take_action_col, 
                                               valid_reject_col, invalid_col, 
                                               additional_info_col, send_to_col, last_row):
        """Add conditional formatting to Validation sheet."""
        
        # Colors
        highlight_yellow = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        highlight_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        
        take_action_letter = get_column_letter(take_action_col)
        valid_reject_letter = get_column_letter(valid_reject_col)
        invalid_letter = get_column_letter(invalid_col)
        additional_letter = get_column_letter(additional_info_col)
        send_to_letter = get_column_letter(send_to_col)
        
        # Highlight "Valid Company → Reject Reason" when Take Action = "Valid Company → Reject"
        for row in range(2, last_row + 1):
            rule1 = FormulaRule(
                formula=[f'${take_action_letter}{row}="Valid Company → Reject"'],
                stopIfTrue=False,
                fill=highlight_yellow
            )
            worksheet.conditional_formatting.add(f"{valid_reject_letter}{row}", rule1)
        
        # Highlight "Invalid Company Reason" when Take Action = "Invalid Company"
        for row in range(2, last_row + 1):
            rule2 = FormulaRule(
                formula=[f'${take_action_letter}{row}="Invalid Company"'],
                stopIfTrue=False,
                fill=highlight_yellow
            )
            worksheet.conditional_formatting.add(f"{invalid_letter}{row}", rule2)
        
        # Highlight "Additional Scoring Information" and "Send to" when Take Action = "Valid Company → MQL"
        for row in range(2, last_row + 1):
            rule3 = FormulaRule(
                formula=[f'${take_action_letter}{row}="Valid Company → MQL"'],
                stopIfTrue=False,
                fill=highlight_green
            )
            worksheet.conditional_formatting.add(f"{additional_letter}{row}", rule3)
            worksheet.conditional_formatting.add(f"{send_to_letter}{row}", rule3)
    
    def _add_conditional_formatting_review(self, worksheet, take_action_col,
                                          reject_reason_col, additional_info_col,
                                          send_to_col, last_row):
        """Add conditional formatting to Review sheet."""
        
        # Colors
        highlight_yellow = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        highlight_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        
        take_action_letter = get_column_letter(take_action_col)
        reject_reason_letter = get_column_letter(reject_reason_col)
        additional_letter = get_column_letter(additional_info_col)
        send_to_letter = get_column_letter(send_to_col)
        
        # Highlight "Reject Reason" when Take Action = "Reject"
        for row in range(2, last_row + 1):
            rule1 = FormulaRule(
                formula=[f'${take_action_letter}{row}="Reject"'],
                stopIfTrue=False,
                fill=highlight_yellow
            )
            worksheet.conditional_formatting.add(f"{reject_reason_letter}{row}", rule1)
        
        # Highlight "Additional Scoring Information" and "Send to" when Take Action = "MQL - Send to Sales"
        for row in range(2, last_row + 1):
            rule2 = FormulaRule(
                formula=[f'${take_action_letter}{row}="MQL - Send to Sales"'],
                stopIfTrue=False,
                fill=highlight_green
            )
            worksheet.conditional_formatting.add(f"{additional_letter}{row}", rule2)
            worksheet.conditional_formatting.add(f"{send_to_letter}{row}", rule2)
    
    def _format_headers(self, worksheet):
        """Format header row with bold text and background color."""
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
        
        # Set header row height
        worksheet.row_dimensions[1].height = 30
    
    def _adjust_column_widths(self, worksheet):
        """Auto-adjust column widths based on content."""
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            
            adjusted_width = min(max_length + 3, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width