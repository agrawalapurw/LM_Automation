import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


class ExcelWriter:
    def write_workbook(self, df_validation: pd.DataFrame, df_review: pd.DataFrame, filepath: str):
        """Write two DataFrames to Excel with separate sheets."""
        wb = Workbook()
        
        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
        
        # Write Validation sheet
        if not df_validation.empty:
            ws_val = wb.create_sheet("Validation")
            self._write_dataframe(ws_val, df_validation)
        
        # Write Review sheet
        if not df_review.empty:
            ws_rev = wb.create_sheet("Review")
            self._write_dataframe(ws_rev, df_review)
        
        wb.save(filepath)
    
    def _write_dataframe(self, worksheet, df: pd.DataFrame):
        """Write DataFrame to worksheet."""
        for row in dataframe_to_rows(df, index=False, header=True):
            worksheet.append(row)
        
        # Freeze header row
        worksheet.freeze_panes = "A2"
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width