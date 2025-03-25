import pandas as pd
from utils import process_excel, apply_alternating_row_colors

class MainProcessor:
    def __init__(self, file_path, input_sheet_name):
        """Initialize with file path and sheet name"""
        self.file_path = file_path
        self.input_sheet_name = input_sheet_name

    def execute(self):
        """Executes data processing and formatting"""
        print(f"Processing file: {self.file_path}")
        process_excel(self.file_path, self.input_sheet_name)

        print(f"Applying formatting to: {self.file_path}")
        apply_alternating_row_colors(self.file_path, self.input_sheet_name)

        print("✅ Data processing and formatting completed successfully!")

if __name__ == "__main__":
    # ✅ You can change the filename here or pass it as an argument
    file_path = "FinancialSample.xlsx"  # Change if needed
    input_sheet_name = "Sheet1"  # Change to match your sheet name

    # Create an instance of MainProcessor and execute
    processor = MainProcessor(file_path, input_sheet_name)
    processor.execute()
