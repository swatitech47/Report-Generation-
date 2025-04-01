import json
import mysql.connector
from utils import process_excel, apply_alternating_row_colors, FoodDataProcessor

class MainProcessor:
    def __init__(self, file_path, input_sheet_name, json_file, db_config):
        """Initialize with file path, sheet name, JSON file, and database file"""
        self.file_path = file_path
        self.input_sheet_name = input_sheet_name
        self.json_file = json_file
        self.db_config = db_config

    def execute(self):
        """Executes Excel data processing, JSON handling, and database update"""
        print(f"Processing file: {self.file_path}")
        process_excel(self.file_path, self.input_sheet_name)
        apply_alternating_row_colors(self.file_path, self.input_sheet_name)
        print("✅ Excel processing and formatting completed!")

        print(f"Processing JSON file: {self.json_file}")
        
        json_processor = FoodDataProcessor(self.json_file, self.db_config)
        json_processor.process_data()  # Process and store JSON data in MySQL
        print("✅ JSON processing completed!")

if __name__ == "__main__":
    file_path = "FinancialSample.xlsx"  # Excel file for financial data
    input_sheet_name = "Sheet1"
    json_file = "food_json.json"  # JSON file with food data

    # Database configuration (adjust these values as needed)
    db_config = {
        'host': 'localhost',  # MySQL server address
        'database': 'food_data',  # MySQL database name
        'user': 'root'  # MySQL username
    
    }

    processor = MainProcessor(file_path, input_sheet_name, json_file, db_config)
    processor.execute()
