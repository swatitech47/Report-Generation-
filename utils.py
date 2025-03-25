import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def process_excel(file_path, input_sheet_name):
    """
    Reads an Excel file, processes data by country and month, 
    and writes output to new sheets in the same file.
    """
    print(f"📂 Loading data from: {file_path}")
    
    try:
        df = pd.read_excel(file_path, sheet_name=input_sheet_name)
    except FileNotFoundError:
        print(f"❌ Error: File '{file_path}' not found!")
        return
    except ValueError:
        print(f"❌ Error: Sheet '{input_sheet_name}' not found in '{file_path}'!")
        return

    df.columns = df.columns.str.strip()  # Trim spaces before and after column names

    print("✅ Data Loaded Successfully!")
    print("🔹 Columns in dataset:", df.columns.tolist())

    # Check if dataframe has data
    if df.empty:
        print("❌ Error: Input sheet is empty!")
        return

    # Define column names
    country_col = "Country"
    month_col = "Month Name"
    units_col = "Units Sold"
    sales_col = "Sales"
    profit_col = "Profit"
    discount_col = "Discounts"

    # Define month sorting order
    month_order = {
        "January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
        "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12
    }

    # Load existing workbook
    book = load_workbook(file_path)

    # Remove previously created sheets if they exist
    for sheet in ["Country_Total", "Month_Average"]:
        if sheet in book.sheetnames:
            del book[sheet]
    book.save(file_path)

    # Open Excel writer
    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a") as writer:
        # Group by country and write each country's data to a new sheet
        country_groups = df.groupby(country_col)
        for country, group in country_groups:
            sheet_name = str(country)[:31]  # Sheet name limit is 31 characters
            print(f"📊 Writing data for country: {country}")
            group.to_excel(writer, sheet_name=sheet_name, index=False)

        # Compute country-wise totals
        country_totals = country_groups.agg({
            units_col: 'sum',
            sales_col: 'sum',
            profit_col: 'sum'
        }).reset_index()
        print("📌 Writing 'Country_Total' sheet...")
        country_totals.to_excel(writer, sheet_name="Country_Total", index=False)

        # Compute month-wise totals and averages
        month_groups = df.groupby(month_col).agg({
            units_col: 'sum',
            sales_col: 'mean',
            profit_col: 'mean',
            discount_col: 'mean'
        }).reset_index()

        # Sort by month order
        month_groups["Month_Order"] = month_groups[month_col].map(month_order)
        month_groups = month_groups.sort_values(by="Month_Order").drop(columns=["Month_Order"])

        print("📌 Writing 'Month_Average' sheet...")
        month_groups.to_excel(writer, sheet_name="Month_Average", index=False)

    print("✅ Data processing completed!")

def apply_alternating_row_colors(file_path, input_sheet_name):
    """
    Applies alternating row colors (blue and white) to newly created sheets.
    """
    print(f"🎨 Applying alternating row colors to: {file_path}")

    # Reload workbook
    book = load_workbook(file_path)
    blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

    # Apply coloring to new sheets
    for sheet_name in book.sheetnames:
        if sheet_name != input_sheet_name:  # Skip input sheet
            ws = book[sheet_name]
            for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=1):
                if row_idx % 2 == 0:  # Apply color to even rows
                    for cell in row:
                        cell.fill = blue_fill

    # Save the formatted file
    book.save(file_path)

    print("✅ Formatting applied successfully!")
