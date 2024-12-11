import pandas as pd
from datetime import datetime, timedelta

# File path (update to the correct path of your Excel file)
input_file = r"C:\Users\Michael Nguyen\Desktop\Python Billing Script\Venafi_Report.xlsx"

# Read the file into a DataFrame
try:
    df = pd.read_excel(input_file, engine="openpyxl")  # Use openpyxl for .xlsx files
    print("File loaded successfully!")
except Exception as e:
    print(f"Error loading file: {e}")
    exit()

# Debug: Print column names to ensure data is loaded correctly
print("Columns in the input file:", df.columns.tolist())

# Initialize the SSL Charges Report structure
ssl_report = pd.DataFrame()

# Dynamically calculate Start Time and End Time
today = datetime.today()
first_of_this_month = datetime(today.year, today.month, 1)
first_of_last_month = first_of_this_month - timedelta(days=1)
start_time = first_of_last_month.replace(day=1)
end_time = start_time.replace(year=start_time.year + 1)

# Assign dynamically calculated dates
ssl_report["Start Time"] = [start_time.strftime("%m/%d/%Y")] * len(df)
ssl_report["End Time"] = [end_time.strftime("%m/%d/%Y")] * len(df)

# Map and transform columns
ssl_report["Activity Code"] = 1869  # Fixed code
ssl_report["Server Name"] = df["Server Name"].str.upper()  # Convert to uppercase
ssl_report["Last Modifier User ID"] = ""  # Blank
ssl_report["Task Code"] = 4120  # Fixed code

# Add more transformations and calculations as required...

# Save the output
output_file = r"C:\Users\Michael Nguyen\Desktop\Python Billing Script\SSL_Charges_Report.xlsx"
ssl_report.to_excel(output_file, index=False)
print(f"Transformed report saved to {output_file}")
