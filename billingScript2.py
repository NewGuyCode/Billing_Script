import pandas as pd
import os
import logging
from datetime import datetime, timedelta

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# File paths
input_file = r"C:\Users\Michael Nguyen\Desktop\Python Billing Script\Venafi_Report.xlsx"
output_file = r"C:\Users\Michael Nguyen\Desktop\Python Billing Script\Processed_Report.xlsx"

logging.info("Starting script...")

# Check file existence
if not os.path.exists(input_file):
    logging.error(f"File not found: {input_file}")
    exit()

# Load the file
try:
    df = pd.read_excel(input_file, engine="openpyxl")
    logging.info("File loaded successfully.")
except Exception as e:
    logging.error(f"Error loading file: {e}")
    exit()

# Function to calculate end time based on start time
def calculate_end_time(start_date_str):
    try:
        start_date = datetime.strptime(start_date_str, "%m/%d/%Y")
        # Move to the first day of the next month, then add one year
        next_month = start_date.replace(day=1) + timedelta(days=32)
        end_date = next_month.replace(year=next_month.year + 1, day=1)
        return end_date.strftime("%m/%d/%Y")
    except Exception as e:
        logging.error(f"Error calculating end time: {e}")
        return None

# User-defined start time
start_time_input = input("Enter Start Time (mm/dd/yyyy): ").strip()  # Example: "11/01/2024"

try:
    # Validate user input and calculate dates
    start_time = datetime.strptime(start_time_input, "%m/%d/%Y").strftime("%m/%d/%Y")
    end_time = calculate_end_time(start_time)
    logging.info(f"Start Time: {start_time}, End Time: {end_time}")
except ValueError:
    logging.error("Invalid date format. Using default Start Time: 06/01/2024.")
    start_time = "06/01/2024"
    end_time = "07/01/2025"

# Initialize required columns
df['Start Time'] = start_time
df['End Time'] = end_time
df['Activity Code'] = 1869
df['Server Name'] = df['Server Name'].str.upper()
df['Last Modifier User ID'] = None
df['Task Code'] = 4120
df['Charge Account Number'] = df['Charge Account Number. Must start with a G, A, or P']

# Define calculation function for the 'Amount' column
def calculate_amount(sans_column):
    try:
        if "*" in sans_column:
            num_wildcards = sans_column.count("*")
            return num_wildcards * 6 * 340
        else:
            num_sans = sans_column.count(",") + 1
            if num_sans <= 3:
                return 340
            else:
                return (num_sans - 3) * 340 + 340
    except Exception as e:
        logging.warning(f"Error calculating amount for SANs: {sans_column} - {e}")
        return 0

# Apply the calculation to the 'SANs (DNS)' column
df['Amount'] = df['SANs (DNS)'].apply(calculate_amount)

# Entry Comment Logic
def generate_entry_comment(sans_column):
    num_sans = sans_column.count(",") + 1
    if num_sans <= 3:
        return "One Year Certificate SSL billed one time"
    return f"One Year Certificate with {num_sans} SAN Certs"

df['Entry Comment'] = df['SANs (DNS)'].apply(generate_entry_comment)

# Billing Description
df['Billing Description'] = "SSL " + df['Nickname'].str.upper()

# Copy SANs (DNS) to Subject Alternative Name
df['Subject Alternative Name'] = df['SANs (DNS)']

# Status
df['Status'] = 'Issued'

# Blank Columns
df['Entry Time'] = None
df['Entry ID'] = None

# Use Valid From and Valid To for Issue Date and Expired Date, formatting dates
df['Issue Date'] = pd.to_datetime(df['Valid From']).dt.strftime('%m/%d/%Y')
df['Expired Date'] = pd.to_datetime(df['Valid To']).dt.strftime('%m/%d/%Y')

# Customer and Department
df['Customer'] = df['Billing Contact Name']

# Updated Department Logic
def generate_department(contact):
    if "AD+hosted" in contact.lower():
        return "Manual Correction Needed"
    elif "local:app-venafi_" in contact.lower():
        try:
            # Extract the department name after the last underscore
            return contact.split("_")[-1]
        except IndexError:
            return "Unknown Department"
    else:
        return "Unknown Department"

df['Department'] = df['Contact'].apply(generate_department)

# Reorder columns based on the example
output_columns = [
    'Start Time', 'End Time', 'Activity Code', 'Server Name',
    'Last Modifier User ID', 'Task Code', 'Amount', 'Charge Account Number',
    'Entry Comment', 'Billing Description', 'Subject Alternative Name',
    'Status', 'Entry Time', 'Entry ID', 'Issue Date', 'Expired Date',
    'Customer', 'Department'
]

try:
    df = df[output_columns]
    df.to_excel(output_file, index=False, engine="openpyxl")
    logging.info(f"Processed file saved successfully at: {output_file}")
except Exception as e:
    logging.error(f"Error saving the file: {e}")
