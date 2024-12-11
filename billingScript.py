import pandas as pd
import os
import logging
from datetime import datetime, timedelta

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# File paths
input_file = r"C:\Users\Michael Nguyen\Desktop\Python Billing Script\Venafi_Report.xlsx"

# Function to prevent overwriting by adding numerical suffix
def get_unique_filename(base_name, extension=".xlsx", directory="."):
    counter = 1
    filename = os.path.join(directory, f"{base_name}{extension}")
    while os.path.exists(filename):
        filename = os.path.join(directory, f"{base_name}({counter}){extension}")
        counter += 1
    return filename

# Configurable fallback for environments without prompts
DEFAULT_NAME = "Processed_Report"
DEFAULT_DIRECTORY = "."  # Current directory

try:
    # Prompt the user for a custom file name
    custom_name = input("Enter a custom name for the output file (leave blank for default): ").strip()
    if not custom_name:
        custom_name = DEFAULT_NAME  # Default name

    # Prompt the user for a custom file location
    file_location = input("Enter the directory to save the output file (leave blank for current directory): ").strip()
    if file_location:
        # Validate the directory
        if not os.path.isdir(file_location):
            logging.error(f"Invalid directory: {file_location}. Defaulting to current directory.")
            file_location = DEFAULT_DIRECTORY
    else:
        # Default to current directory
        file_location = DEFAULT_DIRECTORY
except:
    logging.warning("Input prompts are not supported. Using default values.")
    custom_name = DEFAULT_NAME
    file_location = DEFAULT_DIRECTORY

output_file = get_unique_filename(custom_name, directory=file_location)
logging.info(f"The output file will be saved as: {output_file}")

if not os.path.exists(input_file):
    logging.error(f"File not found: {input_file}")
    exit()

try:
    df = pd.read_excel(input_file, engine="openpyxl")
    logging.info("File loaded successfully.")
except Exception as e:
    logging.error(f"Error loading file: {e}")
    exit()

def calculate_end_time(start_date_str):
    try:
        start_date = datetime.strptime(start_date_str, "%m/%d/%Y")
        next_month = start_date.replace(day=1) + timedelta(days=32)
        end_date = next_month.replace(year=next_month.year + 1, day=1)
        return end_date.strftime("%m/%d/%Y")
    except Exception as e:
        logging.error(f"Error calculating end time: {e}")
        return None

start_time_input = input("Enter Start Time (mm/dd/yyyy): ").strip()
try:
    start_time = datetime.strptime(start_time_input, "%m/%d/%Y").strftime("%m/%d/%Y")
    end_time = calculate_end_time(start_time)
    logging.info(f"Start Time: {start_time}, End Time: {end_time}")
except ValueError:
    logging.error("Invalid date format. Using default Start Time: 06/01/2024.")
    start_time = "06/01/2024"
    end_time = "07/01/2025"

df['Start Time'] = start_time
df['End Time'] = end_time
df['Activity Code'] = 1869
df['Server Name'] = df['Server Name'].str.upper()
df['Last Modifier User ID'] = None
df['Task Code'] = 4120
df['Charge Account Number'] = df['Charge Account Number. Must start with a G, A, or P']

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

df['Amount'] = df['SANs (DNS)'].apply(calculate_amount)

def generate_entry_comment(sans_column):
    num_sans = sans_column.count(",") + 1
    if num_sans <= 3:
        return "One Year Certificate SSL billed one time"
    return f"One Year Certificate with {num_sans} SAN Certs"

df['Entry Comment'] = df['SANs (DNS)'].apply(generate_entry_comment)
df['Billing Description'] = "SSL " + df['Nickname'].str.upper()
df['Subject Alternative Name'] = df['SANs (DNS)']
df['Status'] = 'Issued'
df['Entry Time'] = None
df['Entry ID'] = None
df['Issue Date'] = pd.to_datetime(df['Valid From']).dt.strftime('%m/%d/%Y')
df['Expired Date'] = pd.to_datetime(df['Valid To']).dt.strftime('%m/%d/%Y')
df['Customer'] = df['Billing Contact Name']

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
            return "Manual Correction Needed"
    else:
        return "Manual Correction Needed"

df['Department'] = df['Contact'].apply(generate_department)


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
