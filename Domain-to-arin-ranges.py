# bash pip install pandas openpyxl ipwhois

# python
import pandas as pd
from ipwhois import IPWhois
import socket

# Path to the Excel file
excel_file = 'path/to/your/excel_file.xlsx'

# Read the Excel file
df = pd.read_excel(excel_file)

# Ensure the domain column exists in the dataframe
if 'domain' not in df.columns:
    raise ValueError("The Excel file must contain a 'domain' column.")

# Lists to store the results
ip_ranges = []
arin_numbers = []

# Function to get IP range and ARIN number for a domain
def get_ip_info(domain):
    try:
        ip = socket.gethostbyname(domain)
        obj = IPWhois(ip)
        res = obj.lookup_rdap()
        ip_range = res['network']['cidr']
        arin_number = res['asn_registry']
        return ip_range, arin_number
    except Exception as e:
        print(f"Error processing domain {domain}: {e}")
        return None, None

# Process each domain
for domain in df['domain']:
    ip_range, arin_number = get_ip_info(domain)
    ip_ranges.append(ip_range)
    arin_numbers.append(arin_number)

# Add the results to the dataframe
df['IP Range'] = ip_ranges
df['ARIN Number'] = arin_numbers

# Write the results back to the same Excel file
df.to_excel(excel_file, index=False)

print(f"Updated Excel file saved to {excel_file}")


# Explanation:
#1. **Import Libraries**: Import the required libraries: `pandas`, `ipwhois`, and `socket`.
#2. **Read Excel File**: Load the Excel file using `pandas`.
#3. **Check for Domain Column**: Ensure the Excel file has a column named 'domain'.
#4. **Function to Get IP Info**: Define a function to get the IP range and ARIN number for a given domain using `socket` and `ipwhois`.
#5. **Process Each Domain**: Iterate through each domain, get the IP range and ARIN number, and store them in lists.
#6. **Add Results to DataFrame**: Add the IP range and ARIN number as new columns to the DataFrame.
#7. **Save Back to Excel**: Write the updated DataFrame back to the same Excel file.

### Notes:
#- Ensure the domain names in your Excel file are in a column named 'domain'.
#- Handle exceptions where a domain might not resolve or where IPWhois lookup might fail.
#- Adjust the path to your Excel file in the script. 

#-This script will help you automate the process of obtaining IP ranges and ARIN numbers for a list of domain names and save the results back into the same Excel file.
