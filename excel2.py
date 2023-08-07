import pandas as pd

# Read the Excel file with the first row as the header
df = pd.read_excel('Newsheet.xlsx', sheet_name=None)

# Create an ExcelWriter object to save the modified sheets
with pd.ExcelWriter('Newsheet.xlsx') as writer:
    # Iterate over each sheet and process data
    for sheet_name, sheet_df in df.items():
        if sheet_name in ['Major NC', 'Minor NC', 'OBS', 'OFI']:
            # Add a new column "status" with the value "open"
            sheet_df['status'] = 'open'
        elif sheet_name == 'Closed':
            # Add a new column "status" with the value "closed"
            sheet_df['status'] = 'closed'
        
        # Forward fill the values in the "status" column only if it exists
        if 'status' in sheet_df.columns:
            sheet_df['status'] = sheet_df['status'].ffill()

        # Save the modified sheet to the ExcelWriter object
        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)


import pandas as pd

# Read the Excel file with the first row as the header
df = pd.read_excel('Newsheet.xlsx', sheet_name=None)

# Define a mapping for the headers to be renamed in each sheet
headers_mapping = {
    'Major NC': {
        'Module.22': 'Module',
        'Practice.21': 'Practice',
        'Question.10': 'Question',
        'Finding Detail.10': 'Finding Detail',
        'Severity.10': 'Severity',
        'Opportunity Description.10' : 'OpportunityDescription'
    },
    'Minor NC': {
        'Module.23': 'Module',
        'Practice.22': 'Practice',
        'Question.11': 'Question',
        'Finding Detail.11': 'Finding Detail',
        'Severity.11': 'Severity',
        'Module(Internal).1': 'Module(Internal)',
        'Closure Date (Enter Manually After Validation).1' : 'Closure Date (Enter Manually After Validation)',
        'Opportunity Description.11' : 'OpportunityDescription'

    },
    'OBS': {
        'Module.24': 'Module',
        'Practice.23': 'Practice',
        'Question.12': 'Question',
        'Finding Detail.12': 'Finding Detail',
        'Severity.12': 'Severity',
        'Module(Internal).2': 'Module(Internal)',
        'Closure Date (Enter Manually After Validation).2' : 'Closure Date (Enter Manually After Validation)',
        'Opportunity Description.12' : 'OpportunityDescription'
    },
    'OFI': {
        'Module.25': 'Module',
        'Practice.24': 'Practice',
        'Question.13': 'Question',
        'Finding Detail.13': 'Finding Detail',
        'Severity.13': 'Severity',
        'Module(Internal).3': 'Module(Internal)',
        'Closure Date (Enter Manually After Validation).3' : 'Closure Date (Enter Manually After Validation)',
        'Opportunity Description.13' : 'OpportunityDescription'
    },
    'Closed': {
        'Module.28': 'Module',
        'Practice.27': 'Practice',
        'Question.16': 'Question',
        'Finding Detail.16': 'Finding Detail',
        'Severity.16': 'Severity',
        'Module(Internal).6': 'Module(Internal)',
        'Closure Date (Enter Manually After Validation).6' : 'Closure Date (Enter Manually After Validation)',
        'Opportunity Description.16' : 'OpportunityDescription'
    }
}

# Modify the headers in each sheet
for sheet_name, sheet_df in df.items():
    if sheet_name in headers_mapping:
        # Rename the specified headers
        sheet_df.rename(columns=headers_mapping[sheet_name], inplace=True)

# Save the modified sheets to the Excel file
with pd.ExcelWriter('Newsheet.xlsx') as writer:
    for sheet_name, sheet_df in df.items():
        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

import pandas as pd

# Replace 'your_excel_file.xlsx' with the actual path to your Excel file
excel_file = pd.ExcelFile('Newsheet.xlsx')

# List all sheet names in the Excel file
sheet_names = excel_file.sheet_names

# Initialize an empty DataFrame to store the combined rows
combined_data = pd.DataFrame()

# Iterate through each sheet and concatenate the data into 'combined_data'
for sheet_name in sheet_names:
    # Read the data from the sheet into a DataFrame
    df = excel_file.parse(sheet_name)
    
    # Concatenate the data to 'combined_data'
    combined_data = pd.concat([combined_data, df])

# Reset the index of 'combined_data'
combined_data.reset_index(drop=True, inplace=True)

# Create a new sheet in the Excel file with the combined data
with pd.ExcelWriter('premerge.xlsx', engine='openpyxl') as writer:
    combined_data.to_excel(writer, sheet_name='precombined', index=False)

print("Data combined and saved in 'precombined' sheet.")




