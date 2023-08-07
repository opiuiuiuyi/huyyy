import pandas as pd

# cd to Music then /Music$ source myenv/bin/activate

# Read the Excel file with the first row as the header
df = pd.read_excel('OpEx.xlsx', sheet_name='Data Import Template', skiprows=range(0, 14), header=0)

# Remove additional rows
rows_to_remove = [0, 1, 2,3]  # Example row numbers to remove
df = df.drop(rows_to_remove)

#add column names to remove from here 
columns_to_remove = ['I Accept ','initialized', 'Exited', 'Reviewed']

columns_to_fill = ['BU', 'HU','Client/Account Name','Thread/Sub-Process','Process/LoB','BCSLA Code','Actual Start Date','Exit Report Date ']

# Iterate over each column
for column in columns_to_fill:
    df[column] = df[column].fillna(method='ffill')



df = df.drop(columns=columns_to_remove)
# Save the modified sheet as a new Excel file

columns_major_nc = [ 'BU', 'HU','Client/Account Name','Thread/Sub-Process','Process/LoB','BCSLA Code','Actual Start Date','Exit Report Date ','Module(Internal)','Module.22','Practice.21','Question.10','Opportunity Description.10','Finding Detail.10','Severity.10','Closure Date (Enter Manually After Validation)']

columns_minor_nc = [ 'BU', 'HU','Client/Account Name','Thread/Sub-Process','Process/LoB','BCSLA Code','Actual Start Date','Exit Report Date ','Module(Internal).1','Module.23','Practice.22','Question.11','Opportunity Description.11','Finding Detail.11','Severity.11','Closure Date (Enter Manually After Validation).1']

columns_obs = [ 'BU', 'HU','Client/Account Name','Thread/Sub-Process','Process/LoB','BCSLA Code','Actual Start Date','Exit Report Date ','Module(Internal).2','Module.24','Practice.23','Question.12','Opportunity Description.12','Finding Detail.12','Severity.12','Closure Date (Enter Manually After Validation).2']

columns_ofi = [ 'BU', 'HU','Client/Account Name','Thread/Sub-Process','Process/LoB','BCSLA Code','Actual Start Date','Exit Report Date ','Module(Internal).3','Module.25','Practice.24','Question.13','Opportunity Description.13','Finding Detail.13','Severity.13','Closure Date (Enter Manually After Validation).3']

columns_closed = [ 'BU', 'HU','Client/Account Name','Thread/Sub-Process','Process/LoB','BCSLA Code','Actual Start Date','Exit Report Date ','Module(Internal).6' ,'Module.28','Practice.27','Question.16','Opportunity Description.16','Finding Detail.16','Severity.16','Closure Date (Enter Manually After Validation).6']


writer = pd.ExcelWriter('exported_data.xlsx', engine='xlsxwriter')

# Export the "Major NC" data to a new sheet
df_major_nc = df[columns_major_nc]
df_major_nc.to_excel(writer, sheet_name='Major NC', index=False)

# Export the "Minor NC" data to a new sheet
df_minor_nc = df[columns_minor_nc]
df_minor_nc.to_excel(writer, sheet_name='Minor NC', index=False)

# Export the "OBS" data to a new sheet
df_obs = df[columns_obs]
df_obs.to_excel(writer, sheet_name='OBS', index=False)

# Export the "Major NC" data to a new sheet
df_ofi = df[columns_ofi]
df_ofi.to_excel(writer, sheet_name='OFI', index=False)

df_ofi = df[columns_closed]
df_ofi.to_excel(writer, sheet_name='Closed', index=False)

# Save the changes and close the writer
writer._save()


# Read the Excel file with the first row as the header
df = pd.read_excel('exported_data.xlsx', sheet_name=None)


# Create a new Excel writer using xlsxwriter engine
with pd.ExcelWriter('Newsheet.xlsx', engine='xlsxwriter') as writer:
    # Iterate over each sheet
    for sheet_name, sheet_df in df.items():
        if sheet_name == 'Major NC':
            # Remove rows where "Question.10" is blank
            sheet_df = sheet_df.dropna(subset=['Question.10'])
        elif sheet_name == 'Minor NC':
            # Remove rows where "Question.11" is blank
            sheet_df = sheet_df.dropna(subset=['Question.11'])
        elif sheet_name == 'OBS':
            # Remove rows where "Question.12" is blank
            sheet_df = sheet_df.dropna(subset=['Question.12'])
        elif sheet_name == 'OFI':
            # Remove rows where "Question.13" is blank
            sheet_df = sheet_df.dropna(subset=['Question.13'])
        elif sheet_name == 'Closed':
            # Remove rows where "Question.13" is blank
            sheet_df = sheet_df.dropna(subset=['Question.16'])
        
        # Save the modified sheet to the Excel file
        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

