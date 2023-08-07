import pandas as pd

# Read the Excel file with the first row as the header
df = pd.read_excel('OpEx.xlsx', sheet_name='Data Import Template', skiprows=range(0, 14), header=0)

# Remove additional rows
rows_to_remove = [0, 1, 2, 3]  # Example row numbers to remove
df = df.drop(rows_to_remove)

# Add column names to remove from here
columns_to_remove = ['I Accept ', 'initialized', 'Exited', 'Reviewed']
columns_to_fill = ['BU', 'HU', 'Client/Account Name', 'Thread/Sub-Process', 'Process/LoB', 'BCSLA Code', 'Actual Start Date', 'Exit Report Date ']

# Iterate over each column
for column in columns_to_fill:
    df[column] = df[column].fillna(method='ffill')

df = df.drop(columns=columns_to_remove)

# Save the modified sheet as a new Excel file
df.to_excel('modified.xlsx', index=False)
