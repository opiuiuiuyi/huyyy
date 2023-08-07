import pandas as pd

# Read the Excel file with the first row as the header
df = pd.read_excel('premerge.xlsx', sheet_name=None)

# Function to merge cells with serial numbering and line breaks
def merge_cells_with_serial(values):
    return '\n'.join(f"({i+1}) {value}" for i, value in enumerate(values))

def check_status_closed(group):
    if 'closed' in group['status'].values:
        return 'closed'
    else:
        return 'open'

def merge_rows(df):
    merged_df = pd.DataFrame()
    for _, group in df.groupby(['Practice', 'BCSLA Code', 'Severity']):
        status_check = check_status_closed(group)
        if len(group) > 1:
            merged_value_OpportunityDescription = merge_cells_with_serial(group['OpportunityDescription'])
            merged_value_finding = merge_cells_with_serial(group['Finding Detail'])
            group['OpportunityDescription'] = merged_value_OpportunityDescription
            group['Finding Detail'] = merged_value_finding
            group['status'] = status_check
            merged_df = pd.concat([merged_df, group.tail(1)])
        else:
            group['status'] = status_check
            merged_df = pd.concat([merged_df, group])
    return merged_df

# Process the "precombined" sheet
if 'precombined' in df:
    df['precombined'] = merge_rows(df['precombined'])

# Save the updated Excel file with all sheets
with pd.ExcelWriter('Final.xlsx', engine='xlsxwriter') as writer:
    for sheet_name, sheet_df in df.items():
        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
