import pandas as pd

# Load the Excel file
file_path = 'a.xlsx'  # Replace with the actual file path
df = pd.read_excel(file_path, sheet_name='GST_DETAIL', skiprows=5)

# to modify SGST column
df['SGST'] = pd.to_numeric(df['SGST'], errors='coerce') 
df['SGST'] = df['SGST'] * 2.00 # multiple the sgst tax amount * 2.00
df['SGST_Amount'] = 1 # new SGST Amount column created

# to create CGST column 
df['CGST'] = pd.to_numeric(df['CGST'], errors='coerce') 
df['CGST'] = df['CGST'] * 2.00 # multiple the cgst tax amount * 2.00
df['CGST_Amount'] = 1 # new CGST Amount column created


# Extract relevant columns (adjust column names if necessary)
invoice_data = df[['Invoice Date', 'Desc', 'GSTIN', 'Invoice No', 'Taxable Amount', 'SGST','SGST_Amount','CGST','CGST_Amount','Total GST']].copy()

# to calculate SGST amount by using taxable amount
invoice_data['SGST_Amount'] = df['SGST'] / 100 * df['Taxable Amount']


# to calculate CGST amount by using taxable amount
invoice_data['CGST_Amount'] = df['CGST'] / 100 * df['Taxable Amount']

# Rename columns for better readability
invoice_data.rename(columns={'Desc': 'PartyName'}, inplace=True)

# Fill missing values using ffill()
invoice_data[['PartyName', 'Invoice Date', 'GSTIN', 'Invoice No']] = invoice_data[[
    'PartyName', 'Invoice Date', 'GSTIN', 'Invoice No']].ffill()


# calculate total gst amount by addind cgst + sgst amount 
invoice_data['Total GST'] = invoice_data['CGST_Amount'] + invoice_data['SGST_Amount']



b2c_index = invoice_data[invoice_data['PartyName'].astype(str).str.contains('B2C', case=False, na=False)].index


# Remove rows containing 'B2C (Large) Invoice' or 'B2C (Small) Invoice'
invoice_data = invoice_data[~invoice_data['PartyName'].str.contains(
    r'B2C \(Large\) Invoice|B2C \(Small\) Invoice', case=False, na=False
)]



# Check if 'B2C' was found
if not b2c_index.empty:
    # Get the index of 'B2C' and the following rows
    b2c_start = b2c_index[0]

    # Split the data into B2B and B2C sections
    part1 = invoice_data.iloc[:b2c_start].reset_index(drop=True)  # B2B
    part2 = invoice_data.iloc[b2c_start:].reset_index(drop=True)  # B2C
    # print(part2)
    columns_to_remove = ['GSTIN']
    part2 = part2.drop(columns=columns_to_remove, errors='ignore')

    # Validate GSTINs in B2B data
    invalid_gstin = part1[part1['GSTIN'].isna() | (part1['GSTIN'].str.len() != 15)]
    # print(invalid_gstin)

    # Print invalid GSTIN rows (optional for debugging)
    # print("Invalid GSTIN rows:")
    # print(invalid_gstin)
    # print(part2)
    # Move invalid GSTIN entries to B2C data
    if not invalid_gstin.empty:
        part2 = pd.concat([part2, invalid_gstin], ignore_index=True)
        part1 = part1.drop(invalid_gstin.index).reset_index(drop=True)  # Remove invalid entries from B2B part
    # words that are removed if they present in partyname
    part1 = part1[~part1['PartyName'].str.contains(
    r'Net Total|B2B|Nil Rated/Exempted|Export Invoices|Set/off Tax on Advance of prior period|Gross Total|Less: Credit/Debit Note & Refund Vouche|Tax Liability on Advance', case=False, na=False
)]
    part2 = part2[~part2['PartyName'].str.contains(
    r'Net Total|B2B|Nil Rated/Exempted|Export Invoices|Set/off Tax on Advance of prior period|Gross Total|Less: Credit/Debit Note & Refund Vouche|Tax Liability on Advance', case=False, na=False
)]
    
    
    # Save both parts to separate Excel files
    part1.to_excel('x1.xlsx', index=False)
    part2 = part2.drop(columns=columns_to_remove, errors='ignore')
    part2.to_excel('y1.xlsx', index=False)

    print("Data successfully split into 'B2B.xlsx' and 'B2C.xlsx'.")
else:
    print("No 'B2C' found in the 'PartyName' column. No split performed.")
