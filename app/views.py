import os
import pandas as pd
from django.http import HttpResponse
from django.shortcuts import render
from django.conf import settings
from io import BytesIO
import zipfile


def upload_and_process_excel(request):
    if request.method == 'POST':
        # Retrieve uploaded file
        uploaded_file = request.FILES['excel_file']

        # Load the Excel file using Pandas
        try:
            df = pd.read_excel(uploaded_file, sheet_name='GST_DETAIL', skiprows=5)
        except Exception as e:
            return HttpResponse(f"Error reading the file: {e}")

        # Clean column names: remove all non-alphabetic characters and strip whitespace
        df.columns = df.columns.str.replace(r'[^a-zA-Z]', '', regex=True).str.strip()
        # print(df.columns)
        # to modify SGST column
        df['SGST'] = pd.to_numeric(df['SGST'], errors='coerce') 
        df['SGST'] = df['SGST'] * 2.00 # multiple the sgst tax amount * 2.00
        df['SGST_Amount'] = 1 # new SGST Amount column created

        # to create CGST column 
        df['CGST'] = pd.to_numeric(df['CGST'], errors='coerce') 
        df['CGST'] = df['CGST'] * 2.00 # multiple the cgst tax amount * 2.00
        df['CGST_Amount'] = 1 # new CGST Amount column created


        # Extract relevant columns (adjust column names if necessary)
        invoice_data = df[['InvoiceDate', 'Desc', 'GSTIN', 'InvoiceNo', 'TaxableAmount', 'SGST','SGST_Amount','CGST','CGST_Amount','TotalGST']].copy()

        # to calculate SGST amount by using taxable amount
        invoice_data['SGST_Amount'] = df['SGST'] / 100 * df['TaxableAmount']


        # to calculate CGST amount by using taxable amount
        invoice_data['CGST_Amount'] = df['CGST'] / 100 * df['TaxableAmount']

        # Rename columns for better readability
        invoice_data.rename(columns={'Desc': 'PartyName'}, inplace=True)

        # Fill missing values using ffill()
        invoice_data[['PartyName', 'InvoiceDate', 'GSTIN', 'InvoiceNo']] = invoice_data[[
            'PartyName', 'InvoiceDate', 'GSTIN', 'InvoiceNo']].ffill()


        # calculate total gst amount by addind cgst + sgst amount 
        invoice_data['TotalGST'] = invoice_data['CGST_Amount'] + invoice_data['SGST_Amount']



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
            # Create in-memory buffer for zip file
            in_memory_zip = BytesIO()

            with zipfile.ZipFile(in_memory_zip, mode='w', compression=zipfile.ZIP_DEFLATED) as zip_file:
                # Create in-memory Excel files for part1 and part2
                part1_excel = BytesIO()
                part2_excel = BytesIO()

                part1.to_excel(part1_excel, index=False)
                part2 = part2.drop(columns=columns_to_remove, errors='ignore')
                part2.to_excel(part2_excel, index=False)

                # Save Excel files in the zip archive
                part1_excel.seek(0)
                zip_file.writestr('B2B.xlsx', part1_excel.read())

                part2_excel.seek(0)
                zip_file.writestr('B2C.xlsx', part2_excel.read())

            # Set the buffer's position to the beginning
            in_memory_zip.seek(0)

            # Return the zip file for download
            response = HttpResponse(in_memory_zip.read(), content_type='application/zip')
            response['Content-Disposition'] = 'attachment; filename="processed_data.zip"'
            return response

        else:
            return HttpResponse("No 'B2C' found in the data. No split performed.")

    # Render upload form
    return render(request, 'upload_excel.html')