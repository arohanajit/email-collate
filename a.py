import os
import pdfplumber
import pandas as pd
import re
import traceback
import json

def extract_pdf_data(file_path):
    print(f"Processing PDF: {file_path}")
    try:
        with pdfplumber.open(file_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text()

        # Extract Delivery Date
        delivery_date = None
        order_no = None
        units_3127 = 0
        units_1327 = 0
        gross = 0
        net = 0

        lines = text.split('\n')
        for i, line in enumerate(lines):
            try:
                if "Delivery Date:" in line:
                    delivery_date = line.split()[-1]
                if "Order No:" in line:
                    order_no = line.split(":")[-1].strip()
                elif "OD-" in line:
                    potential_order_no = re.search(r'OD-\d+', line)
                    if potential_order_no:
                        order_no = potential_order_no.group(0)
                if "3127-PRM ETH 10 CARB RFG" in line:
                    parts = line.split()
                    units_3127 += float(re.sub(r'[^0-9.]', '', parts[-3]))
                if "1327-UNL ETH 10 CARB RFG" in line:
                    parts = line.split()
                    units_1327 += float(re.sub(r'[^0-9.]', '', parts[6]))
                if "paid by" in line:
                    net = float(re.sub(r'[^0-9.]', '', line.split()[-2][3:]))
                if "Invoice Total" in line:
                    gross = float(re.sub(r'[^0-9.]', '', line.split()[-1]))
            except Exception as e:
                print(f"Error processing line {i+1} in {file_path}: {str(e)}")
                print(f"Line content: {line.split()}")

        savings = round(gross - net, 2)
        folder_name = os.path.basename(os.path.dirname(file_path))
        print(f"File: {file_path} in folder: {folder_name}, delivery_date: {delivery_date} order_no: {order_no} gross: {gross} net: {net} units_3127: {units_3127} units_1327: {units_1327}")

        print(f"Successfully extracted data from: {file_path}")
        return os.path.basename(file_path), folder_name, delivery_date, order_no, gross, net, units_3127, units_1327, savings
    except Exception as e:
        print(f"Error processing {file_path}:")
        print(traceback.format_exc())
        return None

def process_pdfs(path1, path2, processed_files):
    pdf_files = []
    for path in [path1, path2]:
        print(f"Searching for PDF files in: {path}")
        for file in os.listdir(path):
            if file.lower().endswith('.pdf'):
                full_path = os.path.join(path, file)
                if full_path not in processed_files:
                    pdf_files.append(full_path)
    
    print(f"Found {len(pdf_files)} new PDF files to process")
    print(pdf_files)

    data_list = []
    for pdf_file in pdf_files:
        data = extract_pdf_data(pdf_file)
        if data:
            data_list.append(data)
            processed_files.append(pdf_file)

    return data_list, processed_files

def update_excel(data_list, excel_path):
    print(f"Updating Excel file: {excel_path}")
    try:
        if os.path.exists(excel_path):
            df = pd.read_excel(excel_path, header=None)
        else:
            df = pd.DataFrame()

        new_rows = []
        for data in data_list:
            new_row = [
                data[0],  # Filename
                data[1],  # Folder Name
                data[2],  # Delivery Date
                data[3],  # Order No
                data[7],  # Regular (Gallon)
                data[6],  # Super (Gallon)
                data[7]+data[6],  # Total Gallon
                data[4],  # Gross
                data[5],  # Net
                data[8]   # Savings
            ]
            new_rows.append(new_row)

        new_df = pd.DataFrame(new_rows)
        df = pd.concat([df, new_df], ignore_index=True)

        # Convert 'Delivery Date' column to datetime
        df[2] = pd.to_datetime(df[2], format='%m/%d/%Y', errors='coerce')

        # Sort the dataframe by 'Delivery Date'
        df = df.sort_values(by=2)

        # Convert 'Delivery Date' back to string format
        df[2] = df[2].dt.strftime('%m/%d/%Y')

        # Save the sorted dataframe to Excel
        df.to_excel(excel_path, index=False, header=False)
        print(f"Successfully updated and sorted Excel file with {len(new_rows)} new rows")
        return df
    except Exception as e:
        print(f"Error updating Excel file:")
        print(traceback.format_exc())

def load_processed_files(json_path):
    if os.path.exists(json_path):
        with open(json_path, 'r') as f:
            return json.load(f)
    return []

def save_processed_files(processed_files, json_path):
    with open(json_path, 'w') as f:
        json.dump(processed_files, f)

def process_and_update(path1, path2, excel_path):
    json_path = 'processed_files.json'
    processed_files = load_processed_files(json_path)
    
    data_list, processed_files = process_pdfs(path1, path2, processed_files)
    
    if data_list:
        print("Starting Excel update...")
        update_excel(data_list, excel_path)
        print(f"Excel file updated successfully at {excel_path}")
    else:
        print("No new data extracted from PDFs. Excel file not updated.")
    
    save_processed_files(processed_files, json_path)
    print(f"Updated list of processed files saved to {json_path}")