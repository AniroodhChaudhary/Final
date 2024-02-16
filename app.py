from flask import Flask, request, send_file, render_template
import openpyxl
import fitz  # PyMuPDF
import re

app = Flask(__name__)

    # Add your existing functions here
def parse_transactions_from_pdf(pdf_path):
    # Your existing code
    # ... (Your existing code for parsing transactions)
    try:
        with fitz.open(pdf_path) as pdf:
            text = ''
            for page in pdf:
                text += page.get_text()

            # Adjusted regex pattern to match the format in the image
            pattern = r'OPTSTK\s+(\w+)\s+(\d{2}[a-zA-Z]{3}\d{2})\s+([0-9.]+)\s+(PE|CE)\s+\[\*\w\]\s+(B|S)\s+(-?\d+)\s+([0-9.]+)'
            optstk_transactions = re.findall(pattern, text)

        

        return optstk_transactions
    except Exception as e:
        print(f"An error occurred in parse_transactions_from_pdf: {e}")
        # Optionally, you can re-raise the exception or return a default value
        raise e  # or return []

col_indices = [9, 10, 12, 13]

def is_only_one_cell_filled(row, col_indices):
    """
    Check if only one out of the specified columns in a row is filled.
    
    :param row: The row object from openpyxl.
    :param col_indices: List of column indices to check.
    :return: True if only one cell in the specified columns is filled, False otherwise.
    """
    filled_cells = [row[col - 1].value for col in col_indices if row[col - 1].value is not None]
    return len(filled_cells) == 1
def update_excel_with_transactions(excel_path, transactions, contract_date):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    for trans in transactions:
        # Unpack the transaction details
        # ... (Your existing code for appending transactions
        scrip, expiry, strike_price, option_type, bs_indicator,quantity, rate = trans

        # Convert strings to appropriate numeric types
        quantity = abs(int(quantity))# Make quantity positive
        rate = float(rate)
        strike_price=float(strike_price)

        # New logic to check for existing transactions before appending
        found_matching_row = False
        for row in ws.iter_rows(min_row=2, values_only=False):
            # Your logic to find a matching row
            # If found, set found_matching_row to True and update the row
            # Check for the same scrip, strike price and option type
            if row[5].value == scrip and row[8].value == strike_price and row[4].value == quantity:
                if is_only_one_cell_filled(row, col_indices):
                    if row[9].value:
                        ft=row[9]
                        existing_bs='S'
                        existing_ot='CE'
                    elif row[10].value:
                        ft=row[10]
                        existing_bs='S'
                        existing_ot='PE'
                    elif row[12].value:
                        ft=row[12]
                        existing_bs='B'
                        existing_ot='CE'
                    elif row[13].value:
                        ft=row[13]
                        existing_bs='B'
                        existing_ot='PE'
                # Check if the existing transaction is a complement (buy for sell, sell for buy)
                    if option_type == existing_ot and existing_bs != bs_indicator:
                        if option_type == 'CE' and bs_indicator == 'S':
                            cell_to_up=row[9]
                        elif option_type == 'PE' and bs_indicator == 'S':
                            cell_to_up=row[10]
                        elif option_type == 'CE' and bs_indicator == 'B':
                            cell_to_up=row[12]
                        elif option_type == 'PE' and bs_indicator == 'B':
                            cell_to_up=row[13]
                        cell_to_up.value = rate
                        row[15].value=''
                        if option_type == 'CE':
                            row[14].value= (row[9].value-row[12].value)*row[4].value
                        elif option_type == 'PE':
                            row[14].value= (row[10].value-row[13].value)*row[4].value
                        found_matching_row = True
                        break

        if not found_matching_row:
            new_row = [''] * 16
            new_row[3] = contract_date  # Date
            new_row[4] = quantity  # Quantity
            new_row[5] = scrip  # Scrip
            new_row[8] = strike_price  # Strike Price
            if option_type == 'CE' and bs_indicator == 'B':
                new_row[12] = rate  # Buy Call
            elif option_type == 'CE' and bs_indicator == 'S':
                new_row[9] = rate  # Sell Call
                new_row[15] = new_row[9]*new_row[4]
            elif option_type == 'PE' and bs_indicator == 'B':
                new_row[13] = rate  # Buy Put
            elif option_type == 'PE' and bs_indicator == 'S':
                new_row[10] = rate  # Sell Put
                new_row[15] = new_row[10]*new_row[4]
            ws.append(new_row)  # Append new transaction
    updated_excel_path = 'updated_' + excel_path
    wb.save(updated_excel_path)

    return updated_excel_path


@app.route('/')
def index():
    return render_template('upload_form.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    # Receive files from the user
    pdf_file = request.files['pdf']
    excel_file = request.files['excel']

    # Save files temporarily
    pdf_path = 'contract.pdf'
    excel_path = 'transaction.xlsx'
    pdf_file.save(pdf_path)
    excel_file.save(excel_path)

    # Process files
    transactions = parse_transactions_from_pdf(pdf_path)

    def extract_contract_date(pdf_path):
        with fitz.open(pdf_path) as pdf:
            text = ""
            for page in pdf:
                text += page.get_text()

            # Search for the Trade Date in the text
            trade_date_match = re.search(r'Trade Date\s*(\d{2}/\d{2}/\d{2})', text)
            trade_date = trade_date_match.group(1)
        return trade_date

    trade_date = extract_contract_date(pdf_path)

    # Function to convert date format from "dd/mm/yy" to "dd.mm.yy"
    def convert_date_format(date_str):
        return date_str.replace('/', '.')

    contract_date = convert_date_format(trade_date)

    updated_excel_path=update_excel_with_transactions(excel_path, transactions, contract_date)

    # Send back the processed Excel file
    return send_file(updated_excel_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=False,host='0.0.0.0')

