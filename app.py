from flask import Flask, request, render_template, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # Read the orders file first
        orders_df = pd.read_csv(filepath)

        # Load valid pincodes from pin.csv
        pin_path = os.path.join(UPLOAD_FOLDER, 'pin.csv')
        if not os.path.exists(pin_path):
            return "pin.csv not found. Please upload pin list as 'pin.csv' inside /uploads"

        pin_df = pd.read_csv(pin_path)
        pin_df.columns = pin_df.columns.str.strip()  # Remove any trailing spaces in column names
        valid_pins = set(pin_df['PINCODE'].dropna().astype(str).str.strip())

        # Add Delivery Availability Column
        orders_df['Cleaned Zip'] = orders_df['Shipping Zip'].astype(str).str.replace(r"[^\d]", "", regex=True).str.strip()

        orders_df['Delivery Available'] = orders_df['Cleaned Zip'].apply(
        lambda x: 'yes' if x in valid_pins else 'no'
            )

        # Continue with the rest of your logic...
        cleaned_df = pd.DataFrame({
            'SRNO': range(1, len(orders_df) + 1),
            'ORDER NO': orders_df['Name'],
            'ORDER DATE': pd.to_datetime(orders_df['Created at']).dt.strftime('%Y-%m-%d'),
            'MONTH': pd.to_datetime(orders_df['Created at']).dt.strftime("%b'%y"),
            'BRAND': 'TILTING HEADS',
            'CUSTOMER NAME': orders_df['Billing Name'],
            'ADDRESS': orders_df['Shipping Address1'].fillna('') + ', ' + orders_df['Shipping Address2'].fillna(''),
            'PINCODE': orders_df['Cleaned Zip'],
            'STATE': orders_df['Shipping Province Name'],
            'PHONE NUMBER': orders_df['Shipping Phone'].apply(lambda x: '{:.0f}'.format(x) if pd.notnull(x) else ''),
            'EMAIL ID': orders_df['Email'],
            'PRODUCT NAME': orders_df['Lineitem name'],
            'AMOUNT': orders_df['Total'],
            'COUNT OF ITEMS': orders_df['Lineitem quantity'],
            'PAYMENT MODE': orders_df['Financial Status'].apply(lambda x: 'Online Payment' if x == 'paid' else 'Cash on Delivery'),
            'COURIER COMPANY': '',
            'TRACKING NUMBER': '',
            'DELIVERY STATUS/DATE': '',
            'SOLD/RETURENED': '',
            'CUSTOMER FEEDBACK': '',
            'Delivery Available': orders_df['Delivery Available']
        })

        output_file = os.path.join(OUTPUT_FOLDER, 'cleaned_orders.xlsx')
        cleaned_df.to_excel(output_file, index=False)

        return render_template('download.html', filename='cleaned_orders.xlsx')

    return render_template('index.html')


@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
