from flask import Flask, request, render_template, send_file
import pandas as pd
from datetime import datetime
import os
import re

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# 🔁 Clean repeating phrases from address
def clean_repetitive_phrases(text, min_ngram=3, max_ngram=6):
    if not isinstance(text, str):
        return text

    text = re.sub(r'\s+', ' ', text.strip().lower())  # normalize spaces and case
    words = text.split()
    seen_phrases = set()
    output_words = []
    i = 0

    while i < len(words):
        deduped = False
        for n in range(max_ngram, min_ngram - 1, -1):
            if i + n <= len(words):
                phrase = ' '.join(words[i:i + n])
                if phrase in seen_phrases:
                    deduped = True
                    i += n
                    break
                else:
                    seen_phrases.add(phrase)
        if not deduped:
            output_words.append(words[i])
            i += 1

    return ' '.join(output_words).capitalize()

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # Read orders CSV
        orders_df = pd.read_csv(filepath)

        # Load pin codes
        pin_path = os.path.join(UPLOAD_FOLDER, 'pin.csv')
        if not os.path.exists(pin_path):
            return "pin.csv not found. Please upload pin list as 'pin.csv' inside /uploads"

        pin_df = pd.read_csv(pin_path)
        pin_df.columns = pin_df.columns.str.strip()
        valid_pins = set(pin_df['PINCODE'].dropna().astype(str).str.strip())

        # Clean ZIP and check delivery
        orders_df['Cleaned Zip'] = orders_df['Shipping Zip'].astype(str).str.replace(r"[^\d]", "", regex=True).str.strip()
        orders_df['Delivery Available'] = orders_df['Cleaned Zip'].apply(
            lambda x: 'yes' if x in valid_pins else 'no'
        )

        # Build cleaned DataFrame
        cleaned_df = pd.DataFrame({
            'SRNO': range(1, len(orders_df) + 1),
            'ORDER NO': orders_df['Name'],
            'ORDER DATE': pd.to_datetime(orders_df['Created at']).dt.strftime('%Y-%m-%d'),
            'MONTH': pd.to_datetime(orders_df['Created at']).dt.strftime("%b'%y"),
            'BRAND': 'TILTING HEADS',
            'CUSTOMER NAME': orders_df['Billing Name'],
            'ADDRESS': orders_df.apply(lambda row: clean_repetitive_phrases(' '.join(
                dict.fromkeys([
                    str(row.get('Shipping Address1', '')).strip(),
                    str(row.get('Shipping Address2', '')).strip(),
                    str(row.get('Cleaned Zip', '')).strip(),
                    str(row.get('Shipping City', '')).strip(),
                    str(row.get('Shipping Province Name', '')).strip()
                ])
            ).replace('nan', '').replace('None', '').replace('  ', ' ').strip()), axis=1),
            'PINCODE': orders_df['Cleaned Zip'],
            'STATE': orders_df['Shipping Province Name'],
            'PHONE NUMBER': orders_df['Shipping Phone'].apply(lambda x: '{:.0f}'.format(x) if pd.notnull(x) else ''),
            'EMAIL ID': orders_df['Email'],
            'PRODUCT NAME': orders_df['Lineitem name'],
            'AMOUNT': orders_df['Total'],
            'COUNT OF ITEMS': orders_df['Lineitem quantity'],
            'PAYMENT MODE': orders_df['Financial Status'].apply(lambda x: 'Online Payment' if str(x).strip().lower() == 'paid' else 'Cash on Delivery'),
            'COURIER COMPANY': '',
            'TRACKING NUMBER': '',
            'DELIVERY STATUS/DATE': '',
            'SOLD/RETURENED': '',
            'CUSTOMER FEEDBACK': '',
            'Delivery Available': orders_df['Delivery Available']
        })

        # Save output Excel
        output_file = os.path.join(OUTPUT_FOLDER, 'cleaned_orders.xlsx')
        cleaned_df.to_excel(output_file, index=False)

        return render_template('download.html', filename='cleaned_orders.xlsx')

    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
