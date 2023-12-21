from flask import Flask, render_template, request
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Print form data to debug
    print(request.form)

    # Get form data
    name = request.form['name']
    age = request.form['age']
    sex = request.form['sex']
    contact = request.form.get('contact', '')  # Use get to handle missing key
    R_Power = request.form['R_Power']
    L_Power = request.form['L_Power']
    Near_Add = request.form['Near_Add']
    Lens_Type = request.form['Lens_Type']
    date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    problem = request.form['problem']
    amount = request.form['amount']
    advance = request.form['advance']

    # Specify the desired Excel file path
    excel_path = r'C:\Users\vicky\Downloads\PatientManagementApp\Patient_Details.xlsx'

    try:
        df = pd.read_excel(excel_path)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Patient Name', 'Age', 'Sex', 'Contact no', 'Date', 'R Power', 'L Power', 'Lens Type', 'Amount', 'Advance', 'Problem'])

    # Append new data to the DataFrame
    new_data = {
        'Patient Name': [name],
        'Age': [age],
        'Sex': [sex],
        'Contact no': [contact],
        'Date': [date],
        'R Power': [R_Power],
        'L Power': [L_Power],
        'Near Add': [Near_Add],
        'Lens Type': [Lens_Type],
        'Amount': [amount],
        'Advance': [advance],
        'Problem': [problem]
    }
    df = pd.concat([df, pd.DataFrame(new_data)], ignore_index=True)

    # Save the updated DataFrame to Excel
    df.to_excel(excel_path, index=False)

    return render_template('index.html', message='Data saved successfully!')

if __name__ == '__main__':
    app.run(debug=True)
