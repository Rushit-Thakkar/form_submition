from flask import Flask, request, render_template
import pandas as pd
import os

app = Flask(__name__)

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    data = {
        'Name': request.form['name'],
        'Email': request.form['email'],
        'Message': request.form['message']
    }
    
    df = pd.DataFrame([data])
    
    file_path = 'form_data.xlsx'
    
    if os.path.exists(file_path):
        existing_df = pd.read_excel(file_path)
        df = pd.concat([existing_df, df], ignore_index=True)
    
    df.to_excel(file_path, index=False)
    
    return 'Form submitted successfully!'

if __name__ == '__main__':
    app.run(debug=True)