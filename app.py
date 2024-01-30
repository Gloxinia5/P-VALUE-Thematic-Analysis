import ai
import process
from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file
import pandas as pd
from werkzeug.utils import secure_filename
import os
import json
from io import BytesIO
import test

app = Flask(__name__)

excel_data = {} #Deftonest

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # Check if the post request has the file part
    if 'file' not in request.files:
        return redirect(request.url)

    file = request.files['file']

    # If the user does not select a file, submit an empty part without filename
    if file.filename == '':
        return redirect(request.url)

    # If the file is allowed and the user submitted a valid file
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # Read Excel file into a dictionary
        df = pd.read_excel(filepath, header=None)
        excel_dict = {}

        for index, row in df.iterrows():
            item_name = str(row[0])
            item_data = row[1:].to_dict()
            excel_dict[item_name] = item_data
        excel_data = excel_dict
    data = process.to_array(excel_data)
    test1 = []
    for I in data:
        x = process.json_string_to_dict(ai.prompt(I)["choices"][0]["message"]["content"])
        while x is None:
            x = process.json_string_to_dict(ai.prompt(I)["choices"][0]["message"]["content"])
        test1.append(x)
    return send_file(
        BytesIO(process.dict_to_excel(process.convert_array_to_dict(test1))),
        as_attachment=True,
        download_name='output.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

