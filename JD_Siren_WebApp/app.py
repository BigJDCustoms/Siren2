from flask import Flask, request, render_template, send_file
import os, zipfile, shutil
from werkzeug.utils import secure_filename
from siren_utils import process_zip

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['RESULT_FOLDER'] = 'results'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['RESULT_FOLDER'], exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        uploaded = request.files.get('file')
        if uploaded:
            filename = secure_filename(uploaded.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            uploaded.save(filepath)

            result_txt, result_xlsx = process_zip(filepath, app.config['RESULT_FOLDER'])

            return render_template('index.html', txt=result_txt, xlsx=result_xlsx)

    return render_template('index.html')

@app.route('/download/<path:filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['RESULT_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
