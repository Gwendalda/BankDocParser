
from flask import Flask, render_template, request, redirect, flash, send_file
import os
import zipfile
from docparser import sendFilesToDocParser

app = Flask(__name__)


def get_filenames():
    files = []
    for filename in os.listdir(app.config['UPLOAD_FOLDER']):
        if os.path.isfile(os.path.join(app.config['UPLOAD_FOLDER'], filename)):
            files.append(filename)
    return files


def get_only_excel_files():
    files = []
    for filename in os.listdir(app.config['UPLOAD_FOLDER']):
        if os.path.isfile(os.path.join(app.config['UPLOAD_FOLDER'], filename)) and ".xlsx" in filename:
            files.append(filename)
    return files


def allowed_file(filename):
    if not ".pdf" in filename:
        flash('Only PDF files are allowed')
        return False
    return True


@app.route('/')
def index():
    return render_template('index.html', files=get_only_excel_files())

# upload one or more files


@app.route('/upload', methods=['POST'])
def upload():
    if 'files[]' not in request.files:
        flash('No file part')
        return redirect('/')

    files = request.files.getlist('files[]')

    for file in files:
        if file and allowed_file(file.filename):
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], file.filename))
            sendFilesToDocParser(
                [os.path.join(app.config['UPLOAD_FOLDER'], file.filename)])
            # remove all the files that end in .json and .pdf
            for filename in os.listdir(app.config['UPLOAD_FOLDER']):
                if os.path.isfile(os.path.join(app.config['UPLOAD_FOLDER'], filename)) and (".json" in filename or ".pdf" in filename):
                    os.remove(os.path.join(
                        app.config['UPLOAD_FOLDER'], filename))

    flash('File(s) successfully uploaded')
    return redirect('/')

# download one or more files


@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename),
                     mimetype='application/pdf',
                     as_attachment=True)


if __name__ == '__main__':
    app.config['static_url_path'] = '/views'
    app.config['UPLOAD_FOLDER'] = 'ressources/uploads'
    app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
    app.secret_key = "1234567890"
    app.run(debug=True, host='10.162.0.2', port=5000)
