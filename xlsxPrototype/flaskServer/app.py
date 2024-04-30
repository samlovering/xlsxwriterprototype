from flask import Flask, Blueprint, render_template, send_from_directory
from reports import generateReport

app = Flask(__name__)


@app.route('/generateReport', methods=['GET'])
def generate_report():
    fileName = generateReport.generate_spreadsheet()
    return fileName


@app.route('/download/<path:filename>', methods=['GET', 'POST'])
def download_file(filename):
    return send_from_directory(directory='xlsx', path=filename, as_attachment=True)


@app.route('/')
def index():
    return render_template('index.html')


if __name__ == '__main__':
    app.run(port=5000, debug=True)
