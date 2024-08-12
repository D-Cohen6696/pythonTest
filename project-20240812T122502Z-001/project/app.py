from flask import Flask, request, jsonify, send_file
import pandas as pd
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from io import BytesIO
import matplotlib.pyplot as plt
from reportlab.lib.units import inch

app = Flask(__name__)

# Directory to store uploaded files
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    if file and file.filename.endswith('.xlsx'):
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        excel_file = pd.ExcelFile(filepath)
        response = {
            "filename": file.filename,
            "num_sheets": len(excel_file.sheet_names)
        }
        return jsonify(response)
    else:
        return jsonify({"error": "Invalid file format"}), 400


@app.route('/process', methods=['POST'])
def process_data():
    data = request.json
    filepath = data.get('path')
    sheets_info = data.get('sheets')

    if not filepath or not sheets_info:
        return jsonify({"error": "Invalid data"}), 400

    excel_file = pd.ExcelFile(filepath)
    report = {}

    for sheet_info in sheets_info:
        sheet_name = sheet_info['sheet']
        operation = sheet_info['operation']
        columns = sheet_info['columns']

        if sheet_name not in excel_file.sheet_names:
            return jsonify({"error": f"Sheet {sheet_name} not found"}), 400

        df = pd.read_excel(filepath, sheet_name=sheet_name)

        if operation == 'sum':
            result = df[columns].sum().to_dict()
        elif operation == 'average':
            result = df[columns].mean().to_dict()
        else:
            return jsonify({"error": "Invalid operation"}), 400

        report[sheet_name] = result

    return jsonify(report)


@app.route('/generate_pdf', methods=['POST'])
def generate_pdf():
    report = request.json
    if not report:
        return jsonify({"error": "Invalid report data"}), 400

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)

    c.drawString(100, 750, "Report")
    y_position = 730

    for sheet_name, data in report.items():
        c.drawString(100, y_position, f"Sheet: {sheet_name}")
        y_position -= 20
        for col, value in data.items():
            c.drawString(120, y_position, f"{col}: {value}")
            y_position -= 20
        y_position -= 20

    c.save()
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="report.pdf", mimetype="application/pdf")


@app.route('/generate_graph', methods=['POST'])
def generate_graph():
    data = request.json
    if not data:
        return jsonify({"error": "Invalid data"}), 400

    sheets = data.get('sheets')
    if not sheets:
        return jsonify({"error": "No sheets data"}), 400

    sheet_names = list(sheets.keys())
    values = [sum(data.get(sheet, {}).values()) for sheet in sheet_names]

    plt.figure(figsize=(10, 6))
    plt.bar(sheet_names, values)
    plt.xlabel('Sheets')
    plt.ylabel('Sum')
    plt.title('Sum of Each Sheet')
    plt.grid(True)

    graph_buffer = BytesIO()
    plt.savefig(graph_buffer, format='png')
    plt.close()
    graph_buffer.seek(0)

    return send_file(graph_buffer, as_attachment=True, download_name="graph.png", mimetype="image/png")


@app.route('/generate_detailed_pdf', methods=['POST'])
def generate_detailed_pdf():
    report = request.json
    if not report:
        return jsonify({"error": "Invalid report data"}), 400

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)

    # Adding report content
    c.drawString(100, 750, "Detailed Report")
    y_position = 730

    for sheet_name, data in report.items():
        c.drawString(100, y_position, f"Sheet: {sheet_name}")
        y_position -= 20
        for col, value in data.items():
            c.drawString(120, y_position, f"{col}: {value}")
            y_position -= 20
        y_position -= 20

    # Add a graph
    graph_buffer = BytesIO()
    plt.figure(figsize=(10, 6))
    sheet_names = list(report.keys())
    values = [sum(report.get(sheet, {}).values()) for sheet in sheet_names]
    plt.bar(sheet_names, values)
    plt.xlabel('Sheets')
    plt.ylabel('Sum')
    plt.title('Sum of Each Sheet')
    plt.grid(True)
    plt.savefig(graph_buffer, format='png')
    plt.close()
    graph_buffer.seek(0)

    c.drawImage(graph_buffer, 100, y_position - 200, width=6 * inch, height=4 * inch)
    c.save()

    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="detailed_report.pdf", mimetype="application/pdf")


if __name__ == '__main__':
    app.run(debug=True)
