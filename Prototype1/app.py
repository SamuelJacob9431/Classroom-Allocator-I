from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from openpyxl import Workbook

app = Flask(__name__)


@app.route('/')
def index():
    return render_template("index.html")


@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['excel_file']

    if file.filename == '':
        return "No file uploaded."

    max_capacity = int(request.form["capacity"])

    df = pd.read_excel(file)

    # Validate required columns
    required = {"Name of Student", "Adm No.", "Class"}
    if not required.issubset(df.columns):
        return "❌ Excel must contain: Name of Student, Adm No., Class columns."

    # Group by Class
    class_groups = df.groupby("Class")

    output_files = []
    room_number = 1

    for class_name, group in class_groups:
        group = group.sort_values("Adm No.")

        if len(group) > max_capacity:
            return f"❌ Class {class_name} has {len(group)} students, exceeding capacity {max_capacity}. Cannot allot."

        # Create Excel file for the class (room)
        output_filename = f"Room_{room_number}_{class_name}.xlsx"
        output_path = os.path.join(os.getcwd(), output_filename)

        group.to_excel(output_path, index=False)
        output_files.append(output_filename)

        room_number += 1

    return render_template("result.html", files=output_files)


@app.route('/download/<filename>')
def download(filename):
    file_path = os.path.join(os.getcwd(), filename)
    return send_file(file_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
