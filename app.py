from flask import Flask, render_template, request, send_file
import pdfplumber
import pandas as pd
import re
import os
from io import BytesIO
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def clean_name(raw_name):
    name = re.sub(r'[^A-Z ]', ' ', raw_name.upper())
    words = name.split()

    subjects = {
        "GEOGRAPHY", "HISTORY", "POLITICAL", "SCIENCE",
        "ECONOMICS", "SOCIOLOGY", "PHILOSOPHY", "HINDI",
        "ENGLISH", "PSYCHOLOGY", "HOME", "SCI", "ARTS"
    }

    clean_words = []

    for w in words:
        if w in subjects:
            continue
        if len(w) == 1:
            continue
        clean_words.append(w)

    name = " ".join(clean_words[:3])
    return name.strip()


def extract_data(pdf_path):
    students = []
    course_name = ""
    college_name = ""
    center_name = ""

    with pdfplumber.open(pdf_path) as pdf:

        for page_no, page in enumerate(pdf.pages, start=1):

            text = page.extract_text()
            if not text:
                continue

            if not course_name:
                match = re.search(r"BACHELOR OF .*", text)
                if match:
                    course_name = match.group(0).strip()

            if not college_name:
                match = re.search(r"College Name\s*:\s*(.*)", text)
                if match:
                    college_name = match.group(1).strip()

            if not center_name:
                match = re.search(r"Exam Centre Name\s*:\s*(.*)", text)
                if match:
                    center_name = match.group(1).strip()

            lines = text.split("\n")

            for line in lines:
                line = line.strip()

                match = re.match(
                    r'^\d+\s+(\d{10,})\s+(.*?)\s+(PASS|FAIL|ABSENT)\s+(\d{10,})$',
                    line
                )

                if match:
                    reg_no = match.group(1)
                    raw_name = match.group(2)
                    roll_no = match.group(4)

                    name = clean_name(raw_name)

                    students.append([
                        roll_no,
                        reg_no,
                        name,
                        course_name,
                        college_name,
                        center_name,
                        page_no
                    ])

    columns = [
        "Roll No",
        "Registration Number",
        "Name",
        "Course Name",
        "College Name",
        "Center Name",
        "Page Number"
    ]

    df = pd.DataFrame(students, columns=columns)
    return df


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    file = request.files["pdf"]

    filename = secure_filename(file.filename)
    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(pdf_path)

    df = extract_data(pdf_path)

    # 🔥 MEMORY BASED EXCEL (NO CORRUPTION)
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="result.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == "__main__":
    app.run(debug=True)
