from flask import Flask, render_template, request, send_file
import pdfplumber
import pandas as pd
import re
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def clean_name(raw_name):
    # only keep letters + space
    name = re.sub(r'[^A-Z ]', ' ', raw_name.upper())

    words = name.split()

    # BIG subject list
    subjects = {
        "GEOGRAPHY", "HISTORY", "POLITICAL", "SCIENCE",
        "ECONOMICS", "SOCIOLOGY", "PHILOSOPHY", "HINDI",
        "ENGLISH", "PSYCHOLOGY", "HOME", "SCI", "ARTS"
    }

    clean_words = []

    for w in words:
        # skip subject words
        if w in subjects:
            continue

        # skip 1-letter garbage (A, H etc)
        if len(w) == 1:
            continue

        clean_words.append(w)

    # take only first 3 words (real name)
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

            # ---------------- HEADER ----------------
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

            # ---------------- STUDENT DATA ----------------
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

                    # 🔥 FIXED NAME CLEANING
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

    pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(pdf_path)

    df = extract_data(pdf_path)

    base_name = os.path.splitext(file.filename)[0]
    output_filename = f"{base_name}_result.xlsx"
    output_path = os.path.join(UPLOAD_FOLDER, output_filename)

    df.to_excel(output_path, index=False)

    return send_file(
        output_path,
        as_attachment=True,
        download_name=output_filename
    )


if __name__ == "__main__":
    app.run(debug=True)