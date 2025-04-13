from flask import Flask, render_template, request, send_file
from pdf2docx import Converter
from pdf2image import convert_from_path
import zipfile
import pdfplumber
import pypandoc
import os
import uuid
import subprocess
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert_file():
    file = request.files["file"]
    conversion_type = request.form["conversion_type"]

    if not file:
        return "No file uploaded", 400

    filename = secure_filename(file.filename)
    input_path = os.path.join(app.config["UPLOAD_FOLDER"], f"{uuid.uuid4()}_{filename}")
    file.save(input_path)

    base_name = os.path.splitext(os.path.basename(filename))[0]
    output_ext = conversion_type.split("-")[-1]
    output_filename = f"{base_name}_converted.{output_ext}"
    output_path = os.path.join(app.config["UPLOAD_FOLDER"], output_filename)

    try:
        # Word conversions
        if conversion_type == "word-pdf":
            subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    UPLOAD_FOLDER,
                    input_path,
                ]
            )
            output_path = os.path.join(
                UPLOAD_FOLDER,
                f"{os.path.splitext(os.path.basename(input_path))[0]}.pdf",
            )

        elif conversion_type in [
            "word-txt",
            "word-html",
            "word-odt",
            "word-rtf",
            "word-epub",
        ]:
            pandoc_format = "plain" if output_ext == "txt" else output_ext
            pypandoc.convert_file(input_path, pandoc_format, outputfile=output_path)

        # Excel conversions
        elif conversion_type == "excel-pdf":
            subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    UPLOAD_FOLDER,
                    input_path,
                ]
            )
            output_path = os.path.join(
                UPLOAD_FOLDER,
                f"{os.path.splitext(os.path.basename(input_path))[0]}.pdf",
            )

        elif conversion_type == "excel-csv":
            df = pd.read_excel(input_path)
            df.to_csv(output_path, index=False)

        elif conversion_type == "excel-txt":
            df = pd.read_excel(input_path)
            df.to_csv(output_path, sep="\t", index=False)

        elif conversion_type == "excel-ods":
            subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to",
                    "ods",
                    "--outdir",
                    UPLOAD_FOLDER,
                    input_path,
                ]
            )
            output_path = os.path.join(
                UPLOAD_FOLDER,
                f"{os.path.splitext(os.path.basename(input_path))[0]}.ods",
            )

        # PowerPoint conversions
        elif conversion_type == "ppt-pdf":
            subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    UPLOAD_FOLDER,
                    input_path,
                ]
            )
            output_path = os.path.join(
                UPLOAD_FOLDER,
                f"{os.path.splitext(os.path.basename(input_path))[0]}.pdf",
            )

        elif conversion_type == "ppt-odp":
            subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to",
                    "odp",
                    "--outdir",
                    UPLOAD_FOLDER,
                    input_path,
                ]
            )
            output_path = os.path.join(
                UPLOAD_FOLDER,
                f"{os.path.splitext(os.path.basename(input_path))[0]}.odp",
            )

        elif conversion_type == "ppt-image":
            temp_pdf = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}_slides.pdf")
            subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    UPLOAD_FOLDER,
                    input_path,
                ]
            )
            pdf_name = os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
            pdf_path = os.path.join(UPLOAD_FOLDER, pdf_name)

            slides = convert_from_path(pdf_path)
            img_paths = []
            for i, slide in enumerate(slides):
                img_file = os.path.join(UPLOAD_FOLDER, f"{base_name}_slide_{i + 1}.png")
                slide.save(img_file, "PNG")
                img_paths.append(img_file)

            output_path = img_paths[0]  # Sending only first image for now

        # PDF to DOCX
        elif conversion_type == "pdf-docx":
            output_path = os.path.join(UPLOAD_FOLDER, f"{base_name}.docx")
            cv = Converter(input_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()

        # PDF to Text
        elif conversion_type == "pdf-txt":
            with pdfplumber.open(input_path) as pdf:
                text = ""
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(text)

        # PDF to Image (zipped)
        elif conversion_type == "pdf-img":
            images = convert_from_path(input_path)
            img_paths = []
            for i, image in enumerate(images):
                img_path = os.path.join(UPLOAD_FOLDER, f"{base_name}_page_{i + 1}.png")
                image.save(img_path, "PNG")
                img_paths.append(img_path)

            zip_filename = f"{base_name}_images.zip"
            output_path = os.path.join(UPLOAD_FOLDER, zip_filename)

            with zipfile.ZipFile(output_path, "w") as zipf:
                for img in img_paths:
                    zipf.write(img, os.path.basename(img))

            for img in img_paths:
                os.remove(img)

        else:
            return "Unsupported conversion type", 400

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"Error during conversion: {str(e)}", 500

    finally:
        if os.path.exists(input_path):
            os.remove(input_path)


if __name__ == "__main__":
    app.run(debug=True)
