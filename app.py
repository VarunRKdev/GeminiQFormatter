from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename
import os
import fitz  # PyMuPDF
import requests
from dotenv import load_dotenv
import pytesseract
from PIL import Image
import re
import uuid
import subprocess
import zipfile
import tempfile
import shutil

# Load environment variables
load_dotenv()
API_KEY = os.getenv("GOOGLE_API_KEY")

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"

def convert_docx_to_pdf(docx_path, output_folder):
    try:
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", output_folder, docx_path],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE
        )
        base = os.path.splitext(os.path.basename(docx_path))[0]
        converted_pdf = os.path.join(output_folder, base + ".pdf")
        return converted_pdf if os.path.exists(converted_pdf) else None
    except Exception as e:
        print(f"[DOCX → PDF] Error: {e}")
        return None

def convert_image_to_pdf(image_path, output_folder):
    try:
        img = Image.open(image_path).convert("RGB")
        base = os.path.splitext(os.path.basename(image_path))[0]
        output_path = os.path.join(output_folder, f"{base}.pdf")
        img.save(output_path, "PDF", resolution=100.0)
        return output_path
    except Exception as e:
        print(f"[Image → PDF] Error: {e}")
        return None

def combine_images_to_pdf(image_paths, output_pdf_path):
    images = [Image.open(p).convert("RGB") for p in image_paths if p.lower().endswith(('.jpg', '.jpeg', '.png'))]
    if not images:
        return None
    first_image, *rest = images
    first_image.save(output_pdf_path, save_all=True, append_images=rest)
    return output_pdf_path

@app.route("/", methods=["GET", "POST"])
def index():
    response_text = ""
    table_data = []
    image_files = []
    filename = ""
    filetype = ""
    view_pdf = view_doc = view_image = view_zip = False
    zip_images = []

    if request.method == "POST":
        file = request.files.get("pdf")
        user_prompt = request.form.get("prompt", "").strip()

        if file and file.filename.lower().endswith(('.pdf', '.docx', '.jpg', '.jpeg', '.png', '.zip')) and user_prompt:
            filename = secure_filename(file.filename)
            filetype = filename.split('.')[-1].lower()
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            if filetype == "pdf":
                pdf_path = filepath
                view_pdf = True

            elif filetype == "docx":
                pdf_path = convert_docx_to_pdf(filepath, app.config['UPLOAD_FOLDER'])
                if not pdf_path:
                    response_text = "❌ Failed to convert Word document to PDF."
                    return render_template("index.html", response=response_text)
                view_doc = True

            elif filetype in ("jpg", "jpeg", "png"):
                pdf_path = convert_image_to_pdf(filepath, app.config['UPLOAD_FOLDER'])
                if not pdf_path:
                    response_text = "❌ Failed to convert image to PDF."
                    return render_template("index.html", response=response_text)
                view_image = True

            elif filetype == "zip":
                extracted_image_paths = extract_images_from_zip(filepath)
                if not extracted_image_paths:
                    response_text = "❌ No valid images found in ZIP."
                    return render_template("index.html", response=response_text)

                combined_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{os.path.splitext(filename)[0]}_combined.pdf")
                combine_images_to_pdf(extracted_image_paths, combined_pdf_path)
                pdf_path = combined_pdf_path
                zip_images = [os.path.basename(p) for p in extracted_image_paths]
                view_zip = True

            else:
                response_text = "❌ Unsupported file type."
                return render_template("index.html", response=response_text)

            extracted_text = extract_text_safely(pdf_path)
            images = extract_images_from_pdf(pdf_path)
            image_info = f"\n\nThis document contains {len(images)} image(s)."

            final_prompt = f"{user_prompt}{image_info}\n\nHere is the document content:\n\n{extracted_text[:15000]}"
            response_text = query_gemini(final_prompt)
            table_data = parse_response_to_table(response_text)
            image_files = images

    return render_template("index.html",
                           response=response_text,
                           table=table_data,
                           images=image_files,
                           filename=filename,
                           filetype=filetype,
                           view_pdf=view_pdf,
                           view_doc=view_doc,
                           view_image=view_image,
                           view_zip=view_zip,
                           zip_images=zip_images)

def extract_images_from_zip(zip_path):
    temp_dir = tempfile.mkdtemp()
    image_paths = []

    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        for root, _, files in os.walk(temp_dir):
            for file in files:
                if file.lower().endswith(('.jpg', '.jpeg', '.png')):
                    source_path = os.path.join(root, file)
                    unique_name = f"{uuid.uuid4()}_{file}"
                    destination_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_name)
                    try:
                        Image.open(source_path).save(destination_path)
                        image_paths.append(destination_path)
                    except Exception as e:
                        print(f"❌ Failed to process image {file}: {e}")

    finally:
        shutil.rmtree(temp_dir)

    return image_paths

def extract_text_safely(filepath):
    doc = fitz.open(filepath)
    text = ""
    for page in doc:
        t = page.get_text()
        if not t.strip():
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            t = pytesseract.image_to_string(img, config="--psm 6", lang="eng+equ")
        text += t + "\n"
    return text

def extract_images_from_pdf(filepath):
    doc = fitz.open(filepath)
    images = []
    for page_number in range(len(doc)):
        page = doc[page_number]
        for img_index, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            img_name = f"{uuid.uuid4()}.{image_ext}"
            img_path = os.path.join(app.config['UPLOAD_FOLDER'], img_name)
            with open(img_path, "wb") as f:
                f.write(image_bytes)
            images.append({"page": page_number + 1, "path": img_name})
    return images

def query_gemini(prompt):
    headers = {"Content-Type": "application/json"}
    params = {"key": API_KEY}
    data = {
        "contents": [{
            "parts": [{"text": prompt}]
        }]
    }
    try:
        res = requests.post(GEMINI_API_URL, headers=headers, params=params, json=data)
        res.raise_for_status()
        return res.json()["candidates"][0]["content"]["parts"][0]["text"]
    except Exception as e:
        return f"❌ Error: {e}"

def parse_response_to_table(response_text):
    table = []
    pattern = r"Question No\.:\s*(.+?)\nQuestion:\s*(.+?)\nMarks:\s*(.+?)(?:\n|$)"
    matches = re.findall(pattern, response_text, re.DOTALL)

    for match in matches:
        q_no = match[0].strip()
        question = match[1].strip()
        marks = match[2].strip()
        has_diagram = "diagram" in question.lower() or "(diagram-based)" in question.lower()
        table.append({
            "Question No.": q_no,
            "Question": question,
            "Marks": marks,
            "Diagram": "Yes" if has_diagram else "No"
        })
    return table

@app.route("/uploads/<filename>")
def serve_upload(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route("/view-pdf/<filename>")
def view_pdf(filename):
    return render_template("viewer.html", filename=filename)

@app.route("/view-doc/<filename>")
def view_doc(filename):
    doc_basename = os.path.splitext(filename)[0]
    pdf_name = doc_basename + ".pdf"
    return render_template("viewer_doc.html", filename=pdf_name)

@app.route("/view-image/<filename>")
def view_image(filename):
    return render_template("viewer_image.html", filename=filename)

@app.route("/view-zip/<zipname>")
def view_zip(zipname):
    zip_folder = app.config['UPLOAD_FOLDER']
    base = zipname.replace('.zip', '')
    image_paths = [f for f in os.listdir(zip_folder)
                   if f.lower().endswith(('.jpg', '.jpeg', '.png')) and base in f]
    image_paths.sort()
    return render_template("viewer_zip.html", images=image_paths)

if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)
