from flask import Flask, request, send_file, render_template_string
import os
from werkzeug.utils import secure_filename
from PIL import Image
import subprocess

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def convert_with_libreoffice(input_path):
    libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"  # Update if needed
    output_dir = os.path.dirname(input_path)
    subprocess.run([
        libreoffice_path, '--headless', '--convert-to', 'pdf', '--outdir',
        output_dir, input_path
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def convert_images_to_single_pdf(image_paths, output_path):
    image_list = []
    for path in image_paths:
        img = Image.open(path).convert('RGB')
        image_list.append(img)
    if image_list:
        first_image = image_list[0]
        first_image.save(output_path, save_all=True, append_images=image_list[1:], format='PDF')
    return output_path

@app.route('/')
def index():
    with open('index.html') as f:
        return render_template_string(f.read())

@app.route('/convert', methods=['POST'])
def convert():
    files = request.files.getlist('file')
    if not files:
        return "No files uploaded", 400

    pdf_output = os.path.join(UPLOAD_FOLDER, "output.pdf")
    file_exts = [os.path.splitext(f.filename)[1].lower() for f in files]

    # Case 1: One document (Word/Excel/PPT/TXT)
    if len(files) == 1 and file_exts[0] in ['.doc','.csv', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt']:
        file = files[0]
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        convert_with_libreoffice(filepath)
        output_file = os.path.splitext(filepath)[0] + ".pdf"
        return send_file(output_file, as_attachment=True)

    # Case 2: Multiple images
    elif all(ext in ['.jpg', '.jpeg', '.png', '.gif'] for ext in file_exts):
        image_paths = []
        for file in files:
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)
            image_paths.append(filepath)
        convert_images_to_single_pdf(image_paths, pdf_output)
        return send_file(pdf_output, as_attachment=True)

    else:
        return "Unsupported file type(s)", 400

if __name__ == '__main__':
    app.run(debug=True)
