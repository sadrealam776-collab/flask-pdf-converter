import os
from flask import Flask, render_template, request, send_file, jsonify, after_this_request
from werkzeug.utils import secure_filename
from pdf2docx import Converter

app = Flask(__name__)

# --- CONFIGURATION ---
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
ALLOWED_EXTENSIONS = {'pdf'}

# Create folders if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    try:
        # 1. Try to load it normally (Standard Flask way)
        return render_template('index.html')
        
    except UnicodeDecodeError:
        # 2. If that fails (because of Windows saving), read it manually
        # This forces Python to read the "Windows-style" text without crashing
        with open('templates/index.html', 'r', encoding='cp1252', errors='ignore') as f:
            return f.read()
    except Exception as e:
        return f"Critical Error loading file: {e}"

@app.route('/convert', methods=['POST'])
def convert_file():
    # 1. Check if file was uploaded
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
        
    if file and allowed_file(file.filename):
        try:
            # 2. Save the PDF
            filename = secure_filename(file.filename)
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(pdf_path)
            
            # 3. Define Output Path
            docx_filename = filename.rsplit('.', 1)[0] + '.docx'
            docx_path = os.path.join(app.config['CONVERTED_FOLDER'], docx_filename)
            
            # 4. RUN CONVERSION (With Spacing Fix)
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None, **{
                'x_tolerance': 0.25,  # Fixes "merged words"
                'keep_al': True,      # Strict layout
                'ocr': 0              # Disable OCR (prevents crashes if Tesseract is missing)
            })
            cv.close()
            
            # 5. Send Success Response
            return jsonify({
                "status": "success", 
                "download_url": f"/download/{docx_filename}"
            })

        except Exception as e:
            # This prints the REAL error to your terminal so you can see it
            print(f"❌ SERVER ERROR: {e}") 
            return jsonify({"error": str(e)}), 500

    return jsonify({"error": "Invalid file type"}), 400

@app.route('/download/<filename>')
def download_file(filename):
    filepath = os.path.join(app.config['CONVERTED_FOLDER'], filename)
    
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    else:
        return "File not found", 404

if __name__ == '__main__':
    # For local testing
    app.run(debug=True, port=5000)

