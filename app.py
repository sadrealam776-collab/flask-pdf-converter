import os
from flask import Flask, render_template, request, send_file, jsonify, after_this_request
from werkzeug.utils import secure_filename
from pdf2docx import Converter

app = Flask(__name__)

# CONFIGURATION
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
ALLOWED_EXTENSIONS = {'pdf'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
        
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(pdf_path)
        
        docx_filename = filename.rsplit('.', 1)[0] + '.docx'
        docx_path = os.path.join(app.config['CONVERTED_FOLDER'], docx_filename)
        
        try:
            # --- THE "SPACE FIX" IS HERE ---
            cv = Converter(pdf_path)
            
            # We use a very low x_tolerance (0.5) to stop words from sticking together.
            cv.convert(docx_path, start=0, end=None, **{
                'x_tolerance': 0.5,      # LOWER this number to separate words (Try 0.5, then 0.2)
                'y_tolerance': 2.0,      # Keeps lines separate
                'keep_al': True,         # Enforces strict layout
                'ocr': 0                 # No OCR (faster)
            })
            cv.close()
            
            return jsonify({
                "status": "success", 
                "download_url": f"/download/{docx_filename}"
            })

        except Exception as e:
            return jsonify({"error": str(e)}), 500

    return jsonify({"error": "Invalid file type"}), 400

@app.route('/download/<filename>')
def download_file(filename):
    filepath = os.path.join(app.config['CONVERTED_FOLDER'], filename)
    
    @after_this_request
    def remove_file(response):
        try:
            # Clean up files
            pdf_name = filename.rsplit('.', 1)[0] + '.pdf'
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_name)
            if os.path.exists(pdf_path): os.remove(pdf_path)
        except Exception as e:
            app.logger.error("Error removing file", e)
        return response
        
    return send_file(filepath, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)