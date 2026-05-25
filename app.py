import os
import subprocess
import base64
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path

app = Flask(__name__, static_folder='.')
CORS(app) 

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'temp')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 雲端 Linux 環境的 LibreOffice 預設路徑
LIBREOFFICE_PATH = os.environ.get('LIBREOFFICE_PATH', '/usr/bin/soffice')

# 讓同事輸入網址時，直接看得到你的 HTML 網頁
@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/convert-excel', methods=['POST'])
def convert_excel():
    if 'file' not in request.files:
        return jsonify({'error': '沒有上傳檔案'}), 400
        
    file = request.files['file']
    filename = secure_filename(file.filename)
    excel_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(excel_path)
    
    try:
        subprocess.run([
            LIBREOFFICE_PATH, 
            '--headless', 
            '--convert-to', 'pdf', 
            '--outdir', UPLOAD_FOLDER, 
            excel_path
        ], check=True)
        
        pdf_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.pdf')
        
        # 雲端 Linux 環境不需要指定 poppler_path，系統會自動辨識
        images = convert_from_path(pdf_path, dpi=300)
        
        img_path = os.path.join(UPLOAD_FOLDER, 'page1.jpg')
        images[0].save(img_path, 'JPEG')
        
        with open(img_path, "rb") as img_file:
            encoded_string = base64.b64encode(img_file.read()).decode('utf-8')
            
        os.remove(excel_path)
        os.remove(pdf_path)
        os.remove(img_path)
        
        return jsonify({'image': 'data:image/jpeg;base64,' + encoded_string})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(port=5000)