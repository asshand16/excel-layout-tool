import os
import subprocess
import base64
import platform
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
from PIL import Image, ImageChops

app = Flask(__name__, static_folder='.')
CORS(app) 

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'temp')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 智慧偵測環境 (自動切換 Windows 與 雲端 Linux)
if platform.system() == 'Windows':
    LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
    # ⚠️ 請確保這裡是你在本機測試成功的 poppler 路徑
    POPPLER_PATH = r"C:\poppler\Library\bin" 
else:
    LIBREOFFICE_PATH = '/usr/bin/soffice'
    POPPLER_PATH = None 

# 🌟 斬草除根版：一律強制切除 2.1 公分的頁首與頁尾盲區，消滅隱形雜訊 🌟
def clean_and_trim(img):
    width, height = img.size
    
    # 在 300 DPI 下，250 pixels 大約是 2.1 公分。
    # 這足以切掉 Excel 預設的頁首 (包含 _x000a_ 雜訊) 與頁尾 (包含頁碼)
    # 且不會傷害到正常設定在邊界內的 Logo 或表格內容。
    top_cut = 135
    bottom_cut = 250
    
    # 1. 暴力切除上下邊界盲區
    img_cropped = img.crop((0, top_cut, width, height - bottom_cut))
    
    # 2. 用 ImageChops 精準抓取實際內容邊界，切除多餘留白
    bg = Image.new(img_cropped.mode, img_cropped.size, (255, 255, 255))
    diff = ImageChops.difference(img_cropped, bg)
    bbox = diff.getbbox()
    
    if bbox:
        # 鎖死左右寬度不變 (保證表格 100% 對齊)
        # 僅裁切上下，並保留 2px 緩衝讓線條完美接合
        upper = max(0, bbox[1] - 2)
        lower = min(img_cropped.size[1], bbox[3] + 2)
        return img_cropped.crop((0, upper, width, lower))
        
    return img_cropped

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
        # 1. Excel 轉 PDF
        subprocess.run([
            LIBREOFFICE_PATH, 
            '--headless', 
            '--convert-to', 'pdf', 
            '--outdir', UPLOAD_FOLDER, 
            excel_path
        ], check=True)
        
        pdf_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.pdf')
        
        # 2. PDF 轉圖片
        if POPPLER_PATH:
            images = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
        else:
            images = convert_from_path(pdf_path, dpi=300)
        
        img_path = os.path.join(UPLOAD_FOLDER, 'result.jpg')
        
        # 3. 處理每一頁，全部送進 clean_and_trim 消除雜訊
        trimmed_images = [clean_and_trim(img) for img in images]

        # 4. 動態無縫拼接
        if len(trimmed_images) == 1:
            trimmed_images[0].save(img_path, 'JPEG', quality=95)
        else:
            # 以第一張圖的寬度為基準，高度全部加總
            max_width = trimmed_images[0].size[0]
            total_height = sum(img.size[1] for img in trimmed_images)
            
            new_im = Image.new('RGB', (max_width, total_height), (255, 255, 255))
            y_offset = 0
            for im in trimmed_images:
                # 靠左對齊，保證表格直線完美貼合
                new_im.paste(im, (0, y_offset))
                y_offset += im.size[1]
                
            new_im.save(img_path, 'JPEG', quality=95)
        
        # 5. 轉 Base64 回傳前端
        with open(img_path, "rb") as img_file:
            encoded_string = base64.b64encode(img_file.read()).decode('utf-8')
            
        os.remove(excel_path)
        os.remove(pdf_path)
        os.remove(img_path)
        
        return jsonify({'image': 'data:image/jpeg;base64,' + encoded_string})

    except Exception as e:
        error_msg = str(e)
        print(f"\n❌ 轉檔錯誤: {error_msg}\n")
        return jsonify({'error': error_msg}), 500

if __name__ == '__main__':
    app.run(port=8080, debug=True)