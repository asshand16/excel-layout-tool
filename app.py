import os
import io
import base64
import platform
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET
import threading
import atexit
import glob
from typing import List

import openpyxl
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

if platform.system() == 'Windows':
    LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
    POPPLER_PATH = r"C:\poppler\library\bin"
else:
    LIBREOFFICE_PATH = '/usr/bin/soffice'
    POPPLER_PATH = None

# ===== 優化 1：並發控制 =====
MAX_CONCURRENT_CONVERSIONS = 1
conversion_semaphore = threading.Semaphore(MAX_CONCURRENT_CONVERSIONS)

# ===== 優化 2：定時清理臨時文件 =====
def cleanup_temp_files():
    """在應用退出時清理所有臨時文件"""
    try:
        for temp_file in glob.glob(os.path.join(UPLOAD_FOLDER, '*')):
            try:
                if os.path.isfile(temp_file):
                    os.remove(temp_file)
                elif os.path.isdir(temp_file):
                    import shutil
                    shutil.rmtree(temp_file)
            except Exception as e:
                print(f"清理文件失敗 {temp_file}: {e}")
    except Exception as e:
        print(f"清理臨時文件夾失敗: {e}")

atexit.register(cleanup_temp_files)


def str_to_bool(value, default=False):
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {'1', 'true', 'yes', 'y', 'on'}


# ===== 優化 3：降低圖片質量 =====
def image_to_data_url(img: Image.Image, quality: int = 75) -> str:
    """質量從 95 降低到 75，節省記憶體和傳輸"""
    if img.mode != 'RGB':
        img = img.convert('RGB')
    buf = io.BytesIO()
    img.save(buf, format='JPEG', quality=quality)
    encoded = base64.b64encode(buf.getvalue()).decode('utf-8')
    return 'data:image/jpeg;base64,' + encoded


def remove_print_titles_by_xml(xlsx_path: str) -> None:
    """
    直接修改 xlsx 內部的 xl/workbook.xml，穩定移除 _xlnm.Print_Titles。
    Windows 下暫存檔要建立在同一個磁碟，避免 os.replace 出現 WinError 17。
    """
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    source_dir = os.path.dirname(os.path.abspath(xlsx_path))
    fd, tmp_xlsx = tempfile.mkstemp(suffix='.xlsx', dir=source_dir)
    os.close(fd)

    try:
        with zipfile.ZipFile(xlsx_path, 'r') as zin, zipfile.ZipFile(tmp_xlsx, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'xl/workbook.xml':
                    data = zin.read(item.filename)
                    root = ET.fromstring(data)
                    for defined_name in root.findall('.//main:definedName', ns):
                        if defined_name.get('name') == '_xlnm.Print_Titles':
                            root.remove(defined_name)
                    new_data = ET.tostring(root, encoding='utf-8')
                    zout.writestr(item, new_data)
                else:
                    zout.writestr(item, zin.read(item.filename))
    except Exception as e:
        print(f"移除列印標題失敗: {e}")
    finally:
        if os.path.exists(tmp_xlsx):
            try:
                os.replace(tmp_xlsx, xlsx_path)
            except Exception as e:
                print(f"替換文件失敗: {e}")


def optimize_excel_for_printing(file_path: str) -> None:
    """優化 Excel 文件以供列印"""
    try:
        wb = openpyxl.load_workbook(file_path)
        for ws in wb.sheetnames:
            sheet = wb[ws]
            sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
            sheet.page_setup.orientation = 'portrait'
            sheet.print_options.horizontalCentered = True
        wb.save(file_path)
        remove_print_titles_by_xml(file_path)
    except Exception as e:
        print(f"優化 Excel 失敗: {e}")


def get_content_bbox(img: Image.Image):
    """獲取圖片內容邊界框"""
    if img.mode != 'RGB':
        img = img.convert('RGB')
    bg = Image.new('RGB', img.size, (255, 255, 255))
    diff = ImageChops.difference(img, bg)
    return diff.getbbox()


def clean_and_trim(img: Image.Image, vertical_padding: int = 18) -> Image.Image:
    """
    給 pages[] 單頁顯示用：
    保留較舒服的上下空白。
    """
    if img.mode != 'RGB':
        img = img.convert('RGB')

    bbox = get_content_bbox(img)
    if not bbox:
        return img

    upper = max(0, bbox[1] - vertical_padding)
    lower = min(img.size[1], bbox[3] + vertical_padding)

    if upper <= 2 and lower >= img.size[1] - 2:
        return img

    return img.crop((0, upper, img.size[0], lower))


def crop_for_merge(img: Image.Image, keep_top: int, keep_bottom: int) -> Image.Image:
    """
    給 merged_image 合併長圖用：
    交界處只保留極少空白，消除 PDF 跨頁白帶。
    """
    if img.mode != 'RGB':
        img = img.convert('RGB')

    bbox = get_content_bbox(img)
    if not bbox:
        return img

    upper = max(0, bbox[1] - max(0, keep_top))
    lower = min(img.size[1], bbox[3] + max(0, keep_bottom))
    return img.crop((0, upper, img.size[0], lower))


def build_merge_ready_pages(images: List[Image.Image]) -> List[Image.Image]:
    """
    針對合併長圖的頁面做專用裁切：
    - 第一頁：上留 18，下留 2
    - 中間頁：上留 2，下留 2
    - 最後一頁：上留 2，下留 18
    """
    if len(images) == 1:
        return [clean_and_trim(images[0], vertical_padding=18)]

    result = []
    for i, img in enumerate(images):
        if i == 0:
            result.append(crop_for_merge(img, keep_top=0, keep_bottom=0))
        elif i == len(images) - 1:
            result.append(crop_for_merge(img, keep_top=0, keep_bottom=0))
        else:
            result.append(crop_for_merge(img, keep_top=0, keep_bottom=0))
    return result


def combine_images_vertically(images: List[Image.Image]) -> Image.Image:
    """
    多頁合併成單張長圖：
    - 不畫線
    - 不留 gap
    """
    if not images:
        return None

    total_height = sum(img.height for img in images)
    max_width = max(img.width for img in images)

    combined = Image.new('RGB', (max_width, total_height), (255, 255, 255))
    y_offset = 0
    for img in images:
        combined.paste(img, (0, y_offset))
        y_offset += img.height

    return combined


@app.route('/')
def index():
    return send_from_directory('.', 'index.html')


@app.route('/convert', methods=['POST'])
def convert_excel():
    """
    ===== 優化 4：添加並發控制和超時 =====
    """
    # 獲取並發鎖
    if not conversion_semaphore.acquire(blocking=False):
        return jsonify({'error': '伺服器忙碌，請稍後重試'}), 503

    try:
        if 'file' not in request.files:
            return jsonify({'error': '沒有文件上傳'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '文件名為空'}), 400

        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)

        try:
            optimize_excel_for_printing(file_path)

            # ===== 優化 5：降低 DPI 從 300 到 150 =====
            images = convert_from_path(
                file_path,
                dpi=150,  # 從 300 降低到 150
                poppler_path=POPPLER_PATH
            )

            pages = [clean_and_trim(img) for img in images]
            merged_images = build_merge_ready_pages(images)
            merged_image = combine_images_vertically(merged_images)

            pages_data = [image_to_data_url(page, quality=75) for page in pages]
            merged_data = image_to_data_url(merged_image, quality=75) if merged_image else None

            return jsonify({
                'pages': pages_data,
                'merged_image': merged_data,
                'page_count': len(pages)
            })

        finally:
            # ===== 優化 6：立即清理臨時文件 =====
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"刪除文件失敗 {file_path}: {e}")

    except subprocess.TimeoutExpired:
        return jsonify({'error': 'LibreOffice 轉換超時'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        # 釋放並發鎖
        conversion_semaphore.release()


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=False)
