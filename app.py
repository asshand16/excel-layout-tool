import os
import io
import base64
import platform
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET
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


def str_to_bool(value, default=False):
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {'1', 'true', 'yes', 'y', 'on'}


def image_to_data_url(img: Image.Image, quality: int = 80) -> str:
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
                data = zin.read(item.filename)

                if item.filename == 'xl/workbook.xml':
                    root = ET.fromstring(data)
                    defined_names = root.find('main:definedNames', ns)

                    if defined_names is not None:
                        removed = []
                        for dn in list(defined_names):
                            name = dn.attrib.get('name', '')
                            if name == '_xlnm.Print_Titles':
                                removed.append(name)
                                defined_names.remove(dn)

                        if len(defined_names) == 0:
                            root.remove(defined_names)

                        if removed:
                            print(f'✅ 已從 workbook.xml 移除: {removed}')

                    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)

                zout.writestr(item, data)

        os.replace(tmp_xlsx, xlsx_path)

    finally:
        if os.path.exists(tmp_xlsx):
            try:
                os.remove(tmp_xlsx)
            except Exception:
                pass


def optimize_excel_for_printing(file_path: str) -> None:
    """
    通用型列印優化：
    1. 移除 Print_Titles，避免每頁重複表頭
    2. 清空 header/footer，避免雜訊
    3. 保留多頁，不強制壓成單頁
    """
    wb = openpyxl.load_workbook(file_path)

    for ws in wb.worksheets:
        try:
            ws.print_title_rows = None
            ws.print_title_cols = None
        except Exception as e:
            print(f'⚠️ 清除 print title 屬性失敗: {e}')

        for header_footer in [
            ws.oddHeader, ws.oddFooter,
            ws.evenHeader, ws.evenFooter,
            ws.firstHeader, ws.firstFooter,
        ]:
            if header_footer:
                if hasattr(header_footer, 'left'):
                    header_footer.left.text = None
                if hasattr(header_footer, 'center'):
                    header_footer.center.text = None
                if hasattr(header_footer, 'right'):
                    header_footer.right.text = None

        ws.page_setup.scale = None
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

        if ws.sheet_properties.pageSetUpPr is None:
            from openpyxl.worksheet.properties import PageSetupProperties
            ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        else:
            ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.save(file_path)
    wb.close()
    remove_print_titles_by_xml(file_path)


def get_content_bbox(img: Image.Image):
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
    - 不 overlap
    真正消除空隙靠的是 merge 前專用裁切。
    """
    if len(images) == 1:
        return images[0]

    normalized = []
    for img in images:
        if img.mode != 'RGB':
            img = img.convert('RGB')
        normalized.append(img)

    max_width = max(img.size[0] for img in normalized)
    total_height = sum(img.size[1] for img in normalized)
    canvas = Image.new('RGB', (max_width, total_height), (255, 255, 255))

    y = 0
    for img in normalized:
        canvas.paste(img, (0, y))
        y += img.size[1]

    return canvas


@app.route('/')
def index():
    return send_from_directory('.', 'index.html')


@app.route('/convert-excel', methods=['POST'])
def convert_excel():
    if 'file' not in request.files:
        return jsonify({'error': '沒有上傳檔案'}), 400

    file = request.files['file']
    filename = secure_filename(file.filename)

    if not filename:
        return jsonify({'error': '檔名無效'}), 400

    if filename.lower().endswith('.xls'):
        return jsonify({'error': '【格式太舊】請先另存為 .xlsx 再上傳！'}), 400

    merge_pages = str_to_bool(request.form.get('merge_pages'), default=True)

    excel_path = os.path.join(UPLOAD_FOLDER, filename)
    pdf_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.pdf')

    try:
        file.save(excel_path)
        optimize_excel_for_printing(excel_path)

        result = subprocess.run(
            [
                LIBREOFFICE_PATH,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', UPLOAD_FOLDER,
                excel_path,
            ],
            check=True,
            capture_output=True,
            text=True,
        )
        print('LibreOffice stdout:', result.stdout)
        print('LibreOffice stderr:', result.stderr)

        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f'PDF 轉檔失敗，找不到輸出檔案: {pdf_path}')

        if POPPLER_PATH:
            images = convert_from_path(pdf_path, dpi=180, poppler_path=POPPLER_PATH)
        else:
            images = convert_from_path(pdf_path, dpi=180)

        if not images:
            raise RuntimeError('PDF 已生成，但無法轉成圖片。')

        # pages[]：保留舒適邊界
        page_preview_images = [clean_and_trim(img, vertical_padding=18) for img in images]
        page_data_urls = [image_to_data_url(img) for img in page_preview_images]

        response = {
            'message': '轉檔成功',
            'pages': page_data_urls,
            'page_count': len(page_data_urls),
            'merge_pages': merge_pages,
        }

        if len(page_data_urls) == 1:
            response['image'] = page_data_urls[0]
            response['notice'] = '單頁輸出。'
        else:
            if merge_pages:
                merge_ready_images = build_merge_ready_pages(images)
                merged = combine_images_vertically(merge_ready_images)
                merged_data_url = image_to_data_url(merged)
                response['merged_image'] = merged_data_url
                response['image'] = merged_data_url
            else:
                response['image'] = page_data_urls[0]

            response['notice'] = '此檔案為多頁內容。正式顯示請使用 pages[] 分頁渲染；image/merged_image 僅為相容舊前端或預覽。'

        return jsonify(response)

    except subprocess.CalledProcessError as e:
        print('\n❌ LibreOffice 轉檔錯誤')
        print('returncode:', e.returncode)
        print('stdout:', e.stdout)
        print('stderr:', e.stderr)
        return jsonify({'error': f'LibreOffice 轉檔失敗: {e.stderr or e.stdout or str(e)}'}), 500

    except Exception as e:
        print(f'\n❌ 轉檔錯誤: {e}\n')
        return jsonify({'error': str(e)}), 500

    finally:
        for path in [excel_path, pdf_path]:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception as cleanup_error:
                print(f'⚠️ 清理檔案失敗 {path}: {cleanup_error}')


if __name__ == '__main__':
    app.run(port=8080, debug=True)
