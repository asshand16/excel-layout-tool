import os
import subprocess
import base64
import platform
import tempfile
import zipfile
import xml.etree.ElementTree as ET

import openpyxl
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
from PIL import Image, ImageChops, ImageDraw

app = Flask(__name__, static_folder='.')
CORS(app)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'temp')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 環境偵測
if platform.system() == 'Windows':
    LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
    POPPLER_PATH = r"C:\poppler\library\bin"
else:
    LIBREOFFICE_PATH = '/usr/bin/soffice'
    POPPLER_PATH = None


def remove_print_titles_by_xml(xlsx_path: str) -> None:
    """
    直接修改 xlsx 內部的 xl/workbook.xml，穩定移除 _xlnm.Print_Titles。
    Windows 下暫存檔必須建立在同一個磁碟，否則 os.replace 可能出現 WinError 17。
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
    修正 Excel 列印設定，避免：
    1. Print Titles 重複列印造成 PDF 底部出現看似多出的欄位
    2. Header/Footer 造成奇怪字元或空白
    3. scale 與 fitToWidth/fitToHeight 互相衝突

    注意：
    - 保留原本 print_area（若使用者有設定）
    - 不主動動 hidden rows / hidden cols
    """
    wb = openpyxl.load_workbook(file_path)

    for ws in wb.worksheets:
        # 清除列印標題（表層）
        try:
            ws.print_title_rows = None
            ws.print_title_cols = None
        except Exception as e:
            print(f'⚠️ 清除 print title 屬性失敗: {e}')

        # 清除頁首頁尾，避免雜訊
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

        # 清掉 scale，改用 fitToWidth / fitToHeight
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

    # 底層 XML 再移除一次，確保 _xlnm.Print_Titles 真的不在
    remove_print_titles_by_xml(file_path)


def clean_and_trim(img: Image.Image) -> Image.Image:
    """只裁掉上下純白，保留左右完整寬度。"""
    if img.mode != 'RGB':
        img = img.convert('RGB')

    bg = Image.new(img.mode, img.size, (255, 255, 255))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()

    if bbox:
        upper = max(0, bbox[1] - 5)
        lower = min(img.size[1], bbox[3] + 5)
        return img.crop((0, upper, img.size[0], lower))

    return img


def combine_images_vertically(images, gap=24, add_separator=True):
    """
    多頁 PDF 合併成一張長圖，頁面間加入分隔線，
    避免第 2 頁表頭看起來像第 1 頁底部多出欄位。
    """
    if len(images) == 1:
        return images[0]

    max_width = max(img.size[0] for img in images)
    total_height = sum(img.size[1] for img in images) + gap * (len(images) - 1)

    canvas = Image.new('RGB', (max_width, total_height), (255, 255, 255))
    draw = ImageDraw.Draw(canvas)

    y = 0
    for index, img in enumerate(images):
        canvas.paste(img, (0, y))
        y += img.size[1]

        if index < len(images) - 1:
            if add_separator:
                line_y = y + gap // 2
                draw.line((40, line_y, max_width - 40, line_y), fill=(200, 200, 200), width=2)
            y += gap

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
        return jsonify({'error': '【格式太舊】請在 Excel 中將檔案另存新檔為 .xlsx 後再上傳！'}), 400

    excel_path = os.path.join(UPLOAD_FOLDER, filename)
    pdf_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '.pdf')
    img_path = os.path.join(UPLOAD_FOLDER, filename.rsplit('.', 1)[0] + '_result.jpg')

    try:
        file.save(excel_path)

        # 預處理失敗就直接報錯，不可吞掉
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
            images = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
        else:
            images = convert_from_path(pdf_path, dpi=300)

        if not images:
            raise RuntimeError('PDF 已生成，但無法轉成圖片。')

        trimmed_images = [clean_and_trim(img) for img in images]
        final_image = combine_images_vertically(trimmed_images, gap=24, add_separator=True)
        final_image.save(img_path, 'JPEG', quality=95)

        with open(img_path, 'rb') as img_file:
            encoded_string = base64.b64encode(img_file.read()).decode('utf-8')

        return jsonify({
            'image': 'data:image/jpeg;base64,' + encoded_string,
            'pages': len(images),
            'message': '轉檔成功'
        })

    except subprocess.CalledProcessError as e:
        print('\n❌ LibreOffice 轉檔錯誤')
        print('returncode:', e.returncode)
        print('stdout:', e.stdout)
        print('stderr:', e.stderr)
        return jsonify({
            'error': f'LibreOffice 轉檔失敗: {e.stderr or e.stdout or str(e)}'
        }), 500

    except Exception as e:
        print(f'\n❌ 轉檔錯誤: {e}\n')
        return jsonify({'error': str(e)}), 500

    finally:
        for path in [excel_path, pdf_path, img_path]:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception as cleanup_error:
                print(f'⚠️ 清理檔案失敗 {path}: {cleanup_error}')


if __name__ == '__main__':
    app.run(port=8080, debug=True)
