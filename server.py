# -*- coding: utf-8 -*-
"""
PTI Giám Định - Flask Backend Server
Chạy: python server.py
Truy cập: http://localhost:5000
"""

from flask import Flask, request, jsonify, send_file, send_from_directory, Response
import os, sys, json, re, subprocess
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env'))

def _add_cors(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    response.headers['Access-Control-Allow-Methods'] = 'GET,POST,OPTIONS'
    return response

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, 'input')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
ASSETS_DIR = os.path.join(BASE_DIR, 'assets')
DATA_JSON = os.path.join(INPUT_DIR, 'data.json')

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__, static_folder=BASE_DIR, static_url_path='')
app.after_request(_add_cors)

@app.route('/api/<path:path>', methods=['OPTIONS'])
def handle_options(path):
    return jsonify({}), 200

ALLOWED_IMAGE = {'jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp'}
ALLOWED_PDF   = {'pdf'}

GDV_DEFAULT = [
    {'ma': 'CHINH05',  'ten': 'Nguyễn Hồng Chinh', 'sdt': '0888955673'},
    {'ma': 'TUYENLM',  'ten': 'Lương Minh Tuyến',   'sdt': ''},
    {'ma': 'DUYNT',    'ten': 'Nguyễn Thế Duy',     'sdt': ''},
    {'ma': 'SONTT',    'ten': 'Trần Thanh Sơn',     'sdt': ''},
    {'ma': 'VIETNT05', 'ten': 'Nguyễn Tiến Việt',   'sdt': ''},
    {'ma': 'HUONGNV',  'ten': 'Nguyễn Văn Hướng',   'sdt': ''},
    {'ma': 'TUNGHX',   'ten': 'Hoàng Xuân Tùng',    'sdt': ''},
]

SCRIPT_MAP = {
    'A':   'tao_bien_ban_gd.py',
    'BFH': 'tao_3_bien_ban.py',
    'C':   'tao_bao_cao_so_bo.py',
    'D':   'tao_bao_lanh.py',
    'E':   'tao_tbtn_ycbt.py',
    'G':   'tao_cv_ngan_hang.py',
    'I':   'tao_cham_dut_khoi_phuc.py',
}

def allowed(filename, exts):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in exts


# ── Routes ──────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory(BASE_DIR, 'index.html')

@app.route('/assets/<path:filename>')
def serve_assets(filename):
    return send_from_directory(ASSETS_DIR, filename)

@app.route('/input/<filename>')
def serve_input(filename):
    return send_from_directory(INPUT_DIR, secure_filename(filename))


@app.route('/api/gdv-list')
def gdv_list():
    """Trả về danh sách giám định viên."""
    return jsonify({'gdv': GDV_DEFAULT})


@app.route('/api/upload-images', methods=['POST'])
def upload_images():
    """Lưu ảnh và PDF vào thư mục input/."""
    saved = []
    for key in request.files:
        f = request.files[key]
        if f.filename and allowed(f.filename, ALLOWED_IMAGE | ALLOWED_PDF):
            name = secure_filename(f.filename)
            f.save(os.path.join(INPUT_DIR, name))
            saved.append(name)
    return jsonify({'ok': True, 'saved': saved})


@app.route('/api/input-images')
def list_input_images():
    """Liệt kê ảnh hiện có trong input/."""
    images = []
    if os.path.exists(INPUT_DIR):
        for f in sorted(os.listdir(INPUT_DIR)):
            if f.lower().rsplit('.', 1)[-1] in ALLOWED_IMAGE | ALLOWED_PDF:
                images.append(f)
    return jsonify({'images': images})


@app.route('/api/save-data', methods=['POST'])
def save_data():
    """Ghi dữ liệu form xuống input/data.json cho các script tao_*.py đọc."""
    data = request.json or {}

    payload = {
        'info': data.get('info', {}),
        'phu_tung': [
            {
                'ten': pt.get('ten', ''),
                'phuong_an': pt.get('phuong_an', 'Thay thế có thu hồi'),
                'sl': int(pt.get('sl') or 1),
            }
            for pt in data.get('phu_tung', []) if pt.get('ten')
        ],
    }

    try:
        with open(DATA_JSON, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/generate', methods=['POST'])
def generate():
    """Chạy script sinh công văn theo danh sách được chọn."""
    data = request.json or {}
    docs = data.get('documents', [])  # ['A', 'BFH', 'G', 'I']

    results = []
    errors  = []

    for doc in docs:
        script_name = SCRIPT_MAP.get(doc)
        if not script_name:
            errors.append(f'Không có script cho mẫu {doc}')
            continue
        script_path = os.path.join(BASE_DIR, script_name)
        if not os.path.exists(script_path):
            errors.append(f'Script {script_name} không tìm thấy')
            continue

        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            cwd=BASE_DIR,
        )

        if result.returncode == 0:
            results.append(doc)
        else:
            err_msg = (result.stderr or result.stdout or '').strip()
            errors.append(f'{doc}: {err_msg[:300]}')

    # Liệt kê file output
    files = _list_output()
    return jsonify({'ok': True, 'generated': results, 'errors': errors, 'files': files})


@app.route('/api/output-files')
def list_output_files():
    return jsonify({'files': _list_output()})


@app.route('/api/reset', methods=['POST'])
def reset():
    """Xóa toàn bộ file trong input/ và output/ để bắt đầu hồ sơ mới."""
    deleted = []
    for folder in [INPUT_DIR, OUTPUT_DIR]:
        if os.path.exists(folder):
            for fname in os.listdir(folder):
                fpath = os.path.join(folder, fname)
                try:
                    os.remove(fpath)
                    deleted.append(fname)
                except Exception:
                    pass
    return jsonify({'ok': True, 'deleted': deleted})


def _list_output():
    if not os.path.exists(OUTPUT_DIR):
        return []
    return sorted(f for f in os.listdir(OUTPUT_DIR) if f.endswith('.docx'))


@app.route('/api/download/<filename>')
def download_file(filename):
    safe = secure_filename(filename)
    return send_from_directory(OUTPUT_DIR, safe, as_attachment=True)


@app.route('/api/check-key')
def check_key():
    """Debug: kiểm tra API key đang được load (chỉ hiện 8 ký tự đầu)."""
    key = os.environ.get('GEMINI_API_KEY', '')
    if not key:
        return jsonify({'status': 'MISSING', 'hint': 'GEMINI_API_KEY chưa được set'})
    return jsonify({'status': 'SET', 'prefix': key[:8] + '...', 'length': len(key)})


SCAN_PROMPT = """Đây là các ảnh tài liệu xe cơ giới Việt Nam (giấy đăng ký xe, giấy phép lái xe, màn hình hệ thống PTI, giấy chứng nhận bảo hiểm, báo giá sửa chữa). Hãy trích xuất thông tin và trả về JSON với các trường sau (bỏ trống "" nếu không tìm thấy):

{
  "bien_so_xe": "",
  "hang_xe": "",
  "dong_xe": "",
  "phien_ban": "",
  "nam_sx": "",
  "so_loai": "",
  "so_khung": "",
  "so_may": "",
  "trong_tai": "",
  "so_cho_ngoi": "",
  "giay_phep_luu_hanh": "",
  "gplh_tu_ngay": "",
  "gplh_den_ngay": "",
  "giay_phep_lai_xe": "",
  "hang_gplx": "",
  "gplx_tu_ngay": "",
  "gplx_den_ngay": "",
  "chu_xe": "",
  "dien_thoai_chu_xe": "",
  "dia_chi_chu_xe": "",
  "lai_xe": "",
  "dien_thoai_lai_xe": "",
  "dia_chi_lai_xe": "",
  "so_hop_dong": "",
  "so_gcn_bh": "",
  "gcn_bh_tu_ngay": "",
  "gcn_bh_den_ngay": "",
  "gia_tri_xe": "",
  "phi_bh": "",
  "dk_bs": "",
  "so_ho_so": "",
  "ma_giam_dinh_vien": "",
  "don_vi_quan_ly": "Phòng Giám định và Cứu hộ Quảng Ninh",
  "ngay_giam_dinh": "",
  "ngay_vao_gara": "",
  "ngay_tai_nan": "",
  "khoang_gio_tai_nan": "",
  "dia_diem_tai_nan": "",
  "dien_bien_tai_nan": "",
  "nguyen_nhan_tai_nan": "",
  "tien_tt": "",
  "ten_gara": "",
  "so_tk": "",
  "ten_ngan_hang": "",
  "dia_chi_ngan_hang": "",
  "phu_tung": [
    {"ten": "Tên phụ tùng hoặc hạng mục sửa chữa", "phuong_an": "Thay thế có thu hồi|Thay thế không thu hồi|Sửa chữa", "sl": 1}
  ]
}

Lưu ý quan trọng:
- Ngày tháng định dạng DD/MM/YYYY
- Chỉ trả về JSON thuần, không thêm text hay markdown nào khác
- Biển số xe giữ nguyên định dạng gốc (vd: 14A-894.22)
- Số tiền để nguyên số không có dấu phẩy (vd: 500000000)
- Nếu cùng một thông tin xuất hiện trên nhiều ảnh, ưu tiên ảnh màn hình hệ thống PTI
- Trường "phu_tung": trích xuất từ bảng báo giá / danh sách phụ tùng nếu có; "phuong_an" chỉ nhận 3 giá trị: "Thay thế có thu hồi", "Thay thế không thu hồi", hoặc "Sửa chữa"; nếu không tìm thấy danh sách phụ tùng thì để mảng rỗng []

Hướng dẫn đọc từng loại tài liệu:
- "Phiếu xác minh phí" (XMP / mẫu xe cơ giới PTI): chứa Số GCN BH → "so_gcn_bh"; Số hợp đồng → "so_hop_dong"; Thời hạn bảo hiểm từ/đến → "gcn_bh_tu_ngay"/"gcn_bh_den_ngay"; Phí bảo hiểm (dòng tổng phí hoặc "Phí BH") → "phi_bh"; Điều kiện bổ sung → "dk_bs"
- "Giấy chứng nhận bảo hiểm" (GCN BH): số GCN BH ở đầu trang, thời hạn hiệu lực, phí BH, họ tên chủ xe, biển số
- Màn hình hệ thống PTI tab "Tổn thất - Chi trả": số hồ sơ ("so_ho_so"), mã GĐV ("ma_giam_dinh_vien"), diễn biến tai nạn ("dien_bien_tai_nan"), tên gara ("ten_gara"), tiền thanh toán ("tien_tt")
- Màn hình hệ thống PTI tab "Thông tin giám định": thông tin xe, chủ xe, lái xe, ngày giám định, ngày vào gara
- "Giấy phép lái xe" (GPLX): tìm các nhãn sau trên thẻ: "Số/No:" → dãy số ngay sau đó → "giay_phep_lai_xe"; "Họ tên/Full name:" → tên ngay sau đó → PHẢI điền vào "lai_xe" (bắt buộc, dù tên trùng chủ xe); "Nơi cư trú/Address:" → địa chỉ ngay sau đó → "dia_chi_lai_xe"; "Hạng/Class:" → "hang_gplx"; "Hiệu lực từ ngày/Date" (dạng ngày/tháng/năm trên thẻ) → "gplx_tu_ngay"; "Có giá trị đến/Expires:" → "gplx_den_ngay"; "Ngày, tháng, năm sinh/Date of birth:" không dùng; số điện thoại lái xe không có trên GPLX nên để trống
- "Giấy đăng ký xe" / "Chứng nhận đăng ký xe": "Họ tên chủ xe" → "chu_xe"; "Địa chỉ" → "dia_chi_chu_xe"; "Biển số" → "bien_so_xe"; "Nhãn hiệu" → "hang_xe"; "Loại xe" → "dong_xe"; "Số khung" → "so_khung"; "Số máy" → "so_may"; "Năm sản xuất" → "nam_sx"; "Số chỗ ngồi" → "so_cho_ngoi"
- "Giấy chứng nhận kiểm định" (đăng kiểm): "Số phiếu" → "giay_phep_luu_hanh"; "Có giá trị đến" → "gplh_den_ngay"

QUAN TRỌNG — không được suy đoán lái xe: "lai_xe" và "dia_chi_lai_xe" CHỈ được lấy từ
GPLX hoặc màn hình hệ thống PTI. Nếu không có nguồn nào nêu tên lái xe thì để trống "".
TUYỆT ĐỐI không copy "chu_xe"/"dia_chi_chu_xe" sang, kể cả khi chỉ có giấy đăng ký xe —
chủ xe thường là pháp nhân (công ty) nên điền tên công ty vào chỗ lái xe là sai.
"""


@app.route('/api/scan-images', methods=['POST'])
def scan_images():
    """Dùng Gemini Vision API để trích xuất thông tin từ ảnh trong input/."""
    try:
        from google import genai as google_genai
        from google.genai import types as genai_types
    except ImportError:
        return jsonify({'error': 'Thư viện google-genai chưa cài. Chạy: pip install google-genai'}), 500

    api_key = os.environ.get('GEMINI_API_KEY', '')
    if not api_key:
        return jsonify({'error': 'Chưa cấu hình GEMINI_API_KEY trong biến môi trường'}), 500

    client = google_genai.Client(api_key=api_key)

    # Thu thập ảnh và PDF từ input/ (tối đa 20 mục, lọc trước)
    parts = []
    files_loaded = []
    if os.path.exists(INPUT_DIR):
        all_files = sorted(os.listdir(INPUT_DIR))
        media_files = [f for f in all_files
                       if f.lower().rsplit('.', 1)[-1] in ALLOWED_IMAGE | ALLOWED_PDF][:20]
        for fname in media_files:
            ext = fname.lower().rsplit('.', 1)[-1]
            fpath = os.path.join(INPUT_DIR, fname)
            if ext in ALLOWED_IMAGE:
                with open(fpath, 'rb') as fp:
                    img_bytes = fp.read()
                mime = 'image/jpeg' if ext in ('jpg', 'jpeg') else f'image/{ext}'
                parts.append(genai_types.Part.from_bytes(data=img_bytes, mime_type=mime))
                files_loaded.append(f'{fname} (ảnh)')
            elif ext == 'pdf':
                pdf_text = ''
                try:
                    import pdfplumber
                    with pdfplumber.open(fpath) as pdf:
                        pdf_text = '\n'.join(p.extract_text() or '' for p in pdf.pages)
                except Exception:
                    pass
                if pdf_text.strip():
                    parts.append(genai_types.Part.from_text(text=f'[Nội dung file PDF: {fname}]\n{pdf_text}'))
                    files_loaded.append(f'{fname} (PDF text)')
                else:
                    try:
                        import fitz  # pymupdf
                        doc = fitz.open(fpath)
                        page_count = len(doc)
                        for page in doc:
                            pix = page.get_pixmap(dpi=150)
                            parts.append(genai_types.Part.from_bytes(data=pix.tobytes('png'), mime_type='image/png'))
                        doc.close()
                        files_loaded.append(f'{fname} (PDF ảnh, {page_count} trang)')
                    except Exception:
                        files_loaded.append(f'{fname} (PDF - lỗi đọc)')

    if not parts:
        return jsonify({'error': 'Chưa có ảnh hoặc PDF nào trong thư mục input. Hãy tải file lên trước.'}), 400

    parts.append(genai_types.Part.from_text(text=SCAN_PROMPT))

    max_attempts = 4
    last_error = None
    for attempt in range(max_attempts):
        try:
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=[genai_types.Content(role='user', parts=parts)],
                config=genai_types.GenerateContentConfig(
                    thinking_config=genai_types.ThinkingConfig(thinking_budget=0)
                ),
            )
            last_error = None
            break
        except Exception as e:
            last_error = e
            err_str = str(e).lower()
            # Chỉ retry khi overloaded (503) hoặc rate limit (429)
            if '503' in err_str or 'overload' in err_str or '429' in err_str or 'rate' in err_str:
                if attempt < max_attempts - 1:
                    wait = 2 ** attempt  # 1s, 2s, 4s
                    import time
                    time.sleep(wait)
                    continue
            break
    if last_error is not None:
        return jsonify({'error': f'Lỗi gọi Gemini API: {last_error}'}), 500

    try:
        raw = response.text
    except Exception:
        finish = ''
        try:
            finish = str(response.candidates[0].finish_reason) if response.candidates else 'no candidates'
        except Exception:
            pass
        return jsonify({'error': f'Gemini không trả về nội dung (finish_reason: {finish})'}), 500

    if not raw or not raw.strip():
        return jsonify({'error': 'Gemini trả về nội dung rỗng'}), 500
    raw = raw.strip()

    # Tách JSON từ response
    m = re.search(r'\{[\s\S]+\}', raw)
    if not m:
        return jsonify({'error': 'Không trích xuất được JSON từ phản hồi AI', 'raw': raw[:500]}), 500
    try:
        info = json.loads(m.group())
    except Exception as e:
        return jsonify({'error': f'JSON không hợp lệ: {e}', 'raw': raw[:500]}), 500

    return jsonify({'ok': True, 'info': info, 'files_loaded': files_loaded})


@app.route('/api/preview/<filename>')
def preview_file(filename):
    safe = secure_filename(filename)
    path = os.path.join(OUTPUT_DIR, safe)
    if not os.path.exists(path):
        return jsonify({'error': 'File không tồn tại'}), 404
    try:
        import mammoth
        with open(path, 'rb') as f:
            result = mammoth.convert_to_html(f)
        html = result.value
        page = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  body{{font-family:'Times New Roman',serif;font-size:13pt;margin:30px 40px;line-height:1.6;color:#111}}
  table{{border-collapse:collapse;width:100%}}
  td,th{{border:1px solid #999;padding:4px 8px}}
  img{{max-width:100%}}
</style></head><body>{html}</body></html>"""
        return Response(page, mimetype='text/html; charset=utf-8')
    except ImportError:
        return Response('<p style="font-family:sans-serif;padding:20px">Thư viện <b>mammoth</b> chưa được cài. Chạy: <code>pip install mammoth</code></p>', mimetype='text/html; charset=utf-8')
    except Exception as e:
        return Response(f'<p style="font-family:sans-serif;padding:20px;color:red">Lỗi xem trước: {e}</p>', mimetype='text/html; charset=utf-8')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print('=' * 55)
    print('  PTI Giám Định - Hệ thống lập hồ sơ xe cơ giới')
    print(f'  http://localhost:{port}')
    print('=' * 55)
    app.run(debug=False, host='0.0.0.0', port=port, use_reloader=False)
