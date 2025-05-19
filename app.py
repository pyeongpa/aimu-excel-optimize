from flask import Flask, render_template, request, send_file, redirect, url_for
from werkzeug.utils import secure_filename
from excel_processor import process_excel
import os

# ✅ 현재 실행 경로 기준으로 경로 설정 (Render에서도 잘 동작)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
RESULT_FOLDER = os.path.join(BASE_DIR, 'optimized_files')

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 최대 16MB 업로드 허용

# ✅ 폴더 없으면 자동 생성
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    print("📌 /upload 요청 도착")

    if 'file' not in request.files:
        print("❌ 파일이 request에 없음")
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        print("❌ 파일명이 비어 있음")
        return redirect(url_for('index'))

    filename = secure_filename(file.filename)
    upload_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)

    try:
        file.save(upload_path)
        print(f"📂 파일 저장 완료: {upload_path}")
    except Exception as e:
        print(f"🚨 파일 저장 중 오류: {e}")
        return "파일 저장 실패"

    try:
        print("📥 process_excel 호출 전")
        result_path = process_excel(upload_path, filename)
        print(f"📦 최적화 완료: {result_path}")
    except Exception as e:
        print(f"🚨 process_excel 호출 중 오류: {e}")
        return "엑셀 처리 실패"

    return redirect(url_for('download_file', filename=os.path.basename(result_path)))

@app.route('/download/<filename>')
def download_file(filename):
    download_path = os.path.join(app.config['RESULT_FOLDER'], filename)

    if not os.path.exists(download_path):
        print(f"❌ 다운로드 대상 파일 없음: {download_path}")
        return "파일을 찾을 수 없습니다."

    return send_file(download_path, as_attachment=True)

# ✅ Render 등 클라우드 환경에서 필요함
if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
