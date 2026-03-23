from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import tempfile
import uuid
from datetime import datetime
import threading
import time
import traceback
import subprocess

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = tempfile.gettempdir()

def dosyayi_sil(path, gecikme=300):
    def sil():
        time.sleep(gecikme)
        try:
            if os.path.exists(path):
                os.remove(path)
        except:
            pass
    thread = threading.Thread(target=sil)
    thread.daemon = True
    thread.start()

def benzersiz_dosya(uzanti):
    return os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}{uzanti}")

def libreoffice_donustur(giris, cikis_format, cikis_uzanti):
    """LibreOffice ile herhangi bir dosyayı dönüştür"""
    cikis_klasor = os.path.dirname(giris)
    result = subprocess.run([
        'libreoffice', '--headless', '--convert-to', cikis_format,
        '--outdir', cikis_klasor, giris
    ], capture_output=True, text=True, timeout=120)
    
    # LibreOffice çıktı dosyasını bul
    giris_adi = os.path.splitext(giris)[0]
    cikis = giris_adi + cikis_uzanti
    
    if os.path.exists(cikis):
        return cikis
    
    raise Exception(f'LibreOffice hatası: {result.stderr}')

# ── Word → PDF ──────────────────────────────────────────────
@app.route('/api/word2pdf', methods=['POST'])
def word2pdf():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        uzanti = '.docx' if dosya.filename.endswith('.docx') else '.doc'
        giris = benzersiz_dosya(uzanti)
        dosya.save(giris)
        
        cikis = libreoffice_donustur(giris, 'pdf', '.pdf')
        dosyayi_sil(giris)
        dosyayi_sil(cikis)
        return send_file(cikis, as_attachment=True,
                        download_name=dosya.filename.rsplit('.',1)[0]+'.pdf',
                        mimetype='application/pdf')
    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# ── PDF → Word ──────────────────────────────────────────────
@app.route('/api/pdf2word', methods=['POST'])
def pdf2word():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        giris = benzersiz_dosya('.pdf')
        cikis = benzersiz_dosya('.docx')
        dosya.save(giris)
        from pdf2docx import Converter
        cv = Converter(giris)
        cv.convert(cikis)
        cv.close()
        dosyayi_sil(giris)
        dosyayi_sil(cikis)
        return send_file(cikis, as_attachment=True,
                        download_name=dosya.filename.rsplit('.',1)[0]+'.docx',
                        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# ── Excel → PDF ──────────────────────────────────────────────
@app.route('/api/excel2pdf', methods=['POST'])
def excel2pdf():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        uzanti = '.xlsx' if dosya.filename.endswith('.xlsx') else '.xls'
        giris = benzersiz_dosya(uzanti)
        dosya.save(giris)
        cikis = libreoffice_donustur(giris, 'pdf', '.pdf')
        dosyayi_sil(giris)
        dosyayi_sil(cikis)
        return send_file(cikis, as_attachment=True,
                        download_name=dosya.filename.rsplit('.',1)[0]+'.pdf',
                        mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── PowerPoint → PDF ─────────────────────────────────────────
@app.route('/api/ppt2pdf', methods=['POST'])
def ppt2pdf():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        uzanti = '.pptx' if dosya.filename.endswith('.pptx') else '.ppt'
        giris = benzersiz_dosya(uzanti)
        dosya.save(giris)
        cikis = libreoffice_donustur(giris, 'pdf', '.pdf')
        dosyayi_sil(giris)
        dosyayi_sil(cikis)
        return send_file(cikis, as_attachment=True,
                        download_name=dosya.filename.rsplit('.',1)[0]+'.pdf',
                        mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── PDF Birleştir ────────────────────────────────────────────
@app.route('/api/merge', methods=['POST'])
def merge():
    try:
        dosyalar = request.files.getlist('files')
        if len(dosyalar) < 2:
            return jsonify({'error': 'En az 2 dosya gerekli'}), 400
        from PyPDF2 import PdfMerger
        merger = PdfMerger()
        girdi_dosyalar = []
        for dosya in dosyalar:
            giris = benzersiz_dosya('.pdf')
            dosya.save(giris)
            merger.append(giris)
            girdi_dosyalar.append(giris)
        cikis = benzersiz_dosya('.pdf')
        merger.write(cikis)
        merger.close()
        for f in girdi_dosyalar:
            dosyayi_sil(f)
        dosyayi_sil(cikis)
        return send_file(cikis, as_attachment=True, download_name='birlestirilmis.pdf', mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── PDF Sıkıştır ─────────────────────────────────────────────
@app.route('/api/compress', methods=['POST'])
def compress():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        giris = benzersiz_dosya('.pdf')
        cikis = benzersiz_dosya('.pdf')
        dosya.save(giris)
        from PyPDF2 import PdfReader, PdfWriter
        reader = PdfReader(giris)
        writer = PdfWriter()
        for sayfa in reader.pages:
            sayfa.compress_content_streams()
            writer.add_page(sayfa)
        with open(cikis, 'wb') as f:
            writer.write(f)
        dosyayi_sil(giris)
        dosyayi_sil(cikis)
        return send_file(cikis, as_attachment=True, download_name='sikistirilmis.pdf', mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── PDF → JPG ────────────────────────────────────────────────
@app.route('/api/pdf2jpg', methods=['POST'])
def pdf2jpg():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        giris = benzersiz_dosya('.pdf')
        dosya.save(giris)
        from pdf2image import convert_from_path
        import zipfile
        sayfalar = convert_from_path(giris, dpi=150)
        zip_path = benzersiz_dosya('.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for i, sayfa in enumerate(sayfalar):
                jpg_path = benzersiz_dosya('.jpg')
                sayfa.save(jpg_path, 'JPEG', quality=85)
                zipf.write(jpg_path, f'sayfa_{i+1}.jpg')
                dosyayi_sil(jpg_path)
        dosyayi_sil(giris)
        dosyayi_sil(zip_path)
        return send_file(zip_path, as_attachment=True, download_name='sayfalar.zip', mimetype='application/zip')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── JPG → PDF ────────────────────────────────────────────────
@app.route('/api/jpg2pdf', methods=['POST'])
def jpg2pdf():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        giris = benzersiz_dosya('.jpg')
        dosya.save(giris)
        from PIL import Image
        img = Image.open(giris).convert('RGB')
        cikis = benzersiz_dosya('.pdf')
        img.save(cikis)
        dosyayi_sil(giris)
        dosyayi_sil(cikis)
        return send_file(cikis, as_attachment=True,
                        download_name=dosya.filename.rsplit('.',1)[0]+'.pdf',
                        mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── PDF Böl ──────────────────────────────────────────────────
@app.route('/api/split', methods=['POST'])
def split():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        giris = benzersiz_dosya('.pdf')
        dosya.save(giris)
        from PyPDF2 import PdfReader, PdfWriter
        import zipfile
        reader = PdfReader(giris)
        zip_path = benzersiz_dosya('.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for i, sayfa in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(sayfa)
                pdf_path = benzersiz_dosya('.pdf')
                with open(pdf_path, 'wb') as f:
                    writer.write(f)
                zipf.write(pdf_path, f'sayfa_{i+1}.pdf')
                dosyayi_sil(pdf_path)
        dosyayi_sil(giris)
        dosyayi_sil(zip_path)
        return send_file(zip_path, as_attachment=True, download_name='bolunmus.zip', mimetype='application/zip')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── PDF Döndür ───────────────────────────────────────────────
@app.route('/api/rotate', methods=['POST'])
def rotate():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        derece = int(request.form.get('derece', 90))
        giris = benzersiz_dosya('.pdf')
        cikis = benzersiz_dosya('.pdf')
        dosya.save(giris)
        from PyPDF2 import PdfReader, PdfWriter
        reader = PdfReader(giris)
        writer = PdfWriter()
        for sayfa in reader.pages:
            sayfa.rotate(derece)
            writer.add_page(sayfa)
        with open(cikis, 'wb') as f:
            writer.write(f)
        dosyayi_sil(giris)
        dosyayi_sil(cikis)
        return send_file(cikis, as_attachment=True, download_name='dondurulmus.pdf', mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── Sağlık Kontrolü ─────────────────────────────────────────
@app.route('/api/health', methods=['GET'])
def health():
    durum = {'status': 'ok', 'site': 'hopdf.com', 'zaman': str(datetime.now())}
    # LibreOffice kurulu mu?
    try:
        result = subprocess.run(['libreoffice', '--version'], capture_output=True, text=True, timeout=10)
        durum['libreoffice'] = result.stdout.strip()
    except Exception as e:
        durum['libreoffice'] = f'YOK: {str(e)}'
    return jsonify(durum)

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)
