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
    
    env = os.environ.copy()
    env['HOME'] = '/tmp'
    env['PYTHONPATH'] = ''
    
    result = subprocess.run([
        'libreoffice', '--headless', '--norestore', '--nofirststartwizard',
        '--convert-to', cikis_format,
        '--outdir', cikis_klasor, giris
    ], capture_output=True, text=True, timeout=120, env=env)
    
    app.logger.info(f'LibreOffice stdout: {result.stdout}')
    app.logger.info(f'LibreOffice stderr: {result.stderr}')
    
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

# ── Filigran Ekle ───────────────────────────────────────────
@app.route('/api/watermark', methods=['POST'])
def watermark():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        metin = request.form.get('metin', 'HoPDF')
        giris = benzersiz_dosya('.pdf')
        cikis = benzersiz_dosya('.pdf')
        dosya.save(giris)
        from PyPDF2 import PdfReader, PdfWriter
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        import io

        # Türkçe karakter desteği için sistem fontu kullan
        try:
            pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
            font_name = 'DejaVu'
        except:
            font_name = 'Helvetica'

        reader = PdfReader(giris)
        writer = PdfWriter()
        for sayfa in reader.pages:
            packet = io.BytesIO()
            w = float(sayfa.mediabox.width)
            h = float(sayfa.mediabox.height)
            c = canvas.Canvas(packet, pagesize=(w, h))
            c.setFont(font_name, 40)
            c.setFillColorRGB(0.7, 0.7, 0.7, alpha=0.3)
            c.saveState()
            c.translate(w/2, h/2)
            c.rotate(45)
            c.drawCentredString(0, 0, metin)
            c.restoreState()
            c.save()
            packet.seek(0)
            from PyPDF2 import PdfReader as PR
            filigran = PR(packet).pages[0]
            sayfa.merge_page(filigran)
            writer.add_page(sayfa)
        with open(cikis, 'wb') as f:
            writer.write(f)
        dosyayi_sil(giris)
        dosyayi_sil(cikis)
        return send_file(cikis, as_attachment=True, download_name='filigranli.pdf', mimetype='application/pdf')
    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

# ── PDF Şifrele ──────────────────────────────────────────────
@app.route('/api/encrypt', methods=['POST'])
def encrypt():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        sifre = request.form.get('sifre', '1234')
        giris = benzersiz_dosya('.pdf')
        cikis = benzersiz_dosya('.pdf')
        dosya.save(giris)
        from PyPDF2 import PdfReader, PdfWriter
        reader = PdfReader(giris)
        writer = PdfWriter()
        for sayfa in reader.pages:
            writer.add_page(sayfa)
        writer.encrypt(sifre)
        with open(cikis, 'wb') as f:
            writer.write(f)
        dosyayi_sil(giris)
        dosyayi_sil(cikis)
        return send_file(cikis, as_attachment=True, download_name='sifreli.pdf', mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── Şifre Kaldır ─────────────────────────────────────────────
@app.route('/api/decrypt', methods=['POST'])
def decrypt():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        sifre = request.form.get('sifre', '')
        giris = benzersiz_dosya('.pdf')
        cikis = benzersiz_dosya('.pdf')
        dosya.save(giris)
        from PyPDF2 import PdfReader, PdfWriter
        reader = PdfReader(giris)
        if reader.is_encrypted:
            reader.decrypt(sifre)
        writer = PdfWriter()
        for sayfa in reader.pages:
            writer.add_page(sayfa)
        with open(cikis, 'wb') as f:
            writer.write(f)
        dosyayi_sil(giris)
        dosyayi_sil(cikis)
        return send_file(cikis, as_attachment=True, download_name='sifresiz.pdf', mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── Sayfa Numarası ───────────────────────────────────────────
@app.route('/api/pagenumber', methods=['POST'])
def pagenumber():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        dosya = request.files['file']
        giris = benzersiz_dosya('.pdf')
        cikis = benzersiz_dosya('.pdf')
        dosya.save(giris)
        from PyPDF2 import PdfReader, PdfWriter
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        import io
        reader = PdfReader(giris)
        writer = PdfWriter()
        for i, sayfa in enumerate(reader.pages):
            packet = io.BytesIO()
            w = float(sayfa.mediabox.width)
            h = float(sayfa.mediabox.height)
            c = canvas.Canvas(packet, pagesize=(w, h))
            c.setFont("Helvetica", 10)
            c.setFillColorRGB(0.3, 0.3, 0.3)
            c.drawCentredString(w/2, 20, str(i+1))
            c.save()
            packet.seek(0)
            from PyPDF2 import PdfReader as PR
            numara = PR(packet).pages[0]
            sayfa.merge_page(numara)
            writer.add_page(sayfa)
        with open(cikis, 'wb') as f:
            writer.write(f)
        dosyayi_sil(giris)
        dosyayi_sil(cikis)
        return send_file(cikis, as_attachment=True, download_name='numarali.pdf', mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── PDF İmzala ───────────────────────────────────────────────
@app.route('/api/sign', methods=['POST'])
def sign():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Dosya bulunamadı'}), 400
        if 'signature' not in request.files:
            return jsonify({'error': 'İmza bulunamadı'}), 400

        dosya = request.files['file']
        imza_dosya = request.files['signature']

        # İmza koordinatları ve sayfa bilgisi
        import json
        imzalar = json.loads(request.form.get('signatures', '[]'))

        giris = benzersiz_dosya('.pdf')
        imza_path = benzersiz_dosya('.png')
        cikis = benzersiz_dosya('.pdf')

        dosya.save(giris)
        imza_dosya.save(imza_path)

        import fitz  # PyMuPDF
        from PIL import Image

        doc = fitz.open(giris)

        for imza_bilgi in imzalar:
            sayfa_no = int(imza_bilgi.get('page', 0))
            x = float(imza_bilgi.get('x', 0))
            y = float(imza_bilgi.get('y', 0))
            genislik = float(imza_bilgi.get('width', 150))
            yukseklik = float(imza_bilgi.get('height', 60))

            if sayfa_no >= len(doc):
                continue

            sayfa = doc[sayfa_no]
            rect = fitz.Rect(x, y, x + genislik, y + yukseklik)
            sayfa.insert_image(rect, filename=imza_path)

        doc.save(cikis)
        doc.close()

        dosyayi_sil(giris)
        dosyayi_sil(imza_path)
        dosyayi_sil(cikis)

        return send_file(cikis, as_attachment=True,
                        download_name='imzali.pdf',
                        mimetype='application/pdf')
    except Exception as e:
        app.logger.error(traceback.format_exc())
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
