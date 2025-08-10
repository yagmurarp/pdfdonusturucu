import os
import tempfile
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename

from converters.office import docx_to_pdf, xlsx_to_pdf, find_soffice
from converters.pdf_ops import pdf_to_docx, pdf_to_excel
from converters.word_excel import docx_to_xlsx, xlsx_to_docx

ALLOWED_WORD = {'.docx'}
ALLOWED_EXCEL = {'.xlsx'}
ALLOWED_PDF = {'.pdf'}

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret")  # development only
app.config['UPLOAD_FOLDER'] = os.path.abspath("uploads")
app.config['OUTPUT_FOLDER'] = os.path.abspath("outputs")
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

def allowed(filename, allowed_exts):
    ext = os.path.splitext(filename)[1].lower()
    return ext in allowed_exts

@app.route("/", methods=["GET"])
def index():
    soffice_path = find_soffice()
    return render_template("index.html", soffice_path=soffice_path)

@app.route("/convert", methods=["POST"])
def convert():
    action = request.form.get("action")
    file = request.files.get("file")

    if not action or not file or file.filename.strip() == "":
        flash("Lütfen bir dosya seç ve işlem türünü gönder.")
        return redirect(url_for("index"))

    filename = secure_filename(file.filename)
    src_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(src_path)

    name, ext = os.path.splitext(filename)
    ext = ext.lower()

    try:
        if action == "pdf_to_word":
            if not allowed(filename, ALLOWED_PDF):
                raise ValueError("Sadece PDF kabul edilir.")
            out_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{name}.docx")
            pdf_to_docx(src_path, out_path)
            return send_file(out_path, as_attachment=True, download_name=f"{name}.docx")

        elif action == "word_to_pdf":
            if not allowed(filename, ALLOWED_WORD):
                raise ValueError("Sadece DOCX kabul edilir.")
            out_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{name}.pdf")
            docx_to_pdf(src_path, out_path)
            return send_file(out_path, as_attachment=True, download_name=f"{name}.pdf")

        elif action == "pdf_to_excel":
            if not allowed(filename, ALLOWED_PDF):
                raise ValueError("Sadece PDF kabul edilir.")
            out_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{name}.xlsx")
            pdf_to_excel(src_path, out_path)
            return send_file(out_path, as_attachment=True, download_name=f"{name}.xlsx")

        elif action == "excel_to_pdf":
            if not allowed(filename, ALLOWED_EXCEL):
                raise ValueError("Sadece XLSX kabul edilir.")
            out_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{name}.pdf")
            xlsx_to_pdf(src_path, out_path)
            return send_file(out_path, as_attachment=True, download_name=f"{name}.pdf")

        elif action == "word_to_excel":
            if not allowed(filename, ALLOWED_WORD):
                raise ValueError("Sadece DOCX kabul edilir.")
            out_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{name}.xlsx")
            docx_to_xlsx(src_path, out_path)
            return send_file(out_path, as_attachment=True, download_name=f"{name}.xlsx")

        elif action == "excel_to_word":
            if not allowed(filename, ALLOWED_EXCEL):
                raise ValueError("Sadece XLSX kabul edilir.")
            out_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{name}.docx")
            xlsx_to_docx(src_path, out_path)
            return send_file(out_path, as_attachment=True, download_name=f"{name}.docx")

        else:
            raise ValueError("Bilinmeyen işlem.")
    except Exception as e:
        app.logger.exception("Dönüşüm hatası")
        flash(f"Hata: {e}")
        return redirect(url_for("index"))
    finally:
        # Geçici dosyayı sil
        try:
            os.remove(src_path)
        except Exception:
            pass

if __name__ == "__main__":
    app.run(debug=True)