import os
import shutil
import subprocess
import sys

def find_soffice():
    # 1) Env var
    env_path = os.environ.get("SOFFICE_PATH")
    if env_path and os.path.exists(env_path):
        return env_path

    # 2) MacOS typical
    mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if os.path.exists(mac_path):
        return mac_path

    # 3) Windows typical
    win_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in win_paths:
        if os.path.exists(p):
            return p

    # 4) PATH
    p = shutil.which("soffice")
    if p:
        return p

    return None

def _lo_convert(src_path, out_dir, target="pdf"):
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice bulunamadı. Lütfen LibreOffice kur ve gerekirse SOFFICE_PATH ortam değişkenini ayarla.")

    # Run LibreOffice headless conversion
    # Example: soffice --headless --convert-to pdf --outdir outputs file.docx
    cmd = [soffice, "--headless", "--convert-to", target, "--outdir", out_dir, src_path]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if proc.returncode != 0:
        raise RuntimeError(f"LibreOffice dönüştürme hatası: {proc.stderr or proc.stdout}")

def docx_to_pdf(src_docx, dst_pdf):
    out_dir = os.path.dirname(os.path.abspath(dst_pdf))
    _lo_convert(src_docx, out_dir, target="pdf")
    # LibreOffice aynı isimle .pdf üretir
    base = os.path.splitext(os.path.basename(src_docx))[0] + ".pdf"
    produced = os.path.join(out_dir, base)
    if os.path.abspath(produced) != os.path.abspath(dst_pdf):
        os.replace(produced, dst_pdf)

def xlsx_to_pdf(src_xlsx, dst_pdf):
    out_dir = os.path.dirname(os.path.abspath(dst_pdf))
    _lo_convert(src_xlsx, out_dir, target="pdf")
    base = os.path.splitext(os.path.basename(src_xlsx))[0] + ".pdf"
    produced = os.path.join(out_dir, base)
    if os.path.abspath(produced) != os.path.abspath(dst_pdf):
        os.replace(produced, dst_pdf)