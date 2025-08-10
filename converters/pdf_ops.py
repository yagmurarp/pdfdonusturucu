import os
from pdf2docx import Converter

def pdf_to_docx(src_pdf, dst_docx):
    # Basic conversion using pdf2docx
    cv = Converter(src_pdf)
    cv.convert(dst_docx, start=0, end=None)
    cv.close()

def pdf_to_excel(src_pdf, dst_xlsx):
    # Use camelot to extract tables into Excel
    # NOTE: Requires Ghostscript installed on system.
    import camelot
    import pandas as pd

    tables = camelot.read_pdf(src_pdf, pages="all")
    if tables.n == 0:
        # Create empty Excel with info sheet
        with pd.ExcelWriter(dst_xlsx, engine="openpyxl") as writer:
            empty = pd.DataFrame({"Bilgi": ["Tablo bulunamadı. PDF tablo içermiyor olabilir."]})
            empty.to_excel(writer, index=False, sheet_name="Bilgi")
        return

    with pd.ExcelWriter(dst_xlsx, engine="openpyxl") as writer:
        for i, t in enumerate(tables, start=1):
            df = t.df
            sheet_name = f"Tablo{i}"
            # Excel sheet names max 31 chars
            if len(sheet_name) > 31:
                sheet_name = f"T{i}"
            df.to_excel(writer, index=False, sheet_name=sheet_name)