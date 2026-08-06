# -*- coding: utf-8 -*-
"""render_pptx.py — render a deck to PDF (+ per-slide PNG) for the visual check that structure
checks cannot do. PowerPoint COM (Windows) is primary; falls back to LibreOffice for the PDF.

    python render_pptx.py FILE.pptx [--pdf OUT.pdf] [--png-dir DIR]

You MUST open the PNGs and LOOK: clipped card lines, footnotes bleeding off the slide, overrun titles
and images are invisible to python-pptx. Requires pywin32 (PowerPoint installed) for PNG export."""
import sys, io, os, argparse, glob, subprocess
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PP_PDF = 32   # ppSaveAsPDF
PP_PNG = 18   # ppSaveAsPNG (exports every slide as a PNG into a folder)

def via_com(path, pdf, png_dir):
    import win32com.client
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(os.path.abspath(path), WithWindow=False)
    try:
        pres.SaveAs(os.path.abspath(pdf), PP_PDF)
        if png_dir:
            os.makedirs(png_dir, exist_ok=True)
            pres.SaveAs(os.path.abspath(png_dir), PP_PNG)  # writes Slide1.PNG, ... into png_dir
    finally:
        pres.Close(); app.Quit()

def via_soffice(path, pdf):
    outdir = os.path.dirname(os.path.abspath(pdf)) or "."
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir,
                    os.path.abspath(path)], check=True)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--pdf", default=None)
    ap.add_argument("--png-dir", default=None)
    a = ap.parse_args()
    pdf = a.pdf or os.path.splitext(a.file)[0] + ".pdf"
    png_dir = a.png_dir or (os.path.splitext(a.file)[0] + "_png")
    try:
        via_com(a.file, pdf, png_dir)
        pngs = sorted(glob.glob(os.path.join(png_dir, "*.PNG")) + glob.glob(os.path.join(png_dir, "*.png")))
        print(f"PDF: {pdf}\nPNGs: {len(pngs)} in {png_dir}  (OPEN THEM AND LOOK)")
    except Exception as e:
        print("COM render failed (%s); trying LibreOffice for PDF only..." % e)
        via_soffice(a.file, pdf)
        print(f"PDF: {pdf}  (LibreOffice; layout is approximate for .pptx)")

if __name__ == "__main__":
    main()
