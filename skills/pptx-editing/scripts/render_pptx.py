# -*- coding: utf-8 -*-
"""render_pptx.py — render a deck to PDF (+ per-slide PNG) for the visual check that structure
checks cannot do. PowerPoint COM (Windows) is primary; falls back to LibreOffice for the PDF.

    python render_pptx.py FILE.pptx                  # PDF + PNGs beside the deck (the usual call)
    python render_pptx.py FILE.pptx --png-dir DIR    # PNGs somewhere else
    python render_pptx.py FILE.pptx --pdf OUT.pdf    # PDF only -- no PNGs are written

You MUST open the PNGs and LOOK: clipped card lines, footnotes bleeding off the slide, overrun titles
and images are invisible to python-pptx. Requires pywin32 (PowerPoint installed) for PNG export.

⚠ Asking for --pdf ONLY does not scatter PNGs any more (2026-08-31). It used to always create
`<deck>_png/` next to the source, which quietly littered folders that get zipped and shipped --
the caller who wanted an email attachment got 36 PNGs in the distribution folder and had to
notice and delete them. Bare call (no --pdf, no --png-dir) keeps the old render-and-look
behaviour, because there the PNGs ARE the point."""
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
    # Only fall back to "<deck>_png beside the source" for the bare render-and-look call.
    # If the caller explicitly asked for a PDF and said nothing about PNGs, they want a PDF.
    if a.png_dir:
        png_dir = a.png_dir
    elif a.pdf:
        png_dir = None
    else:
        png_dir = os.path.splitext(a.file)[0] + "_png"
    try:
        via_com(a.file, pdf, png_dir)
        if png_dir:
            pngs = sorted(glob.glob(os.path.join(png_dir, "*.PNG")) + glob.glob(os.path.join(png_dir, "*.png")))
            print(f"PDF: {pdf}\nPNGs: {len(pngs)} in {png_dir}  (OPEN THEM AND LOOK)")
        else:
            print(f"PDF: {pdf}\n(PNG 없음 — --png-dir 를 주면 만듭니다. 눈으로 볼 거면 인자 없이 부르세요.)")
    except Exception as e:
        print("COM render failed (%s); trying LibreOffice for PDF only..." % e)
        via_soffice(a.file, pdf)
        print(f"PDF: {pdf}  (LibreOffice; layout is approximate for .pptx)")

if __name__ == "__main__":
    main()
