# -*- coding: utf-8 -*-
"""Find text that actually breaks the layout — measured by PowerPoint, not eyeballed.

python-pptx has no reliable autofit and cannot tell you whether text outgrew its box, so
the usual advice is "render it and look", which turns every tweak into a render-inspect
cycle. PowerPoint already knows: COM exposes `TextFrame2.TextRange.BoundWidth/BoundHeight`
— the size the text ACTUALLY occupies after layout.

⚠️ The subtlety that makes a naive version useless: **a PowerPoint textbox does not clip.**
Text bigger than its box simply spills and stays fully visible. A first cut of this script
reported "clipped!" for every box whose text was taller than the box and was wrong on 3/3
cases — all three rendered perfectly (a 2-line legend under a formula, a wrapped subtitle).
So spill alone is NOT an error. What actually breaks a slide is spill that

  * runs off the slide edge  -> text genuinely disappears, or
  * lands on top of another shape -> overlapping, unreadable text

Those are what this reports. Plain spill is listed only with --all, as information.

    python audit_text_fit.py DECK.pptx           # real problems only
    python audit_text_fit.py DECK.pptx --all     # every text shape + spill amount

Exit 1 if anything runs off-slide or collides, so a build can gate on it.
Requires PowerPoint (Windows); opens the deck read-only and never saves.

Blind spot: a table shape has no shape-level TextFrame (its text lives per-cell), so this script
cannot measure it and table rows never appear in the report — not even as spill. PowerPoint expands
a table's row height to fit its content, which can push the table past the slide bottom with no
warning here. For any deck with tables, also run the lab's `agent/tools/deck_render_audit.py`, which
checks the rendered PNG's bottom edge instead of measured text bounds and catches exactly this case.
"""
import argparse
import io
import os
import sys

if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PT = 72.0


def measure(path):
    """Per slide: slide size + [(name, l, t, w, h, text_w, text_h, text)] for every shape."""
    import win32com.client as win32
    app = win32.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(os.path.abspath(path), WithWindow=False, ReadOnly=True)
    slides = []
    try:
        sw, sh_ = float(pres.PageSetup.SlideWidth) / PT, float(pres.PageSetup.SlideHeight) / PT
        for si in range(1, pres.Slides.Count + 1):
            sl = pres.Slides(si)
            shapes = []
            for shi in range(1, sl.Shapes.Count + 1):
                sp = sl.Shapes(shi)
                try:
                    l, t = float(sp.Left) / PT, float(sp.Top) / PT
                    w, h = float(sp.Width) / PT, float(sp.Height) / PT
                    name = sp.Name
                except Exception:
                    continue
                tw = th = None
                txt = ""
                try:
                    if sp.HasTextFrame and sp.TextFrame.HasText:
                        tr = sp.TextFrame2.TextRange
                        tw, th = float(tr.BoundWidth) / PT, float(tr.BoundHeight) / PT
                        txt = tr.Text
                except Exception:
                    pass
                try:                       # msoPicture=13, msoLinkedPicture=11
                    is_pic = int(sp.Type) in (11, 13)
                except Exception:
                    is_pic = False
                try:                       # a table has no TextFrame, so it would be invisible
                    is_tbl = int(sp.HasTable) == -1      # to collision checks unless flagged
                except Exception:
                    is_tbl = False
                shapes.append(dict(name=name, l=l, t=t, w=w, h=h, tw=tw, th=th,
                                   txt=txt, is_pic=is_pic, is_tbl=is_tbl))
            slides.append((si, shapes))
    finally:
        pres.Close()
    return sw, sh_, slides


def overlap(a_l, a_t, a_r, a_b, b_l, b_t, b_r, b_b, tol=0.02):
    ox = min(a_r, b_r) - max(a_l, b_l)
    oy = min(a_b, b_b) - max(a_t, b_t)
    return ox > tol and oy > tol


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--all", action="store_true", help="also list plain spill (not an error)")
    ap.add_argument("--tol", type=float, default=0.02, help="inches of slack (default .02)")
    ap.add_argument("--min", dest="minover", type=float, default=0.05,
                    help="ignore overruns smaller than this, in inches (default .05). Below "
                         "roughly this, nothing is visibly cut at normal font sizes — a real "
                         "deck reported +0.02in and rendered perfectly clean.")
    a = ap.parse_args()

    sw, sh_, slides = measure(a.deck)
    offslide, collide, spill = [], [], []

    for si, shapes in slides:
        for s in shapes:
            if s["tw"] is None:
                continue
            # region the text really occupies (text is left/top anchored in these decks)
            t_r, t_b = s["l"] + s["tw"], s["t"] + s["th"]
            grew = s["th"] > s["h"] + a.tol or s["tw"] > s["w"] + a.tol
            if grew:
                spill.append((si, s, max(s["th"] - s["h"], s["tw"] - s["w"])))
            off = max(t_b - sh_, t_r - sw)
            if off > a.minover:
                offslide.append((si, s, off))
                continue
            if off > a.tol:
                continue                    # past the edge but not visibly — geometry noise
            if not grew:
                continue
            # Only the strip that spilled past the declared box can hit a neighbour — and only
            # a neighbour that CARRIES CONTENT counts. Landing on a background fill or an accent
            # bar is normal layering, not a defect: a first cut flagged all 3 spills in a clean
            # deck because a full-slide "Rectangle 1" background overlaps everything.
            for o in shapes:
                if o is s or o["w"] <= 0:
                    continue
                carries_content = (o["tw"] is not None) or o["is_pic"] or o["is_tbl"]
                if not carries_content:
                    continue
                if overlap(s["l"], s["t"] + s["h"], t_r, t_b,          # spilled strip
                           o["l"], o["t"], o["l"] + o["w"], o["t"] + o["h"], a.minover):
                    collide.append((si, s, o))
                    break

    n_text = sum(1 for _, shp in slides for s in shp if s["tw"] is not None)
    print(f"{os.path.basename(a.deck)} — 텍스트 도형 {n_text}개 (PowerPoint 실측)")

    if offslide:
        print(f"\n[슬라이드 밖 {len(offslide)}건] 글자가 화면 밖으로 나갑니다")
        for si, s, over in offslide:
            print(f"  슬{si:3d} {s['name'][:24]:24s} 초과 {over:+.2f}in  "
                  f"{s['txt'].replace(chr(13),' / ')[:52]}")
    if collide:
        print(f"\n[겹침 {len(collide)}건] 넘친 글자가 다른 도형 위에 얹힙니다")
        for si, s, o in collide:
            print(f"  슬{si:3d} {s['name'][:22]:22s} → {o['name'][:22]:22s}  "
                  f"{s['txt'].replace(chr(13),' / ')[:40]}")

    if a.all and spill:
        print(f"\n[참고] 박스보다 커진 텍스트 {len(spill)}건 — PowerPoint는 자르지 않으므로 "
              f"그 자체로는 문제 아님")
        for si, s, over in sorted(spill, key=lambda x: -x[2])[:15]:
            print(f"  슬{si:3d} {s['name'][:22]:22s} +{over:.2f}in  "
                  f"{s['txt'].replace(chr(13),' / ')[:44]}")

    if not offslide and not collide:
        print(f"  OK 슬라이드 밖 0건, 겹침 0건"
              f"{f'  (박스 초과 {len(spill)}건은 정상 범위)' if spill else ''}")
    return 1 if (offslide or collide) else 0


if __name__ == "__main__":
    sys.exit(main())
