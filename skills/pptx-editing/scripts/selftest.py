# -*- coding: utf-8 -*-
"""selftest.py — prove pptx_kit works without a real template.

Exercises the part that actually bit us: writing a speaker note to a notes slide whose master has
NO body placeholder (the case where python-pptx returns notes_text_frame is None). We reproduce that
by stripping the placeholder, then assert the note round-trips (write -> save -> reopen -> read).
    python selftest.py   # prints PASS / raises on failure
"""
import sys, io, os, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from pptx import Presentation
from pptx_kit import (new_deck, blank_slide_layout, speaker_note, fit_picture, overflows,
                      hang, text_units, wrapped_row_count, check_surface_leaks, save_and_check)

def _strip_notes_placeholders(slide):
    ns = slide.notes_slide                       # creates the notes slide
    for ph in list(ns.placeholders):
        ph._element.getparent().remove(ph._element)   # force notes_text_frame -> None

def test_text_and_leak_mechanics():
    # text_units: Hangul/Jamo count double, Latin counts single
    assert text_units("ab") == 2
    assert text_units("가나") == 4
    assert text_units("a가") == 3

    # wrapped_row_count: a line under wrap_units is 1 row; over it wraps to more
    assert wrapped_row_count(["short"], wrap_units=50) == 1
    assert wrapped_row_count(["가" * 30], wrap_units=50) == 2   # 30 hangul = 60 units -> wraps once
    assert wrapped_row_count(["a", "b"], wrap_units=50) == 2     # two separate 1-row lines

    # check_surface_leaks: no default term list (caller-supplied only), finds an exact hit
    prs = new_deck()
    s = prs.slides.add_slide(blank_slide_layout(prs))
    s.shapes.add_textbox(0, 0, 100, 100).text_frame.text = "이전 버전과 비교하면"
    assert check_surface_leaks(prs, []) == [], "empty term list must report nothing"
    hits = check_surface_leaks(prs, ["이전 버전"])
    assert len(hits) == 1 and hits[0][1] == "이전 버전", f"expected one leak hit, got {hits}"

    # save_and_check: raises on a leak even though the file is on disk
    with tempfile.TemporaryDirectory() as d:
        out = os.path.join(d, "leak.pptx")
        try:
            save_and_check(prs, out, leak_terms=["이전 버전"])
            raise AssertionError("save_and_check should have raised on a known leak term")
        except SystemExit:
            pass
        assert os.path.exists(out), "save_and_check must still write the file before gating"

    print("PASS  text_units/wrapped_row_count estimate wrap correctly; "
          "check_surface_leaks/save_and_check gate on caller-supplied terms only.")

def main():
    prs = new_deck()                             # default template, blank deck
    assert len(prs.slides) == 0, "new_deck did not clear slides"
    s = prs.slides.add_slide(blank_slide_layout(prs))

    # force the hard path: no notes body placeholder, then write a multi-line note
    _strip_notes_placeholders(s)
    assert s.notes_slide.notes_text_frame is None, "expected a placeholder-less notes master"
    NOTE = "line one\nline two with an em-dash — and 한글"
    speaker_note(s, NOTE)

    # overflow helper sanity: a shape pushed off the slide is reported
    from pptx.util import Inches
    s.shapes.add_textbox(Inches(12.0), Inches(7.0), Inches(3), Inches(2))  # runs off 13.33x7.5
    assert overflows(s), "overflows() failed to flag an off-slide shape"

    # hang(): sets marL/indent on the paragraph's XML (python-pptx has no property for this)
    box = s.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(3), Inches(1))
    p = box.text_frame.paragraphs[0]
    p.text = "wraps under the bullet marker, not back to the margin"
    hang(p, 137160)  # ~0.15in in EMU
    assert p._pPr is not None and p._pPr.get("marL") == "137160" and p._pPr.get("indent") == "-137160", \
        "hang() did not set marL/indent as expected"

    with tempfile.TemporaryDirectory() as d:
        out = os.path.join(d, "selftest.pptx")
        prs.save(out)
        r = Presentation(out)
        got = r.slides[0].notes_slide.notes_text_frame.text
        assert got == NOTE, f"note round-trip mismatch:\n  wrote={NOTE!r}\n  read ={got!r}"

    print("PASS  speaker_note round-trips on a placeholder-less notes master; "
          "overflows() and hang() work.")

def test_dtable():
    """dtable() 이 «진짜 표»를 만들고, 학술 서식(세로선 없음)을 실제로 강제하는가.

    XML 을 손으로 쓰는 함수라 조용히 깨지기 쉽다 — a:tcPr 자식의 스키마 순서를 어기면
    PowerPoint 가 복구 모드로 열고, noFill 을 빠뜨리면 템플릿 표 스타일의 세로선이 비쳐 나온다.
    """
    from lxml import etree
    from pptx.dml.color import RGBColor
    from pptx_kit import dtable, _q, _TC_ORDER

    prs = Presentation()
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rows = [["Variable", "n", "OR (95% CI)"],
            ["A cell long enough that it must wrap at this narrow column width", "12",
             "1.23 (0.98-1.56)"],
            ["**Total", "34", "**2.10 (1.4-3.1)"]]
    h = dtable(sl, rows, 0.5, 0.5, 6.0, [3, 1, 2], RGBColor(0, 0, 0), size_pt=12,
               name="pk:selftest")

    tbl = [s for s in sl.shapes if s.has_table][0].table
    assert len(tbl.rows) == 3 and len(tbl.columns) == 3, "표 크기가 다르다"

    # 줄바꿈되는 행은 «다른 행보다 높아야» 한다 — 고정 높이면 그 셀이 아래 행을 덮는다
    hs = [r.height for r in tbl.rows]
    assert hs[1] > hs[0], "줄바꿈되는 행의 높이를 재지 않았다 (고정 높이로 되돌아갔다)"
    assert abs(h - sum(hs) / 914400.0) < 1e-6, "반환 높이가 행 합과 다르다"

    for i in range(3):
        for j in range(3):
            tcPr = tbl.cell(i, j)._tc.find(_q("tcPr"))
            assert tcPr is not None, "tcPr 가 없다"
            names = [etree.QName(c).localname for c in tcPr]
            known = [n for n in names if n in _TC_ORDER]
            assert known == sorted(known, key=_TC_ORDER.index), \
                "a:tcPr 자식 순서가 스키마와 다르다 — PowerPoint 가 복구 모드로 연다: %s" % names
            # 세로선은 «명시적으로» 꺼야 한다. 비워 두면 템플릿 스타일이 비쳐 나온다.
            for side in ("lnL", "lnR"):
                ln = tcPr.find(_q(side))
                assert ln is not None and ln.find(_q("noFill")) is not None, \
                    "세로선을 명시적으로 끄지 않았다 (%s, 셀 %d,%d)" % (side, i, j)
    assert tbl.cell(0, 0)._tc.find(_q("tcPr")).find(_q("lnT")).find(_q("solidFill")) is not None, \
        "표 상단 선이 없다"
    assert tbl.cell(2, 0)._tc.find(_q("tcPr")).find(_q("lnB")).find(_q("solidFill")) is not None, \
        "표 하단 선이 없다"
    print("PASS  dtable(): 진짜 표 · 행 높이를 잼 · 세로선 명시적 차단 · tcPr 스키마 순서 유지.")


def test_tag_korean_runs():
    """한글 run 에만, 도형·표 칸·그룹 안까지 lang="ko-KR" 이 붙는가. save_and_check 도 붙이는가.

    태그가 빠지면 PowerPoint 가 한글을 음절 중간에서 줄바꿈한다(guide §5) — 렌더로만 보이는
    결함이라 여기서 XML 로 지킨다."""
    from pptx.util import Inches
    from pptx_kit import tag_korean_runs

    def _lang(run):
        rPr = run._r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
        return None if rPr is None else rPr.get("lang")

    prs = new_deck()
    sl = prs.slides.add_slide(blank_slide_layout(prs))
    p = sl.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(4), Inches(1)).text_frame.paragraphs[0]
    r_ko = p.add_run(); r_ko.text = "표본 크기 산정 근거"
    r_en = p.add_run(); r_en.text = " n = 120 "
    tbl = sl.shapes.add_table(1, 2, Inches(0.5), Inches(2), Inches(4), Inches(0.5)).table
    tbl.cell(0, 0).text = "120명"; tbl.cell(0, 1).text = "n"
    grp = sl.shapes.add_group_shape()
    grp.shapes.add_textbox(Inches(5), Inches(0.5), Inches(3), Inches(1)).text_frame.text = "그룹 안 한글"

    n = tag_korean_runs(prs)
    assert n == 3, f"한글 run 3개(본문·표 칸·그룹)여야 한다, got {n}"
    assert _lang(r_ko) == "ko-KR" and _lang(r_en) is None, "한글 run 에만 붙어야 한다"
    assert _lang(tbl.cell(0, 0).text_frame.paragraphs[0].runs[0]) == "ko-KR", "표 칸이 빠졌다"
    assert _lang(tbl.cell(0, 1).text_frame.paragraphs[0].runs[0]) is None
    g_run = list(grp.shapes)[0].text_frame.paragraphs[0].runs[0]
    assert _lang(g_run) == "ko-KR", "그룹 안 도형이 빠졌다"

    prs2 = new_deck()
    s2 = prs2.slides.add_slide(blank_slide_layout(prs2))
    s2.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(4), Inches(1)).text_frame.text = "저장 게이트"
    with tempfile.TemporaryDirectory() as d:
        out = os.path.join(d, "ko.pptx")
        save_and_check(prs2, out)
        # 기본 템플릿엔 빈 레이아웃이 없어 제목 placeholder 가 먼저 온다 — 글상자를 텍스트로 찾는다
        box = [sh for sh in Presentation(out).slides[0].shapes
               if sh.has_text_frame and sh.text_frame.text == "저장 게이트"][0]
        run = box.text_frame.paragraphs[0].runs[0]
        assert _lang(run) == "ko-KR", "save_and_check 가 태그를 붙이지 않았다"
    print("PASS  tag_korean_runs(): 한글 run 에만 · 표 칸·그룹 포함 · save_and_check 경유 저장본에 남음.")


if __name__ == "__main__":
    main()
    test_text_and_leak_mechanics()
    test_dtable()
    test_tag_korean_runs()
