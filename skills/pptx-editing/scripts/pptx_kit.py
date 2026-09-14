# -*- coding: utf-8 -*-
"""pptx_kit — palette-agnostic PowerPoint MECHANICS, shared across the lab's deck builders.

Single home for the fiddly, template-dependent, easy-to-get-wrong parts of building a deck with
python-pptx. Project-specific STYLE (palette, fonts, card/section components) lives in each project's
own kit_common.py / ppt_common.py, which import the mechanics from here.

Why: speaker notes, template loading + slide clearing, blank-layout lookup, and overflow-safe image
placement are identical across every deck and each has a non-obvious failure mode. Duplicating them
per project is how they drift (e.g. one project "discovers" a template has no notes placeholder and
concludes notes are impossible — the fix is to inject the placeholder; see speaker_note).

Nothing here references a colour or font: callers pass those in. Keep it that way so it stays shared.
Import path: this file lives in the `pptx-editing` skill's scripts/; add that dir to sys.path.
"""
# 표 폭은 «한 계산»을 셋이 나눠 쓴다 — docx·pptx·hwpx.
# 계산은 같은 폴더의 col_widths.py — 정본은 docx-editing 이고 여기 «복사본»을 둔다
# (스킬 하나만 받아도 동작해야 한다: ops3 공개 조건 1 자족성)
import os as _os, sys as _sys, re as _re
_CW = _os.path.dirname(_os.path.abspath(__file__))
if _CW not in _sys.path:
    _sys.path.insert(0, _CW)
from col_widths import content_col_widths  # noqa: E402
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE


def new_deck(template=None, width_in=13.333, height_in=7.5):
    """Load a template (inheriting its theme/master) and clear every slide, returning the deck.

    Removing the template's own slides while keeping its masters is the standard 'use the theme, build
    fresh' idiom, done by hand because python-pptx has no delete-slide API. slide_width/height are set;
    the caller tracks them if its styled components need SW/SH."""
    prs = Presentation(template) if template else Presentation()
    if not template:
        prs.slide_width = Inches(width_in); prs.slide_height = Inches(height_in)
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        rId = (sid.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
               or sid.get("r:id"))
        if rId:
            try: prs.part.drop_rel(rId)
            except KeyError: pass
        lst.remove(sid)
    return prs


def blank_slide_layout(prs):
    """The blank layout. Templates differ: some expose a single placeholder-free 'DEFAULT' layout
    rather than the canonical index 6, so pick the first layout with no placeholders."""
    for lay in prs.slide_layouts:
        if len(lay.placeholders) == 0:
            return lay
    return prs.slide_layouts[0]


def rect(slide, l, t, w, h, color):
    """A filled, borderless rectangle (bars, panels, backgrounds). Coordinates in inches."""
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    r.fill.solid(); r.fill.fore_color.rgb = color; r.line.fill.background()
    return r


def slide_number(prs, slide, color, font, size=11, l=12.3, t=6.98, w=0.85, h=0.32):
    """Bottom-right slide number = current slide count. Colour and font are passed by the caller."""
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    p = box.text_frame.paragraphs[0]; p.text = str(len(prs.slides))
    p.alignment = PP_ALIGN.RIGHT; p.font.size = Pt(size)
    p.font.color.rgb = color; p.font.name = font
    return box


def fit_picture(slide, path, left, top, max_w, max_h, center_x=True):
    """Add an image fitted inside a (max_w x max_h) box preserving aspect ratio, measured from the
    real pixel dimensions (PIL). Returns the Picture, or None if the file is missing (caller decides
    how to signal that). Overflow-safe: the image never exceeds the box. Coordinates in inches."""
    import os
    from PIL import Image
    if not os.path.exists(path):
        return None
    iw, ih = Image.open(path).size
    ar = iw / ih
    w = max_w; h = w / ar
    if h > max_h:
        h = max_h; w = h * ar
    x = left + (max_w - w) / 2 if center_x else left
    return slide.shapes.add_picture(path, Inches(x), Inches(top), Inches(w), Inches(h))


def overflows(slide, sw=13.333, sh=7.5, tol=0.02):
    """Return [(name, right_in, bottom_in), ...] for shapes whose right/bottom leaves the slide.
    Run before delivering: the list must be empty. python-pptx will not tell you otherwise."""
    out = []
    for sh_ in slide.shapes:
        try:
            r = (sh_.left + sh_.width) / 914400.0
            b = (sh_.top + sh_.height) / 914400.0
        except TypeError:
            continue  # some shapes (e.g. placeholders) may lack explicit geometry
        if r > sw + tol or b > sh + tol:
            out.append((sh_.name, round(r, 2), round(b, 2)))
    return out


def speaker_note(slide, text):
    """Attach a PowerPoint speaker note that SURVIVES a rebuild (it is written from the build script).

    Failure mode this guards against: some templates ship a notes master with no body placeholder, so
    python-pptx returns notes_text_frame is None and `slide.notes_slide.notes_text_frame.text = ...`
    raises. That is NOT 'this template can't hold notes' — PowerPoint shows the note fine; the
    placeholder is simply absent from the master. We inject a body placeholder into the notes slide's
    spTree directly, then the note renders normally. Newlines become separate paragraphs.

    A note written this way (from the script) is regenerated on every rebuild; a note typed into the
    .pptx by hand is wiped by the next rebuild. Put speaker notes in the script, not in PowerPoint."""
    from lxml import etree
    from xml.sax.saxutils import escape
    ns = slide.notes_slide
    tf = ns.notes_text_frame
    if tf is not None:
        tf.text = text
        return
    paras = "".join("<a:p><a:r><a:t>%s</a:t></a:r></a:p>" % escape(ln) for ln in text.split("\n"))
    xml = ('<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
           '<p:nvSpPr><p:cNvPr id="10" name="Notes Placeholder"/>'
           '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
           '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
           '<p:spPr/><p:txBody><a:bodyPr/>' + paras + '</p:txBody></p:sp>')
    ns._element.spTree.append(etree.fromstring(xml))


# ---------------------------------------------------------------------------
# Native OOXML Math (mc:AlternateContent) — building NEW equations from scratch
# ---------------------------------------------------------------------------
# python-pptx has no math API at all, and (verified 2026-08-05 by introspecting the
# generated typelib module) common PowerPoint COM automation does NOT either: TextRange2
# exposes a `MathZones(Start, Length)` METHOD that returns an existing math zone as a
# TextRange2 — it reads, it does not create. There is no OMaths/MathZones.Add, no
# BuildUp, no LinearFormat setter anywhere in the typelib. Don't spend time on COM for
# this; SendKeys-driving the UI is the only COM-adjacent option and it's not worth it.
#
# What DOES work: place an empty textbox placeholder while building the deck, then AFTER
# prs.save() reopen the .pptx as a zip and string-replace that placeholder's <p:sp> with a
# real <mc:AlternateContent> block (Choice = native m:oMath, Fallback = plain text — the
# Fallback does NOT need to be a rendered image; PowerPoint only shows it to viewers that
# don't support the a14 math extension, so plain text is a perfectly valid fallback and
# skips the "render an image" problem entirely).
_MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_MATH_RPR = ('<a:rPr><a:latin typeface="Cambria Math" panose="02040503050406030204" '
             'pitchFamily="18" charset="0"/></a:rPr>')
_MATH_CTRLPR = (f'<m:ctrlPr><a:rPr i="1"><a:latin typeface="Cambria Math" '
                f'panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr></m:ctrlPr>')


def m_run(t):
    """A plain run inside a formula (one operator/variable/number)."""
    return f'<m:r>{_MATH_RPR}<m:t>{t}</m:t></m:r>'


def m_frac(num_xml, den_xml):
    """A fraction. num_xml/den_xml are fragments built from m_run/m_sub/m_sup/etc."""
    return f'<m:f><m:fPr>{_MATH_CTRLPR}</m:fPr><m:num>{num_xml}</m:num><m:den>{den_xml}</m:den></m:f>'


def m_sub(base_t, sub_t):
    """Subscript, e.g. n_rel."""
    return (f'<m:sSub><m:sSubPr>{_MATH_CTRLPR}</m:sSubPr>'
            f'<m:e>{m_run(base_t)}</m:e><m:sub>{m_run(sub_t)}</m:sub></m:sSub>')


def m_sup(base_xml, sup_t):
    """Superscript. Pass an m_sub(...) as base_xml to get a combined sub+superscript."""
    return (f'<m:sSup><m:sSupPr>{_MATH_CTRLPR}</m:sSupPr>'
            f'<m:e>{base_xml}</m:e><m:sup>{m_run(sup_t)}</m:sup></m:sSup>')


def m_delim(inner_xml, beg="(", end=")"):
    """Auto-sizing delimiters: (…) [ ] { } | |. The brackets grow with their content.

    Plain "(" / "[" runs stay text-height, so a squared or fractional argument pokes out of them
    and the formula reads as broken. Nearly every real formula needs this, so it belongs next to
    m_frac/m_sub/m_sup — it was written twice in separate deck builders before being lifted here."""
    return (f'<m:d><m:dPr><m:begChr m:val="{beg}"/><m:endChr m:val="{end}"/>'
            f'{_MATH_CTRLPR}</m:dPr><m:e>{inner_xml}</m:e></m:d>')


def m_nary(e_xml, chr_="∑"):
    """N-ary operator (Σ, Π, ∫, ...) with limits hidden (bare-symbol style)."""
    return (f'<m:nary><m:naryPr><m:chr m:val="{chr_}"/><m:limLoc m:val="undOvr"/>'
            f'<m:subHide m:val="on"/><m:supHide m:val="on"/>{_MATH_CTRLPR}</m:naryPr>'
            f'<m:sub/><m:sup/><m:e>{e_xml}</m:e></m:nary>')


def m_acc(base_t, chr_="̄"):
    """Accent over a single variable, e.g. x̄ (x-bar): m_acc('x'). Default chr_ is combining
    macron (U+0304, bar-above); pass another combining mark (e.g. '̂' for hat) for others."""
    return (f'<m:acc><m:accPr><m:chr m:val="{chr_}"/>{_MATH_CTRLPR}</m:accPr>'
            f'<m:e>{m_run(base_t)}</m:e></m:acc>')


def m_sqrt(e_xml):
    """Square root (degree hidden, i.e. a plain √, not an nth-root)."""
    return (f'<m:rad><m:radPr><m:degHide m:val="on"/>{_MATH_CTRLPR}</m:radPr>'
            f'<m:deg/><m:e>{e_xml}</m:e></m:rad>')


def equation_slot(slide, l, t, w, h, marker):
    """Empty placeholder textbox for a not-yet-real equation. promote_equations() finds it
    by this name (`EQN::marker`) after save and replaces it with the real oMath."""
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    box.name = f"EQN::{marker}"
    return box


def promote_equations(pptx_path, equations):
    """Call AFTER prs.save(pptx_path). equations = {marker: omath_run_xml, ...} where
    omath_run_xml is the m_run/m_frac/... fragments concatenated (no outer <m:oMath> tag —
    this function adds it). Rewrites each EQN::marker placeholder in the saved file into a
    real mc:AlternateContent equation, in place, preserving its position/size. Raises if a
    marker isn't found anywhere (fail loud, not silent).

    Gotcha this avoids: a naive extraction regex `<m:oMath[^>]*>` also matches the WRAPPER
    tag `<m:oMathPara ...>` (since `[^>]*` doesn't respect tag-name boundaries), silently
    nesting an extra oMathPara/oMath pair inside your formula. If you're lifting real m:oMath
    XML out of an existing hand-edited deck to reuse verbatim (rather than building fresh with
    m_run/m_frac/...), anchor the opening tag with a lookahead so it can't swallow "Para":
    `re.search(r'<m:oMath(?=[\\s>])[^>]*>(.*)</m:oMath>', xml, re.DOTALL)`.
    """
    import zipfile
    import re
    with zipfile.ZipFile(pptx_path, "r") as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}

    slide_names = [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)]
    remaining = dict(equations)
    for slide_name in slide_names:
        xml = data[slide_name].decode("utf-8")
        changed = False
        for marker in list(remaining):
            pat = re.compile(
                r'<p:sp><p:nvSpPr><p:cNvPr id="(\d+)" name="EQN::' + re.escape(marker) +
                r'"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>.*?</p:sp>', re.DOTALL)
            m = pat.search(xml)
            if not m:
                continue
            shape_id = m.group(1)
            xfrm_m = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>', m.group(0))
            off_x, off_y, ext_cx, ext_cy = xfrm_m.groups()
            omath = f'<m:oMath xmlns:m="{_MATH_NS}">{remaining[marker]}</m:oMath>'
            choice_sp = (
                f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="EQN::{marker}"/>'
                '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
                f'<p:spPr><a:xfrm><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{ext_cx}" cy="{ext_cy}"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
                '<p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/>'
                f'<a:p><a:pPr algn="l"/><a14:m><m:oMathPara xmlns:m="{_MATH_NS}">'
                f'<m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>{omath}</m:oMathPara></a14:m>'
                '<a:endParaRPr/></a:p></p:txBody></p:sp>'
            )
            fallback_sp = (
                f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="EQN::{marker}"/>'
                '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
                f'<p:spPr><a:xfrm><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{ext_cx}" cy="{ext_cy}"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
                '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" altLang="en-US" sz="1800"/>'
                f'<a:t>[{marker}]</a:t></a:r></a:p></p:txBody></p:sp>'
            )
            new_shape = (
                '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">'
                f'<mc:Choice Requires="a14">{choice_sp}</mc:Choice>'
                f'<mc:Fallback xmlns="">{fallback_sp}</mc:Fallback></mc:AlternateContent>'
            )
            xml = xml[:m.start()] + new_shape + xml[m.end():]
            changed = True
            del remaining[marker]
        if changed:
            data[slide_name] = xml.encode("utf-8")

    if remaining:
        raise SystemExit(f"promote_equations: markers not found: {list(remaining)} — "
                          f"check equation_slot() names/slides")

    with zipfile.ZipFile(pptx_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])


def hang(paragraph, width):
    """Hanging indent of `width` (an EMU-ish int, e.g. Pt(sz)*0.95) on one paragraph.

    python-pptx exposes no API for this. Without it, a bullet whose text wraps to a second line
    reads as a separate bullet (the wrapped line falls back to the paragraph's left edge, not
    under the first line's text) -- this is what a hanging indent fixes.
    """
    pPr = paragraph._pPr if paragraph._pPr is not None else paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(width)))
    pPr.set("indent", str(int(-width)))


def text_units(s):
    """Width of one line in 'half-width' units: Hangul/Jamo count as 2, everything else as 1.

    Korean glyphs render roughly 2x the width of Latin ones in most UI fonts (Malgun Gothic
    included) -- summing raw character count under-predicts how many lines a Korean-heavy string
    will wrap to. Use this instead of len(s) when estimating wrap.
    """
    return sum(2 if ("가" <= ch <= "힣" or "ㄱ" <= ch <= "ㆎ") else 1 for ch in s)


_FONT_FILES = {
    ("arial", False): r"C:\Windows\Fonts\arial.ttf",
    ("arial", True): r"C:\Windows\Fonts\arialbd.ttf",
    ("malgun gothic", False): r"C:\Windows\Fonts\malgun.ttf",
    ("malgun gothic", True): r"C:\Windows\Fonts\malgunbd.ttf",
    ("calibri", False): r"C:\Windows\Fonts\calibri.ttf",
    ("calibri", True): r"C:\Windows\Fonts\calibrib.ttf",
    ("consolas", False): r"C:\Windows\Fonts\consola.ttf",
    ("consolas", True): r"C:\Windows\Fonts\consolab.ttf",
}

# 한글 이름으로 지정된 폰트 -- PowerPoint 에서는 «맑은 고딕» 으로 쓰지만 파일명은 malgun 이다.
# ⚠ 이 별칭이 없으면 조용히 옛 근사로 떨어진다(실측: "가나다라마바사 abc" 16pt 에서
#    2.67in vs 실제 1.99in, 34% 과대추정). 예외도 경고도 없이 «값만» 틀리므로 눈치채기 어렵다.
_FONT_ALIASES = {
    "맑은 고딕": "malgun gothic",
    "맑은고딕": "malgun gothic",
    "굴림": "gulim",
    "돋움": "dotum",
    "바탕": "batang",
}
for _ko, _en in (("굴림", "gulim.ttc"), ("돋움", "dotum.ttc"), ("바탕", "batang.ttc")):
    for _b in (False, True):
        _FONT_FILES.setdefault((_FONT_ALIASES[_ko], _b), r"C:\Windows\Fonts\%s" % _en)


def text_width_in(txt, size_pt, font="Arial", bold=False, _cache={}):
    """한 줄로 놓았을 때의 **실제 렌더 폭(inch)** — 평균 문자폭으로 «추정»하지 않는다.

    ⭐ 상자 높이를 미리 잡는 모든 계산의 바닥. 폭을 추정하면 그 오차가 그대로 «여백»에
    실리고, 거기서 「빌드 -> 넘침 -> 줄임 -> 빌드 -> 여백 -> 키움」 루프가 시작된다.
    2026-08-26 실사고: 포스터 빌드가 「여백 0.2cm」라 찍은 판의 실제 여백이 4.3cm 였고
    그날 포스터를 25회 넘게 재빌드했다. 실측으로 바꾸니 오차가 0.2cm 가 됐다.

    폰트 파일이 없으면 옛 근사(라틴 0.52em / 굵게 0.60em / 한중일 1.42em)로 떨어진다.
    """
    name = _FONT_ALIASES.get(font.strip(), font).lower()
    key = (name, bool(bold), round(size_pt, 2))
    if key not in _cache:
        try:
            from PIL import ImageFont
            path = _FONT_FILES.get((name, bool(bold)))
            _cache[key] = ImageFont.truetype(path, int(round(size_pt * 4))) if path else None
        except Exception:
            _cache[key] = None
    f = _cache[key]
    if f is None:
        em = size_pt / 72.0
        wide = 0.60 if bold else 0.52
        return sum(1.42 if ord(c) > 0x2E80 else wide for c in txt) * em
    return f.getlength(txt) / (4 * 72.0)


def wrapped_lines(txt, width_in, size_pt, font="Arial", bold=False):
    """`width_in` 폭에서 실제로 몇 줄이 되는가 — 단어 단위 줄바꿈을 시뮬레이션한다."""
    usable = max(width_in, 0.05)
    n = 0
    for seg in str(txt).split("\n"):
        cur, lines = "", 1
        for word in seg.split(" "):
            trial = word if not cur else cur + " " + word
            if not cur or text_width_in(trial, size_pt, font, bold) <= usable:
                cur = trial
            else:
                lines += 1
                cur = word
        n += lines
    return n


def wrapped_row_count(lines, wrap_units=50, width_in=None, size_pt=None,
                      font="Arial", bold=False):
    """Total rendered rows for `lines` once wrapped. Pre-size a card/textbox before drawing it.

    ⭐ **`width_in` 과 `size_pt` 를 주면 실측한다 — 그렇게 부르는 것이 기본이다.**
    옛 방식(`wrap_units` 만 주기)은 「한글=2, 나머지=1」 근사에 손으로 보정한 상수를 나누는
    것이라 레이아웃마다 다시 보정해야 했고, 그 보정을 «렌더를 보며» 하는 것이 곧 재빌드
    시행착오였다(2026-08-26). 하위호환으로 남기지만 새 코드에서 쓰지 마라.
    """
    if width_in and size_pt:
        return sum(wrapped_lines(ln, width_in, size_pt, font, bold) for ln in lines)
    return sum(1 + max(0, (text_units(ln) - 1) // wrap_units) for ln in lines)


##################################################################
#####  MEASURED-HEIGHT CACHE — 정본  #####
##################################################################
# 상자·표 높이를 «추정»하면 그 오차가 여백에 실리고 「빌드→넘침→줄임→빌드」 루프가 시작된다.
# PIL 로 폭을 재도 근사는 남는다(CJK 금칙처리·커닝을 재현할 수 없다). PowerPoint 는 정답을
# 안다 -- 그래서 추정을 더 정교하게 만드는 대신 **되먹인다**:
#
#   1회차 빌드 -- 추정으로 그리고, 도형 이름에 `pk:<키해시>` 를 새긴다
#   measure_boxes.py -- COM 1회로 실측을 캐시에 적는다
#   2회차 빌드 -- 같은 해시를 찾아 **정확한 값**을 쓴다. 오차 0.
#
# 저장되는 값의 «모양»이 대상을 구분한다:
#   float  = 텍스트 상자 높이(cm, 프레임 여백 포함)
#   list   = 표의 **행별** 높이(cm). 표는 총높이만으로는 행을 배치할 수 없다.
# 캐시는 프로젝트 밖에 살아서 다음 포스터·덱은 1회차부터 정확하다.
import hashlib as _hashlib
import io as _io
import json as _json
import os as _os

# Where the cache lives. Three tries, in order:
#   1. PPTX_KIT_CACHE          -- explicit override wins.
#   2. <skill>/../../../agent/cache  -- resolved from THIS file, so no machine path is
#      baked in. In the lab checkout that lands on the shared cache; in a standalone
#      install of this skill the directory does not exist and we fall through.
#   3. ~/.cache/pptx-editing   -- portable default, created on first write.
# ⚠ The old code hardcoded one author's absolute drive path as its fallback, so the cache
#   only ever worked on a single machine -- and its first branch pointed two levels ABOVE
#   the home directory, which exists nowhere. Silent failure: a cache miss just recomputes,
#   so nobody noticed it had never worked for anyone else.
_HERE = _os.path.dirname(_os.path.abspath(__file__))
_LAB_CACHE = _os.path.normpath(_os.path.join(_HERE, "..", "..", "..", "..", "agent", "cache"))
if _os.environ.get("PPTX_KIT_CACHE"):
    CACHE_PATH = _os.environ["PPTX_KIT_CACHE"]
elif _os.path.isdir(_LAB_CACHE):
    CACHE_PATH = _os.path.join(_LAB_CACHE, "pptx_box_heights.json")
else:
    CACHE_PATH = _os.path.join(_os.path.expanduser("~"), ".cache", "pptx-editing",
                               "pptx_box_heights.json")

_HCACHE = None
_HSTATS = {"hit": 0, "miss": 0}


def _load_cache():
    global _HCACHE
    if _HCACHE is None:
        try:
            with _io.open(CACHE_PATH, encoding="utf-8") as f:
                _HCACHE = _json.load(f)
        except Exception:
            _HCACHE = {}
    return _HCACHE


def box_key(kind, text, w_cm, size, **kw):
    """높이를 결정하는 모든 입력의 지문. 하나라도 빠지면 캐시가 «조용히» 거짓말을 한다.

    호출자는 렌더 높이에 영향을 주는 것을 **전부** kw 로 넘겨야 한다 -- 여백·간격·머리표·
    최소글자크기까지. 특히 상수를 나중에 바꿀 생각이라면 그 상수도 키에 넣어라.
    """
    payload = repr((kind, text, round(float(w_cm), 4), float(size),
                    tuple(sorted((k, v) for k, v in kw.items()))))
    return "pk:" + _hashlib.md5(payload.encode("utf-8")).hexdigest()[:12]


def cached(key):
    """실측값(float=상자 높이cm · list=표 행높이cm) 또는 캐시 미스면 None."""
    v = _load_cache().get(key)
    _HSTATS["hit" if v is not None else "miss"] += 1
    return v


def cached_height(key):
    v = cached(key)
    return v if isinstance(v, (int, float)) else None


def cached_rows(key, n=None):
    """표의 행별 실측 높이. `n` 을 주면 행 수가 맞을 때만 돌려준다(내용이 바뀌면 무효)."""
    v = cached(key)
    if not isinstance(v, list):
        return None
    if n is not None and len(v) != n:
        return None
    return list(v)


def cache_report():
    """빌드 끝에 찍는다 -- 미스가 남아 있으면 그만큼 «아직 추정»이라는 뜻이다."""
    n = _HSTATS["hit"] + _HSTATS["miss"]
    if not n:
        return "높이 캐시: 사용 안 함"
    return ("높이 캐시: %d/%d 실측 (%.0f%%)%s"
            % (_HSTATS["hit"], n, 100.0 * _HSTATS["hit"] / n,
               "" if not _HSTATS["miss"] else "  -- measure_boxes.py 로 나머지를 재고 다시 빌드"))


##################################################################
#####  TABLES — a real table, not textboxes pretending  #####
##################################################################
# 왜 이 함수가 필요한가 (2026-08-27)
# ---------------------------------
# 랩의 덱 빌더들은 표를 «진짜 표»가 아니라 **셀마다 텍스트박스 + 가로줄 도형**으로 그려 왔다
# (예: 선택교과4 `ppt_common.table_slide`). 이유는 셋이고, 둘은 정당했다:
#
#   ① python-pptx 에는 셀 «테두리» API 가 없다. 채우기는 한 줄이지만 선은 a:lnL/R/T/B XML 을
#      손으로 써야 한다. 그런데 학술 표의 표준 서식은 **선만 있고 채우기가 없는 것**이라,
#      하필 python-pptx 가 못 하는 게 우리가 늘 원하는 서식이었다.
#      (`poster_kit.ptable` 은 이 문제를 «테두리를 안 쓰고 줄무늬 채우기로» 피해 갔다.
#       포스터에선 통하지만 강의·발표 덱의 학술 표에는 안 통한다.)
#   ② 진짜 표는 템플릿의 표 스타일을 상속한다 — 원치 않는 띠 색·테마 폰트·테두리가 딸려 온다.
#   ③ 그리고 진짜 이유: **여기에 표 헬퍼가 없었다.** 그래서 덱마다 각자 만들었고, 완전한
#      통제가 가장 쉬운 방법이 텍스트박스였다. 설계가 아니라 표류다.
#
# 텍스트박스 표가 치르는 대가 (모두 실측·확인됨)
# ----------------------------------------------
#   • **셀이 줄바꿈되면 아래 행을 «밀지 않고 겹친다».** 모든 행이 고정 높이라서다. 진짜 표는
#     PowerPoint 가 그 행을 늘려 밀어낸다(음성대조 실측: 지정 1.20cm -> 실제 2.54cm, +1.34cm).
#     `ptable` 독스트링이 「한 빌드에 세 번 독립적으로 터졌다」고 적어둔 그 결함이다.
#   • **감사 도구가 못 본다.** `shape.has_table` 을 도는 수치 검증기는 가짜 표를 통째로
#     통과시킨다 — 가짜 표를 쓰면 «어떤 감사가 적용되는지»가 조용히 바뀐다.
#   • 손으로 행 하나를 넣으면 아래 전부를 다시 배치해야 한다. 진짜 표는 흐른다.
#
# 그래서 이 함수는 ①을 한 번만 제대로 풀어(테두리 XML) 나머지를 전부 되찾는다.
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
# a:tcPr 자식의 «스키마 순서». 순서를 어기면 PowerPoint 가 파일을 복구 모드로 연다.
_TC_ORDER = ("lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D", "fill",
             "headers", "extLst")


def _q(tag):
    return "{%s}%s" % (_A_NS, tag)


def _cell_borders(cell, top=None, bottom=None, left=None, right=None, color="000000"):
    """셀 테두리를 «네 변 모두» 명시한다. None = 선 없음(noFill).

    ⚠ 원하지 않는 변도 반드시 `noFill` 로 명시해야 한다. 비워 두면 템플릿 표 스타일의
    테두리가 그대로 비쳐 나온다 -- 「세로선을 안 그렸는데 세로선이 있다」의 원인이 이것이다.
    """
    from lxml import etree
    tc = cell._tc
    tcPr = tc.find(_q("tcPr"))
    if tcPr is None:
        tcPr = etree.SubElement(tc, _q("tcPr"))
        tc.remove(tcPr)
        tc.insert(len(tc), tcPr)      # tcPr 는 a:tc 의 «마지막» 자식이다
    for tag in ("lnL", "lnR", "lnT", "lnB"):
        old = tcPr.find(_q(tag))
        if old is not None:
            tcPr.remove(old)
    want = {"lnL": left, "lnR": right, "lnT": top, "lnB": bottom}
    for tag in ("lnL", "lnR", "lnT", "lnB"):
        pt = want[tag]
        ln = etree.Element(_q(tag))
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        if pt:
            ln.set("w", str(int(round(pt * 12700))))
            fill = etree.SubElement(ln, _q("solidFill"))
            etree.SubElement(fill, _q("srgbClr")).set("val", color)
        else:
            etree.SubElement(ln, _q("noFill"))
        # 스키마 순서를 지켜 끼워 넣는다
        idx = 0
        for child in tcPr:
            name = etree.QName(child).localname
            if name in _TC_ORDER and _TC_ORDER.index(name) < _TC_ORDER.index(tag):
                idx += 1
            else:
                break
        tcPr.insert(idx, ln)


def dtable(slide, rows, left_in, top_in, width_in, col_frac=None, ink=None, size_pt=14,
           font="Arial", header=True, rule_color="404040", rule_top=1.5,
           rule_head=1.0, rule_bottom=1.5, rule_row=0.0, align=None,
           row_pad_in=0.10, min_row_in=0.28, name=None):
    """학술 서식(선만, 채우기 없음)의 **진짜 표**. 실제 높이(inch)를 돌려준다.

    rows[0] 은 `header=True` 면 머리행이다. 셀 문자열이 `**` 로 시작하면 그 셀만 굵게.
    `col_frac` 은 열 폭의 «비율»(합이 1일 필요 없다). `align` 은 열별 PP_ALIGN 리스트,
    생략하면 첫 열만 왼쪽·나머지 가운데(수치표의 관례).

    행 높이는 `wrapped_lines()` 로 **재서** 정한다 -- 고정 높이로 두면 긴 셀이 행을 넘치고,
    PowerPoint 가 그 행만 늘려서 «호출자가 받은 높이»가 거짓이 된다. 그래도 최종 진실은
    `measure_boxes.py` 가 COM 으로 읽는 실제 높이다(행 자동확장은 정적 검사로 안 보인다).
    """
    if col_frac is None:
        # 폭을 안 주면 «내용량 비례»로 — docx·hwpx 와 같은 계산(2026-09-09).
        _cm = content_col_widths([[str(c) for c in r] for r in rows], 16.0,
                                 max(len(r) for r in rows))
        col_frac = [x / sum(_cm) for x in _cm]
    from pptx.util import Cm, Emu
    n_row, n_col = len(rows), len(col_frac)
    tot = float(sum(col_frac))
    col_w = [width_in * f / tot for f in col_frac]
    line_in = size_pt * 1.2 / 72.0

    def cell_h(r):
        n = 1
        for j, v in enumerate(r):
            txt = v[2:] if str(v).startswith("**") else str(v)
            # 셀 좌우 여백(기본 0.1in x2)을 빼야 «실제로» 몇 줄인지 나온다
            n = max(n, wrapped_lines(txt, max(col_w[j] - 0.2, 0.1), size_pt, font,
                                     bold=(header and r is rows[0]) or str(v).startswith("**")))
        return max(min_row_in, n * line_in + row_pad_in)

    heights = [cell_h(r) for r in rows]
    gf = slide.shapes.add_table(n_row, n_col, Inches(left_in), Inches(top_in),
                                Inches(width_in), Inches(sum(heights)))
    tbl = gf.table
    # 템플릿 표 스타일의 «띠»를 끈다. 이걸 안 끄면 채우기를 지워도 첫 행이 색을 갖는다.
    tbl.first_row = False
    tbl.horz_banding = False
    tbl.first_col = False
    tbl.vert_banding = False

    for j, w in enumerate(col_w):
        tbl.columns[j].width = Emu(int(round(w * 914400)))
    for i, h in enumerate(heights):
        tbl.rows[i].height = Emu(int(round(h * 914400)))

    for i, r in enumerate(rows):
        is_head = header and i == 0
        for j in range(n_col):
            c = tbl.cell(i, j)
            c.fill.background()                     # 채우기 없음 -- 학술 표의 기본
            c.margin_left = c.margin_right = Inches(0.10)
            c.margin_top = c.margin_bottom = Inches(0.02)
            _cell_borders(
                c, color=rule_color,
                top=(rule_top if i == 0 else (rule_row or None)),
                bottom=(rule_head if is_head else
                        (rule_bottom if i == n_row - 1 else (rule_row or None))),
                left=None, right=None)              # 세로선 없음 -- 명시적으로 끈다
            raw = str(r[j]) if j < len(r) else ""
            bold = is_head or raw.startswith("**")
            txt = raw[2:] if raw.startswith("**") else raw
            p = c.text_frame.paragraphs[0]
            p.alignment = (align[j] if align else
                           (PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER))
            run = p.add_run()
            run.text = txt
            run.font.size = Pt(size_pt)
            run.font.bold = bold
            run.font.name = font
            run.font.color.rgb = ink
    if name:
        gf.name = name
    return sum(heights)


def check_surface_leaks(prs, terms):
    """Scan every text frame in `prs` for any string in `terms` (editing traces, cross-document
    references, AI-tell phrasing -- e.g. "이전 버전", "다음과 같이", meta-commentary asides).

    Returns [(slide_no, term, snippet), ...] -- empty means clean.

    `terms` has NO default and is not itself shared: what counts as a "leak" is register- and
    project-specific (a phrase that's a tell in a lab-meeting deck may be normal prose in a teaching
    deck), and a broad shared list produces false positives (documented: "다음과 같다" flagged in a
    context where it was legitimate prose). Each project supplies its own list, informed by
    output_surface.md's register test, not this function.

    See also: `audit_surface_text.py` in this same skill runs a small register-agnostic baseline
    (English academic-writing tells: "justified", "see Table 3", "I.e.,") as a standalone script with
    its own exit-1 gate -- run it in addition to this in-process check, not instead of it. For any
    non-pptx deliverable (.tex, .md, rendered .pdf), `agent/tools/surface_leak_scan.py` generalizes
    this same no-default, caller-supplied-terms approach.
    """
    hits = []
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text
            for term in terms:
                if term in text:
                    hits.append((i + 1, term, text.strip()[:60]))
    return hits


_HANGUL_RE = _re.compile(r"[ㄱ-ㆎ가-힣]")


def _text_frames(shapes):
    """도형·표 칸·그룹 안 도형의 text_frame 을 전부 낸다."""
    from pptx.shapes.group import GroupShape
    for sh in shapes:
        if isinstance(sh, GroupShape):
            yield from _text_frames(sh.shapes)
            continue
        if sh.has_text_frame:
            yield sh.text_frame
        if getattr(sh, "has_table", False):
            for row in sh.table.rows:
                for c in row.cells:
                    yield c.text_frame


def tag_korean_runs(prs):
    """한글이 든 run 에 lang="ko-KR"(altLang="en-US") 을 붙인다. 반환 = 붙인 run 수.

    python-pptx 는 run 에 언어를 적지 않는다. 그러면 PowerPoint 가 한글을 영어 문맥으로 줄바꿈해
    「하였/다」「연구/자」처럼 음절 중간에서 끊는다. ko-KR 이 붙은 run 은 같은 상자에서 어절 단위로
    넘어간다(2026-09-14 COM 렌더 실측 — altLang 유무는 무관). latinLnBrk 로는 못 고친다: 1 로 두면
    한글은 그대로 끊기고 「120명」이 「1/20명」으로 쪼개진다. 렌더로만 보이는 결함이다.

    저장 직전에 «한 번» 부른다 — 도형·표를 만드는 함수가 여럿이라 각자 붙이게 하면 하나는 빠진다.
    한글이 없는 run 은 건드리지 않는다(영문 덱에서는 0 을 반환하고 아무것도 안 바뀐다).
    """
    n = 0
    for sl in prs.slides:
        for tf in _text_frames(sl.shapes):
            for p in tf.paragraphs:
                for r in p.runs:
                    if _HANGUL_RE.search(r.text or ""):
                        rPr = r._r.get_or_add_rPr()
                        rPr.set("lang", "ko-KR")
                        rPr.set("altLang", "en-US")
                        n += 1
    return n


def save_and_check(prs, path, leak_terms=None, sw=13.333, sh=7.5, tol=0.02):
    """Save + gate on surface leaks and boundary overflow in one call. Raises SystemExit (does not
    write a "saved" log line the caller can mistake for success) if either check fails -- a deck
    that fails either check should not reach the user un-flagged.

    Reuses `overflows()` per-slide (same tolerance as everywhere else in this module) rather than a
    second boundary check with different math, so there is one definition of "overflowing," not two
    that can disagree. Tags Korean runs first (`tag_korean_runs`) so wrapping is by word.
    """
    tag_korean_runs(prs)
    prs.save(path)
    leaks = check_surface_leaks(prs, leak_terms or [])
    overflow_hits = []
    for i, slide in enumerate(prs.slides):
        for name, r, b in overflows(slide, sw=sw, sh=sh, tol=tol):
            overflow_hits.append((i + 1, name, r, b))
    print("saved:", path, "slides=", len(prs.slides._sldIdLst))
    if leaks:
        print("[surface leaks]")
        for slide_no, term, snippet in leaks:
            print(f"  slide {slide_no}: '{term}' in \"{snippet}\"")
    if overflow_hits:
        print("[overflow]")
        for slide_no, name, r, b in overflow_hits:
            print(f"  slide {slide_no}: {name} right={r:.2f} bottom={b:.2f}")
    if leaks or overflow_hits:
        raise SystemExit(f"save_and_check failed -- leaks={len(leaks)}, overflow={len(overflow_hits)}. "
                          f"Fix and rebuild before delivering.")
    return prs
