# -*- coding: utf-8 -*-
"""poster_kit — palette-agnostic MECHANICS for large-format academic posters (python-pptx).

Same split as pptx_kit.py: this file has no colour/font opinions, callers pass those in. What lives
here is the fiddly, easy-to-get-wrong part of poster layout specifically — which is different enough
from a 16:9 slide deck (cm-scale canvas, CJK+Latin mixed text, a hard minimum legible font size read
from 1.5m away, figures whose *internal* labels must stay legible) that it doesn't belong in pptx_kit.

Extracted 2026-08-10 from `KWCS/teacher_violence_qol/R/03_poster.py`, whose 2-column English poster
was judged (by a co-author, informally, and by the user directly) clearly higher quality than a
contemporaneous 3-column poster built ad hoc for another project (`matching/brush_cog`). The
difference wasn't one trick — it was five habits every poster in the lab should now default to:

  1. **Two columns, not three.** Three columns forces small figures and short line lengths that read
     as cramped from viewing distance. Two columns give each figure real width and each line of text
     room to breathe. Use `two_col_grid()`.
  2. **One enforced minimum font size, no exceptions.** `kf()` asserts on every run — a poster is read
     from ~1.5m away, and "just this one caption a bit smaller" is how a poster ends up with a
     6pt-equivalent-effective-size line nobody proofread at that distance. Default 24pt; pass
     `min_pt=None` only for genuinely decorative micro-text (e.g. a copyright line), never for content.
  3. **Figures saved at layout size, not scaled down after.** `pic_cm()` locks width to the column and
     computes height from the real aspect ratio (PIL) — never pass a `max_h` that silently shrinks the
     image, because the point sizes *baked into the PNG* (axis labels, legends) shrink with it and fall
     below the 24pt floor invisibly. If a figure must be shorter, regenerate it at a shorter aspect
     ratio in the source plotting script — don't squeeze it here.
  4. **Narrative captions, not label captions.** `caption()` is happy to hold 2-3 sentences. "Figure 1.
     X vs Y" wastes the one place on a poster where a passerby who reads nothing else still gets the
     finding. Write what the figure *means*, the way an abstract's last sentence would.
  5. **One accent colour, used rarely.** Everything is BAR (section colour) or INK (body) or GRAY
     (captions/footnotes) except the single most important callout per section, which gets the accent
     colour — never four different tinted card-fill colours competing on one canvas. `bullets()` and
     `ptable()` both take an `accent` param for exactly this: mark ONE line or row, not several.

Pre-flight checklist before calling a poster "done" (beyond the usual pptx-editing skill steps):
  - No literal "TBD" / "to be confirmed" / placeholder text anywhere on the canvas — a poster is a
    finished artifact, not a draft; if something is genuinely unconfirmed, resolve it or drop it,
    don't ship the disclaimer (this happened for real: an author affiliation sat as "to be confirmed"
    on a rendered poster for two build cycles after the affiliation had already been confirmed
    elsewhere in the project).
  - `min_font_report()` returns empty (catches text created without going through `kf()`, e.g. hand-set
    table-cell runs).
  - `pptx_kit.overflows()` returns empty, using the poster's own sw_in/sh_in (NOT the 13.333x7.5 slide
    default — a poster canvas is tens of cm, always pass the real size).
  - Render via `render_pptx.py` and look at it at "arm's length" zoom, not full-screen — full-screen
    hides exactly the too-small-to-read-from-3-feet-away problem this kit exists to prevent.

Import path: lives in the pptx-editing skill's scripts/, alongside pptx_kit.py. Load it the same way
(see any *_poster.py in the lab for the `_load_pptx_kit`-style loader; add a second loader for this
file or extend that one to load both from the same directory).
"""
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import hashlib
import io
import json
import os
import re

DEFAULT_MIN_PT = 24

# 텍스트 상자에 더하는 여유(cm). python-pptx 기본 상하 여백 0.05in x2 = 0.254cm + 줄높이 반올림.
BOX_PAD_CM = 0.4

# ⚠ 이 값을 «추정 오차를 덮는 안전마진»으로 쓰지 마라 -- 그러면 반환 높이가 실제 글자보다
#   커져 빌드가 찍는 여백이 통째로 거짓말이 된다(2026-08-26, 최대 4cm).


def pt2cm(pt):
    return pt / 72 * 2.54


def kf(run, size, color, bold=False, italic=False, font="Arial", min_pt=DEFAULT_MIN_PT):
    """Set a run's font, enforcing the poster's minimum legible size.

    min_pt=None (or 0) opts a specific run out — reserve that for genuinely decorative text
    (a copyright line, a tiny logo caption), never for anything a reader is meant to learn from."""
    if min_pt:
        assert size >= min_pt, "font %.1fpt below poster minimum %dpt" % (size, min_pt)
    f = run.font
    f.size, f.color.rgb, f.bold, f.name, f.italic = Pt(size), color, bold, font, italic
    return run


def tb(sl, l, t, w, h, anchor=None):
    """A word-wrapped textbox at (l, t, w, h) in cm."""
    x = sl.shapes.add_textbox(Cm(l), Cm(t), Cm(w), Cm(h))
    x.text_frame.word_wrap = True
    if anchor:
        x.text_frame.vertical_anchor = anchor
    return x


def _pk():
    """옆에 있는 pptx_kit 을 불러온다 -- 폭 실측의 «정본»은 거기 하나뿐이어야 한다.

    두 킷이 각자 복사본을 들면 언젠가 갈리고, 그때 어느 쪽이 참인지 알 방법이 없다.
    """
    import importlib.util
    import os
    import sys
    if "pptx_kit" in sys.modules:
        return sys.modules["pptx_kit"]
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pptx_kit.py")
    spec = importlib.util.spec_from_file_location("pptx_kit", p)
    m = importlib.util.module_from_spec(spec)
    sys.modules["pptx_kit"] = m
    spec.loader.exec_module(m)
    return m


##################################################################
#####  MEASURED-HEIGHT CACHE  #####
##################################################################
# PIL 폭 실측으로도 상자 높이는 «근사»다 -- CJK 금칙처리·커닝·자간을 재현할 수 없어서,
# 긴 문단에서 한 줄이 갈리면 그대로 오차가 된다(2026-08-26 실측 잔차 1.4cm).
#
# 그런데 PowerPoint 는 정답을 알고 있다: `TextFrame2.TextRange.BoundHeight` 가 **실제로 그린
# 텍스트 높이**다. 그래서 추정을 더 정교하게 만드는 대신 **되먹인다**:
#
#   1회차 빌드 -- 추정으로 그리고, 각 상자 이름에 `pk:<키해시>` 를 새긴다
#   measure_boxes.py -- COM 1회로 BoundHeight 를 읽어 해시별 실측 높이를 캐시에 적는다
#   2회차 빌드 -- 같은 해시를 찾아 **정확한 높이**를 쓴다. 오차 0.
#
# 캐시는 프로젝트 밖(`agent/cache/`)에 살아서 **다음 포스터·덱은 1회차부터 정확**하다.
# 키가 높이를 결정하는 모든 것(문자열·폭·크기·간격·폰트·머리표·최소크기)을 담으므로
# 프로젝트가 달라도 같은 키면 같은 높이다.
# ⭐ 구현은 `pptx_kit` 에 «하나»만 둔다 -- 여기 사본을 들면 언젠가 갈리고, 그때 어느 쪽이
#   참인지 알 방법이 없다(폭 실측에서 이미 겪은 일이다). 여기서는 «포스터 규약»만 덧입힌다.
def box_key(kind, text, w_cm, size, **kw):
    """포스터용 지문. `BOX_PAD_CM` 을 키에 포함시키는 것이 여기의 규약이다.

    ⚠ 캐시가 담는 값이 «텍스트 높이 + 패드» 라서, 패드를 바꾸면 저장된 항목 전부가 조용히
    그 차이만큼 틀린다. 키에 넣으면 자동으로 무효화된다.
    """
    return _pk().box_key(kind, text, w_cm, size, _pad=BOX_PAD_CM, **kw)


def cached_height(key):
    """실측 높이(cm) 또는 캐시 미스면 None."""
    return _pk().cached_height(key)


def cached_rows(key, n=None):
    """표의 행별 실측 높이(cm). 행 수가 다르면 None -- 내용이 바뀐 것이다."""
    return _pk().cached_rows(key, n)


def cache_report():
    return _pk().cache_report()


CACHE_PATH = _pk().CACHE_PATH        # 정본은 pptx_kit -- 여기서는 이름만 빌려 쓴다


def text_width_cm(txt, size, font="Arial", bold=False, _cache={}):
    """한 줄로 놓았을 때의 **실제 렌더 폭(cm)**. 평균 문자폭으로 «추정»하지 않는다.

    Arial 은 가변폭이라 같은 글자 수라도 폭이 크게 다르다("iii" vs "WWW"). 평균값으로
    줄 수를 세면 상자 높이가 실제 글자보다 크거나 작게 나오고, 그 오차가 그대로 «여백»
    계산에 실린다 -- 2026-08-26 teacher 포스터에서 빌드가 「여백 0.2cm」라고 찍은 판의
    실제 여백이 4.3cm 였다. 폰트 메트릭으로 재면 그 오차가 사라진다.
    (같은 기법이 `agent/tools/audit_table_widths.py` 에 이미 있었다 -- 킷만 안 쓰고 있었다.)

    폰트 파일이 없는 환경에서는 옛 근사(라틴 0.52em / 굵게 0.60em / 한글 1.42em)로 떨어진다.
    """
    return _pk().text_width_in(txt, size, font, bold) * 2.54


def est_lines(txt, w_cm, size, char_w_factor=0.52, pad_cm=0.0):
    """Wrapped line count for auto-sizing a textbox before drawing it.

    ⭐ 2026-08-26 부터 «추정»이 아니라 **실측**이다 -- `text_width_cm()` 으로 단어를 하나씩
    붙여 보며 실제 줄바꿈을 시뮬레이션한다. `char_w_factor` 는 하위호환으로 남아 있고,
    0.56 이상이면 «굵은 글씨» 신호로만 쓴다(폭 계산에는 안 쓴다).

    char_w_factor is the average glyph width as a fraction of the em (point size). This differs
    sharply by script AND by weight — CALIBRATE, don't guess (measured via PIL ImageFont
    .getbbox() on the actual TTF, 2026-08-26 — see brush_cog SESSION_LOG that date for the
    method if recalibrating for another font):
      - Latin/English, regular (Arial):        ~0.52   (this default)
      - Latin/English, BOLD (Arial Bold):       ~0.58-0.60 — meaningfully wider per character,
        not just "a bit more". A short bold word (e.g. an 8-char button-style label) can wrap
        when an equal- or longer-length regular-weight line right next to it does not, at the
        SAME box width — this was mistaken for a box-width bug across several rebuild-render
        cycles before the actual cause (weight, not width) was found by measuring glyph width
        directly instead of iterating on renders. Pass char_w_factor≈0.6 for bold labels rather
        than reusing the regular default.
      - Hangul:                                 ~1.42   (roughly square glyphs, far fewer chars per line)
    Overestimating lines is the safe failure mode (extra whitespace, not a clipped line) — if
    unsure which way to round, round the factor UP.

    ⚠ **`pad_cm` matters and defaults to 0.0, but `tb()` in this file does NOT zero out
    python-pptx's own default text-frame margins** (0.1in left+right = ~0.508cm total, plus
    top/bottom). Calling `est_lines(text, box_w, size, factor)` with the default `pad_cm=0.0`
    will UNDER-predict wrapping for any box made with `tb()` — it will happily say "1 line"
    for text that actually wraps to 2 once real margins eat into the usable width (verified
    2026-08-26: this made a first calibration check look fine when the actual render still
    wrapped). Pass `pad_cm=0.51` (or measure the real margin if a helper other than `tb()`
    made the box) to get a prediction that matches what will actually render.

    ⭐ Call this BEFORE drawing a text box whose fit is uncertain (short labels next to a
    fixed-size image, e.g. a QR code) — check `est_lines(text, box_w, size, factor, pad_cm=0.51) == 1`
    for the width you're about to use, rather than building the whole slide and rendering to
    find out. A full rebuild+audit+render cycle costs far more than one function call, and
    "guess a width, rebuild, look, adjust" is not a substitute for checking the one number
    that actually determines wrapping.
    """
    usable = max(w_cm - pad_cm, 0.1)
    bold = char_w_factor is not None and char_w_factor >= 0.56
    n = 0
    for seg in txt.split("\n"):
        cur, lines = "", 1
        for word in seg.split(" "):
            trial = word if not cur else cur + " " + word
            if not cur or text_width_cm(trial, size, bold=bold) <= usable:
                cur = trial
            else:
                lines += 1
                cur = word
        n += lines
    return n


def two_col_grid(page_w, margin, gutter, ncol=2):
    """Return (xs, colw) for an ncol-column grid spanning [margin, page_w - margin].

    Default the poster to 2 columns (see module docstring #1) — pass ncol=3 only when content
    genuinely cannot compress into two without a wall of tiny multi-panel figures; that was the
    actual failure mode this kit was extracted to fix, so treat 3 as the exception, not the default.
    """
    colw = (page_w - 2 * margin - (ncol - 1) * gutter) / ncol
    xs = [margin + i * (colw + gutter) for i in range(ncol)]
    return xs, colw


def sectitle(sl, txt, l, t, w, color, size=42, font="Arial", upper=True, min_pt=DEFAULT_MIN_PT):
    """Section header + underline rule. Returns the y (cm) where content should start next."""
    h = pt2cm(size) * 1.2
    x = tb(sl, l, t, w, h)
    kf(x.text_frame.paragraphs[0].add_run(), size, color, bold=True, font=font, min_pt=min_pt).text \
        = txt.upper() if upper else txt
    ry = t + h * 0.94
    r = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(l), Cm(ry), Cm(w), Cm(0.2))
    r.fill.solid(); r.fill.fore_color.rgb = color; r.line.fill.background(); r.shadow.inherit = False
    return ry + 0.2 + 0.7


def _spans(txt):
    """'**bold** rest' -> [(text, is_bold), ...]. A line starting with a lone leading '**' with no
    matching close is treated as whole-line emphasis by the caller (bullets/ptable), not here."""
    out, i = [], 0
    for m in re.finditer(r"\*\*(.+?)\*\*", txt):
        if m.start() > i:
            out.append((txt[i:m.start()], False))
        out.append((m.group(1), True))
        i = m.end()
    if i < len(txt):
        out.append((txt[i:], False))
    return out or [(txt, False)]


def bullets(sl, items, l, t, w, ink, accent=None, size=27, gap=14, font="Arial", char_w_factor=0.52,
            mark="•  ", min_pt=DEFAULT_MIN_PT):
    """Auto-height bullet list. Inline `**bold**` spans; a whole item wrapped in a single leading
    `**...` (i.e. starts with '**' and has no closing pair elsewhere) is rendered entirely in
    `accent` colour+bold — use this for AT MOST one bullet per section (module docstring #5), not as
    a general emphasis tool. Returns the total height (cm) consumed."""
    plain = [s.replace("**", "") for s in items]
    key = box_key("bullets", tuple(plain), w, size, gap=gap, font=font, mark=mark,
                  min_pt=min_pt, indent=1.2)
    total = cached_height(key)
    if total is None:
        total = sum(est_lines(s, w - 1.2, size, char_w_factor, pad_cm=0.51) for s in plain) \
            * pt2cm(size) * 1.2 + len(items) * pt2cm(gap)
        # 예전엔 +1.0cm 였다. 그 여유는 est_lines 가 «추정»이던 시절의 안전마진이었고, 그만큼
        # 반환 높이가 실제 글자보다 커서 빌드가 찍는 여백이 최대 4cm 틀렸다(2026-08-26).
        # 이제 폭을 실측하므로 python-pptx 텍스트 프레임의 실제 상하 여백(0.05in x2 = 0.254cm)만
        # 덮으면 된다. 줄높이 반올림 여유를 조금 더해 0.4.
        total += BOX_PAD_CM
    x = tb(sl, l, t, w, total)
    x.name = key
    tf = x.text_frame
    for i, s in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if i:
            p.space_before = Pt(gap)
        whole = s.startswith("**") and s.count("**") == 1
        s2 = s[2:] if whole else s
        whole_color = accent if (whole and accent) else ink
        kf(p.add_run(), size, whole_color, bold=whole, font=font, min_pt=min_pt).text = mark
        for seg, bd in _spans(s2):
            col = accent if ((whole or bd) and accent) else ink
            kf(p.add_run(), size, col, bold=(whole or bd), font=font, min_pt=min_pt).text = seg
    return total


def caption(sl, txt, l, t, w, color, accent=None, size=24, font="Arial", char_w_factor=0.52,
            min_pt=DEFAULT_MIN_PT):
    """A narrative caption (module docstring #4 — write what the figure MEANS). Returns height (cm).

    Supports the same inline `**bold**` markup as bullets() (rendered bold, in `accent` colour if
    given, else just bold in `color`) — pass `accent` if the caption emphasises one clause. A caption
    with no `**` in it renders as a single plain run, unchanged from before this parameter existed.
    """
    plain = txt.replace("**", "")
    # ⚠ pad_cm 를 빠뜨리면 python-pptx 의 좌우 기본 여백(0.1in x2 = 0.508cm)을 무시하게 되어
    #   «한 줄에 더 들어간다»고 낙관한다. est_lines 독스트링이 경고하던 바로 그 함정이고,
    #   bullets() 는 지키는데 여기만 안 지키고 있었다(2026-08-26).
    key = box_key("caption", plain, w, size, font=font, min_pt=min_pt)
    h = cached_height(key)
    if h is None:
        h = est_lines(plain, w, size, char_w_factor, pad_cm=0.51) * pt2cm(size) * 1.2 + BOX_PAD_CM
    x = tb(sl, l, t, w, h)
    x.name = key
    p = x.text_frame.paragraphs[0]
    for seg, bd in _spans(txt):
        col = accent if (bd and accent) else color
        kf(p.add_run(), size, col, bold=bd, font=font, min_pt=min_pt).text = seg
    return h


def ptable(sl, rows, l, t, w, widths, header_color, alt_color, white, ink, accent=None, size=24,
           rowh=1.9, title=None, font="Arial", min_pt=DEFAULT_MIN_PT, char_w_factor=0.55):
    """Poster table. `**cell` bolds that cell; a title (rendered as a merged first row, matching the
    convention seen in published posters) goes in `title`, not as a separate textbox above the table
    — keeps the table a single self-contained object when the layout height is recalculated.

    `rowh` is a FLOOR, not the assumed height of every row. Each row's real height is computed from
    the wrapped line count of its longest cell (est_lines(), the same estimator bullets()/caption()
    already use for prose) — a row only gets taller than rowh if some cell in it actually needs to
    wrap at that column's width. The returned height is the true sum, not `rowh * nrow`.

    Why this matters (2026-08-14, brush_cog poster, hit 3 times independently in one build before
    this fix existed): a plain `rowh * nrow` silently under-predicts any row whose cell wraps — a
    long row label ("After-lunch brushing" in a narrow first column), a header cell ("CHS
    (N=79,826)" in a column sized for short data values), or a long table `title`. PowerPoint then
    auto-grows JUST that row at render time (this is real, static geometry checks can't see it —
    see deck_render_audit.py's docstring), so the table's true rendered height exceeds what this
    function told the caller, and whatever the caller places next (almost always a caption)
    overlaps the table's actual last row. No amount of nudging the caption's y-offset fixes this
    reliably, because the mismatch scales with content, not a fixed pixel amount — the fix has to
    be that the returned height is right in the first place.
    """
    ncol = len(widths)
    col_w = [ww * (w / sum(widths)) for ww in widths]
    line_h = pt2cm(size) * 1.2

    def cell_lines(val, cw):
        plain = val[2:] if val.startswith("**") else val
        return est_lines(str(plain), cw, size, char_w_factor, pad_cm=0.4)

    def row_height(rr):
        return max(rowh, max(cell_lines(v, col_w[j]) for j, v in enumerate(rr)) * line_h + 0.3)

    nrow = len(rows) + (1 if title else 0)
    # ⭐ 실측 되먹임 (2026-08-27). est_lines 는 «근사»라 셀 하나가 예상보다 한 줄 더 접히면
    #   반환 높이가 작아지고, 바로 아래 놓은 캡션이 표의 아래 선을 밟는다 -- 실제로 그랬다
    #   (brush_cog KSEPI2026: 예측 14.00cm vs 실제 14.73cm). `audit_text_fit` 은 표를 못 보고
    #   빌드 로그도 조용하다. `measure_boxes.py` 가 **행별** 실측을 캐시에 넣고, 2회차 빌드가
    #   그걸 쓴다. 표는 총높이만으로는 행을 배치할 수 없어 «리스트»로 담는다.
    key = box_key("ptable", tuple(tuple(str(c) for c in r) for r in rows), w, size,
                  widths=tuple(widths), rowh=rowh, title=str(title or ""),
                  font=font, min_pt=min_pt, cwf=char_w_factor)
    row_heights = cached_rows(key, nrow)
    if row_heights is None:
        row_heights = []
        if title:
            title_lines = est_lines(title, w, size + 2, char_w_factor, pad_cm=0.35)
            row_heights.append(max(rowh, title_lines * pt2cm(size + 2) * 1.2 + 0.3))
        row_heights += [row_height(rr) for rr in rows]
    total_h = sum(row_heights)

    gt = sl.shapes.add_table(nrow, ncol, Cm(l), Cm(t), Cm(w), Cm(total_h))
    gt.name = key
    tblx = gt.table
    for j, ww in enumerate(widths):
        tblx.columns[j].width = Cm(col_w[j])
    off = 0
    if title:
        tblx.rows[0].height = Cm(row_heights[0])
        c0 = tblx.cell(0, 0)
        c0.merge(tblx.cell(0, ncol - 1))
        c0.fill.solid(); c0.fill.fore_color.rgb = white
        c0.margin_left = c0.margin_right = Cm(0.15)
        kf(c0.text_frame.paragraphs[0].add_run(), size + 2, header_color, bold=True, font=font,
           min_pt=min_pt).text = title
        off = 1
    for i, rr in enumerate(rows):
        tblx.rows[i + off].height = Cm(row_heights[i + off])
        head = (i == 0)
        for j, val in enumerate(rr):
            c = tblx.cell(i + off, j)
            c.margin_left = c.margin_right = Cm(0.18)
            c.margin_top = c.margin_bottom = Cm(0.05)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.fill.solid()
            c.fill.fore_color.rgb = header_color if head else (alt_color if (i % 2 == 0) else white)
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            bold = head or val.startswith("**")
            val2 = val[2:] if val.startswith("**") else val
            col = white if head else (accent if (val.startswith("**") and accent) else ink)
            kf(p.add_run(), size, col, bold=bold, font=font, min_pt=min_pt).text = val2
    return total_h


def stack_box_height(w, lines, size=24, char_w_factor=0.55, pad_cm=0.5):
    """The height stack_box() would use for this (w, lines) — call this FIRST when two or more
    boxes must share one height (e.g. two lanes of a flowchart row): compute each side's natural
    height, take the max, then pass it as `min_h` to every stack_box() call in that row so they
    align without any box being under-sized for its own content."""
    line_h = pt2cm(size) * 1.2
    n_lines = sum(est_lines(txt, w - pad_cm, size, char_w_factor) for txt, _, _ in lines)
    return n_lines * line_h + pad_cm


def stack_box(sl, l, t, w, lines, fill, border=None, size=24, font="Arial",
              min_pt=DEFAULT_MIN_PT, char_w_factor=0.55, pad_cm=0.5, align=PP_ALIGN.CENTER,
              min_h=0.0):
    """A rounded-rect box auto-sized to fit `lines` (a list of (text, bold, color) tuples, one per
    line, centred and stacked top-to-bottom) — height computed from est_lines() at the box's own
    width, the same pattern ptable() now uses for cells. Returns (shape, height_cm). For several
    boxes that must share one height, see stack_box_height() first.

    Extracted 2026-08-14 from a project-local flowchart box() that was hand-sized with a guessed
    constant three separate times in the same build before someone finally computed it from the
    real line count — this is that computation, done once, shared, so the next STROBE/flow-diagram
    box in any project starts from a box that actually fits its own text instead of reinventing
    (and re-breaking) the same helper. Multi-line wrapping WITHIN one logical line is estimated via
    est_lines(); this does not merge/re-wrap the list itself, so pass already-short phrases (one
    idea per list item), not paragraphs.
    """
    h = max(min_h, stack_box_height(w, lines, size, char_w_factor, pad_cm))
    r = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(l), Cm(t), Cm(w), Cm(h))
    r.fill.solid(); r.fill.fore_color.rgb = fill
    r.adjustments[0] = 0.03
    if border:
        r.line.color.rgb = border; r.line.width = Pt(1.5)
    else:
        r.line.fill.background()
    r.shadow.inherit = False
    b = tb(sl, l + pad_cm / 2, t, w - pad_cm, h, anchor=MSO_ANCHOR.MIDDLE)
    for i, (txt, bold, color) in enumerate(lines):
        p = b.text_frame.paragraphs[0] if i == 0 else b.text_frame.add_paragraph()
        p.alignment = align
        kf(p.add_run(), size, color, bold=bold, font=font, min_pt=min_pt).text = txt
    return r, h


def pic_cm(sl, path, l, t, w_cm):
    """Insert an image at exactly `w_cm` wide, height from its real aspect ratio (PIL).

    Deliberately has NO max_h parameter (module docstring #3) — a poster figure's width should be set
    by the column it lives in, and its height should follow from that. If a figure is too tall for the
    space left on the page, the fix is regenerating the figure at a different aspect ratio (in its own
    plotting script) or reordering sections, not shrinking it here — shrinking silently drops any
    in-figure text below the poster's minimum legible size.
    """
    from PIL import Image
    iw, ih = Image.open(path).size
    h_cm = w_cm * ih / iw
    sl.shapes.add_picture(path, Cm(l), Cm(t), Cm(w_cm), Cm(h_cm))
    return h_cm


def min_font_report(slide, min_pt=DEFAULT_MIN_PT):
    """Mechanical pre-flight check: walk every run on the slide (textboxes AND table cells) and
    return [(shape_name, text_snippet, size_pt), ...] for anything under min_pt.

    Catches text created without going through kf() — e.g. a table cell font set by hand, or a run
    whose size was computed and happens to round under the floor. kf()'s assert only protects text
    that actually goes through it; this is the safety net for text that didn't. Empty return = clean.

    Narrow by design: this only re-checks the [RUN] case (a literal font size on a run). It does NOT
    catch an image placed smaller than native — baked-in labels shrink with it, 11pt drawn at 0.7x
    placement renders at 7.7pt — or an undersized cex=/size=/fontsize= literal in the R/Python script
    that generated a figure. This function has no dependency outside this file, so a standalone
    install of the public skill only gets this narrower check. If the lab's shared tooling is also
    available, run `agent/tools/audit_font_sizes.py` on the built .pptx for the fuller three-pronged
    check (RUN + SCALE + SRC) — it is not imported here, so the two can drift; if the definition of
    "too small" changes in one, check whether it should change in the other."""
    hits = []
    for shp in slide.shapes:
        runs = []
        if shp.has_text_frame:
            for p in shp.text_frame.paragraphs:
                runs.extend(p.runs)
        if shp.has_table:
            for row in shp.table.rows:
                for cell in row.cells:
                    for p in cell.text_frame.paragraphs:
                        runs.extend(p.runs)
        for r in runs:
            sz = r.font.size
            if sz is not None and sz.pt < min_pt and r.text.strip():
                hits.append((shp.name, r.text.strip()[:40], round(sz.pt, 1)))
    return hits
