# -*- coding: utf-8 -*-
"""렌더에서 «열별 실제 잉크 끝»을 잰다 — 빌드 스크립트가 찍는 여백 숫자를 믿지 않는다.

## 이 도구가 막는 사고

포스터·덱 빌드 스크립트는 보통 블록을 쌓으며 `y` 를 누적하고 마지막에 여백을 찍는다:

    print("RIGHT end y=%.1f  (margin %.1f cm)" % (y, BOTTOM - y))

**이 숫자는 실제 여백이 아니다.** 텍스트 상자 높이는 «문자폭 추정 + 고정 패딩»으로 잡혀
글자가 실제로 차지하는 높이보다 크다(`poster_kit.bullets` 는 호출마다 1.0cm 를 더한다).
그래서 두 방향으로 다 틀린다:

  · 스크립트가 «꽉 찼다»는데 렌더엔 빈 공간이 남는다 → 그림을 작게 두고 여백을 남긴다
  · 스크립트가 «넘쳤다»는데 글자는 안 넘친다 → 없는 문제를 고치려 재빌드를 반복한다

**실측(2026-08-26, teacher KSPM 포스터):** 스크립트가 우측 여백 **0.2cm** 라고 찍은 그 빌드를
이 도구로 재니 **4.3cm** 였다. 사용자가 눈으로 먼저 알아챘고("오른쪽에 공간이 좀 나는데"),
그 4.3cm 로 Figure 1 을 1.3배 키울 수 있었다. 그 세션에서 포스터를 **25회 넘게** 재빌드했는데
상당수가 이 어긋난 숫자를 좇은 시행착오였다.

## 쓰는 법

    python ink_extent.py <포스터.pptx>                  # 캔버스·띠를 스스로 알아낸다
    python ink_extent.py <렌더PNG> --canvas 90x120
    python ink_extent.py <포스터.pptx> --cols "LEFT:3.5-44.25,RIGHT:45.75-86.5"

  .pptx 를 주면 캔버스 크기를 파일에서 읽고 렌더 PNG 를 `<이름>_png/` 에서 찾는다.
  없으면 `render_pptx.py` 로 만든다.

  --canvas   PNG 를 직접 줄 때만 필요(cm).
  --cols     열 이름과 좌우 경계(cm). 생략하면 캔버스를 좌/우 반으로 나눈다.
  --bottom   내용이 끝나야 하는 y(cm). 생략하면 푸터 띠를 자동 검출, 없으면 캔버스 높이.
  --top      이 y(cm)부터 본다. 생략하면 머리글 띠를 자동 검출.

⚠️ 머리글·푸터 «띠»(폭 전체 단색 사각형)는 자동 배제된다 — 안 그러면 띠가 잉크로 잡혀
   늘 「여백 0」이 나오고 검사가 늑대소년이 된다.

## 읽는 법 (`poster_rules` §1)

  · 남은 여백 > 3cm  → 「할 말이 없다」로 읽힌다. **결과를 더 넣거나 그림을 키운다.**
    캡션을 늘려 메우지 않는다.
  · 0.2~1.5cm       → 목표 대역
  · < 0.2cm         → 인쇄물에서 답답하다
  · 두 열 차이 > 3cm → 균형이 깨진 것

⚠️ 이 도구는 «세로로 어디까지 찼나»만 본다. 겹침·이탈은 `audit_text_fit.py`,
   표 행높이 확장으로 인한 잘림은 `deck_render_audit.py` 가 본다. 축이 다르다.
"""
import argparse
import glob
import os
import subprocess
import sys

from PIL import Image

INK_MAX = 200      # 이보다 어두우면 잉크(0=검정, 255=흰색)
MIN_DARK = 3       # 한 행에 이만큼은 어두워야 «내용» — 점 하나로 오판하지 않게
STEP_X = 3         # 가로 표본 간격(px). 전수는 느리고 불필요하다
HERE = os.path.dirname(os.path.abspath(__file__))


def canvas_of(pptx):
    from pptx import Presentation
    prs = Presentation(pptx)
    return prs.slide_width / 360000.0, prs.slide_height / 360000.0


def render_png(pptx):
    """렌더 PNG 를 찾고, 없으면 render_pptx.py 로 만든다."""
    d = os.path.splitext(pptx)[0] + "_png"
    hits = sorted(glob.glob(os.path.join(d, "*.[pP][nN][gG]")))
    if not hits:
        subprocess.run([sys.executable, os.path.join(HERE, "render_pptx.py"), pptx],
                       check=True, capture_output=True)
        hits = sorted(glob.glob(os.path.join(d, "*.[pP][nN][gG]")))
    if not hits:
        raise SystemExit("렌더 PNG 를 만들지 못했다: " + pptx)
    return hits[0]


def resolve_target(target, canvas):
    if target.lower().endswith(".pptx"):
        cw, ch = canvas_of(target)
        return render_png(target), cw, ch
    if not canvas:
        raise SystemExit("PNG 를 직접 줄 때는 --canvas 90x120 처럼 캔버스 크기가 필요하다")
    a, _, b = canvas.partition("x")
    return target, float(a), float(b)


def _modal(vals):
    """행의 최빈값 -- 띠 위에 «글자»가 있어도 배경색이 최빈값으로 남는다."""
    c = {}
    for v in vals:
        k = v // 5                       # 5단계로 뭉쳐 안티에일리어싱을 흡수
        c[k] = c.get(k, 0) + 1
    return max(c, key=c.get) * 5 + 2


def find_band(im, from_bottom, px_cm, min_cm=0.5):
    """가장자리의 «색 띠»의 안쪽 경계(px). 없으면 None.

    ⚠ 균일한 행만 세면 안 된다 -- 푸터 띠에는 보통 글자가 있어서 그 행에서 끊긴다.
      2026-08-26 첫 판이 그래서 푸터 «글자»를 잉크로 세고 「여백 0」이라는 오탐을 냈다.
      행의 «최빈값»으로 보면 글자가 있어도 배경색이 남으므로 띠 전체를 제대로 잡는다.
    """
    W, H = im.size
    p = im.load()
    xs = list(range(0, W, STEP_X))
    rows = range(H - 1, -1, -1) if from_bottom else range(H)
    band, run = None, 0
    for y in rows:
        m = _modal([p[x, y] for x in xs])
        if band is None:
            if m > 245:                  # 흰 배경에서 시작하면 띠가 없다
                return None
            band, run = m, 1
            continue
        if abs(m - band) <= 6:
            run += 1
        else:
            break
    if run / px_cm < min_cm:
        return None
    return (H - run) if from_bottom else run


def parse_cols(spec, canvas_w):
    if not spec:
        half = canvas_w / 2.0
        return [("LEFT", 0.0, half), ("RIGHT", half, canvas_w)]
    out = []
    for part in spec.split(","):
        name, _, rng = part.partition(":")
        a, _, b = rng.partition("-")
        out.append((name.strip(), float(a), float(b)))
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("target", help="포스터 .pptx 또는 렌더 PNG")
    ap.add_argument("--canvas", default=None)
    ap.add_argument("--cols", default=None)
    ap.add_argument("--bottom", type=float, default=None)
    ap.add_argument("--top", type=float, default=None)
    a = ap.parse_args(argv)

    png, cw, ch = resolve_target(a.target, a.canvas)
    im = Image.open(png).convert("L")
    W, H = im.size
    px_y, px_x = H / ch, W / cw

    bot_px = find_band(im, True, px_y)
    top_px = find_band(im, False, px_y)
    bottom = a.bottom if a.bottom is not None else (bot_px / px_y if bot_px else ch)
    top = a.top if a.top is not None else (top_px / px_y if top_px else 0.0)
    y_lo, y_hi = int(top * px_y), min(H, int(bottom * px_y))
    p = im.load()

    print("%s  %dx%d px  캔버스 %.0fx%.0f cm" % (os.path.basename(png), W, H, cw, ch))
    print("  검사 구간 %.1f-%.1f cm%s"
          % (top, bottom, "  (머리글·푸터 띠 자동 배제)" if (bot_px or top_px) else ""))

    gaps = []
    for name, x0, x1 in parse_cols(a.cols, cw):
        xa, xb = max(0, int(x0 * px_x)), min(W, int(x1 * px_x))
        last = None
        for y in range(y_hi - 1, y_lo, -1):
            if sum(1 for x in range(xa, xb, STEP_X) if p[x, y] < INK_MAX) >= MIN_DARK:
                last = y
                break
        if last is None:
            print("  %-6s 잉크 없음" % name)
            continue
        cm = last / px_y
        gap = bottom - cm
        gaps.append(gap)
        flag = "빈 공간 과다" if gap > 3 else ("빡빡함" if gap < 0.2 else "목표 대역")
        print("  %-6s 잉크 끝 %6.1f cm   남은 여백 %5.1f cm   [%s]" % (name, cm, gap, flag))

    if len(gaps) > 1 and (max(gaps) - min(gaps)) > 3:
        print("  ** 두 열의 여백 차이가 %.1f cm — 균형이 깨졌다" % (max(gaps) - min(gaps)))
    return 1 if gaps and min(gaps) < 0 else 0


if __name__ == "__main__":
    sys.exit(main())
