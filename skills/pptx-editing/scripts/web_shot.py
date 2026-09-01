# -*- coding: utf-8 -*-
"""웹 페이지를 «슬라이드에 놓을 크기에 맞춰» 캡처한다 (headless Chrome).

## 왜 이 도구가 스킬 안에 있나

수업자료·발표덱에 웹 화면(다운로드 페이지, 대시보드, 문서)을 넣는 일이 잦다. 그냥 찍으면
대개 흐리거나 여백투성이인데, **원인은 캡처가 아니라 «슬라이드에서 얼마나 넓게 놓느냐»** 다.

    실효 해상도(DPI) = 잘라낸 CSS 폭 x 배율 / 슬라이드에 놓인 인치

세 값이 함께 정해져야 하므로, 셋 중 둘을 주면 나머지를 «계산해서» 찍는다. 눈대중으로
배율을 정하면 어떤 그림은 262 DPI 이고 어떤 그림은 121 DPI 가 된다(실측 사고).
찍은 뒤에는 `audit_image_dpi.py` 로 덱 전체를 다시 확인한다.

## 설계에서 지키는 것 (전부 실제로 한 번씩 당한 것)

1. **CSS 뷰포트 폭을 고정한다.** 그래야 레이아웃이 매번 같고, 브라우저에서
   `getBoundingClientRect` 로 «잰» 좌표를 그대로 쓸 수 있다. 배율을 올려도 레이아웃은
   그대로이므로 좌표는 유효하다 — 배율은 «래스터 밀도»만 바꾼다.
2. **좌표는 CSS px 로 받는다** (배율을 곱하는 것은 이 도구가 한다). 배율을 바꿀 때마다
   좌표를 다시 재야 한다면 그 좌표는 틀린 단위로 저장된 것이다.
3. **찍힌 폭이 기대와 다르면 멈춘다.** 배율만 올리고 원본을 다시 안 찍으면, CSS 좌표에
   큰 배율을 곱해 «엉뚱한 데»를 자른다 — 조용히.
4. **강조 상자는 글자 «위»에 얹지 않는다.** 대상을 두르기만 한다. 번호 배지로 링크
   문구를 가려 버린 사고가 있었다.

## 쓰는 법

    # 배율을 직접 준다
    python web_shot.py https://example.com/ out.png --viewport 1707 --scale 3

    # 슬라이드 배치 폭에서 «배율을 계산»하게 한다 (권장)
    python web_shot.py https://example.com/ out.png \
        --crop 0,300,462,398 --slide-width 7.6 --min-dpi 200

    # 강조 상자(들)를 두른다 — CSS px 로 l,t,w,h
    python web_shot.py URL out.png --crop 0,8,1150,156 --box 21,66,280,33

    python web_shot.py --selftest

⚠ Chrome(또는 Edge)이 필요하다. 없으면 어디를 찾았는지 말하고 멈춘다.
⚠ 이 도구는 «무엇이 찍혔는가» 는 모른다. 페이지가 개편되면 좌표가 조용히 어긋나므로
  찍은 PNG 를 반드시 눈으로 확인한다 — 그것이 좌표의 유일한 검증이다.
"""
import argparse
import os
import subprocess
import sys
import tempfile

from PIL import Image, ImageDraw

BOX_RGB = (0xDC, 0x32, 0x32)

CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "/usr/bin/google-chrome", "/usr/bin/chromium", "/usr/bin/chromium-browser",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
]


def find_chrome(explicit=None):
    if explicit:
        if not os.path.exists(explicit):
            raise SystemExit("지정한 브라우저가 없습니다: %s" % explicit)
        return explicit
    for c in CHROME_CANDIDATES:
        if os.path.exists(c):
            return c
    raise SystemExit("Chrome/Edge 를 찾지 못했습니다. --chrome 으로 경로를 주세요.\n  찾아본 곳:\n    "
                     + "\n    ".join(CHROME_CANDIDATES))


def scale_for(crop_w_css, slide_width_in, min_dpi):
    """놓일 폭과 목표 DPI 에서 «필요한 배율»을 올림으로 정한다."""
    if crop_w_css <= 0:
        raise SystemExit("crop 폭이 0 이하입니다.")
    need = min_dpi * slide_width_in / float(crop_w_css)
    return max(1, int(need) + (1 if need > int(need) else 0))


def shoot(url, out, viewport=(1707, 980), scale=2, chrome=None, budget_ms=12000):
    """페이지를 통째로 찍어 out 에 저장하고 (w, h) 를 돌려준다."""
    exe = find_chrome(chrome)
    with tempfile.TemporaryDirectory() as prof:
        cmd = [exe, "--headless=new", "--disable-gpu", "--hide-scrollbars",
               "--no-first-run", "--no-default-browser-check",
               "--default-background-color=ffffff",
               "--force-device-scale-factor=%d" % scale,
               "--window-size=%d,%d" % viewport,
               "--user-data-dir=" + prof,
               "--virtual-time-budget=%d" % budget_ms,
               "--screenshot=" + out, url]
        subprocess.run(cmd, check=False, capture_output=True, timeout=180)
    if not os.path.exists(out):
        raise SystemExit("촬영 실패: %s" % url)
    w, h = Image.open(out).size
    # ③ 기대한 배율로 찍혔는지 — 아니면 아래의 CSS 좌표 계산이 전부 어긋난다
    if abs(w - viewport[0] * scale) > 2 * scale:
        raise SystemExit("배율 x%d 로 찍히지 않았습니다 (폭 %d, 기대 %d)."
                         % (scale, w, viewport[0] * scale))
    return w, h


def annotate(path, out, scale, crop=None, boxes=(), stack=None, gap=0):
    """CSS px 좌표로 상자를 두르고 잘라낸다. stack 을 주면 조각들을 세로로 붙인다."""
    img = Image.open(path).convert("RGB")
    d = ImageDraw.Draw(img)
    lw = max(2, round(3 * scale))
    for (bl, bt, bw, bh) in boxes:
        d.rectangle([bl * scale, bt * scale, (bl + bw) * scale, (bt + bh) * scale],
                    outline=BOX_RGB, width=lw)
    if stack:
        parts = [img.crop((l * scale, t * scale, r * scale, b * scale))
                 for (l, t, r, b) in stack]
        g = gap * scale
        W = max(p.width for p in parts)
        H = sum(p.height for p in parts) + g * (len(parts) - 1)
        res = Image.new("RGB", (W, H), "white")
        y = 0
        for p in parts:
            res.paste(p, (0, y))
            y += p.height + g
    elif crop:
        l, t, r, b = crop
        res = img.crop((l * scale, t * scale, r * scale, b * scale))
    else:
        res = img
    res.save(out, dpi=(150 * scale, 150 * scale))
    return res.size


def _quad(s, n=4):
    v = [int(x) for x in s.split(",")]
    if len(v) != n:
        raise argparse.ArgumentTypeError("값 %d개가 필요합니다: %s" % (n, s))
    return tuple(v)


# ------------------------------------------------------------------ 자체 시험
def _selftest():
    """망 없이 확인한다 — 로컬 html 을 찍고, 배율·좌표·상자가 맞는지 본다."""
    tmp = tempfile.mkdtemp()
    html = os.path.join(tmp, "t.html")
    with open(html, "w", encoding="utf-8") as f:
        f.write("<body style='margin:0'>"
                "<div style='width:400px;height:200px;background:#000'></div>"
                "</body>")
    url = "file:///" + html.replace("\\", "/")
    fails = []

    # ① 배율 계산 — 순수 함수라 브라우저 없이도 검사한다
    if scale_for(462, 7.6, 200) != 4:       # 200*7.6/462 = 3.29 -> 4
        fails.append("scale_for(462, 7.6, 200) 은 4 여야 한다: %d"
                     % scale_for(462, 7.6, 200))
    if scale_for(1600, 12.2, 200) != 2:     # 200*12.2/1600 = 1.53 -> 2
        fails.append("scale_for(1600, 12.2, 200) 은 2 여야 한다")

    try:
        raw = os.path.join(tmp, "raw.png")
        w, h = shoot(url, raw, viewport=(800, 400), scale=2)
        if (w, h) != (1600, 800):
            fails.append("배율 2 로 800x400 을 찍으면 1600x800 이어야 한다: %dx%d" % (w, h))
        # ② CSS 좌표로 자른다 — 검은 블록만 나와야 한다
        cut = os.path.join(tmp, "cut.png")
        size = annotate(raw, cut, 2, crop=(0, 0, 400, 200))
        if size != (800, 400):
            fails.append("CSS 400x200 을 배율 2 로 자르면 800x400: %s" % (size,))
        px = Image.open(cut).convert("L").getpixel((400, 200))
        if px > 40:
            fails.append("잘라낸 자리가 검은 블록이어야 한다 (밝기 %d)" % px)
        # ③ 배율이 안 맞으면 «멈춰야» 한다
        try:
            shoot(url, os.path.join(tmp, "x.png"), viewport=(800, 400), scale=99)
            fails.append("배율 99 로는 찍히지 않으므로 멈춰야 한다")
        except SystemExit:
            pass
    except SystemExit as e:
        if "찾지 못했습니다" in str(e):
            print("  (Chrome 이 없어 브라우저 부분은 건너뜀 — 순수 계산만 검사)")
        else:
            fails.append(str(e))

    print("\n" + "=" * 60)
    if fails:
        for f in fails:
            print("  ✗ " + f)
        raise SystemExit("selftest 실패 %d건" % len(fails))
    print("selftest 통과 — 배율 계산 · CSS 좌표 자르기 · 배율 불일치 감지")
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="슬라이드용 웹 페이지 캡처")
    ap.add_argument("url", nargs="?")
    ap.add_argument("out", nargs="?")
    ap.add_argument("--viewport", type=int, default=1707, help="CSS 뷰포트 폭 (기본 1707)")
    ap.add_argument("--height", type=int, default=980, help="창 높이 (기본 980)")
    ap.add_argument("--scale", type=int, help="device scale factor. 안 주면 아래에서 계산")
    ap.add_argument("--slide-width", type=float, help="슬라이드에 놓을 폭(인치)")
    ap.add_argument("--min-dpi", type=int, default=200, help="목표 실효 DPI (기본 200)")
    ap.add_argument("--crop", type=lambda s: _quad(s), help="l,t,r,b (CSS px)")
    ap.add_argument("--box", action="append", default=[], type=lambda s: _quad(s),
                    help="강조 상자 l,t,w,h (CSS px). 여러 번 줄 수 있다")
    ap.add_argument("--chrome", help="브라우저 실행 파일 경로")
    ap.add_argument("--keep-raw", help="자르기 전 원본을 이 경로에 남긴다")
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args()
    if a.selftest:
        sys.exit(_selftest())
    if not a.url or not a.out:
        ap.error("url 과 out 이 필요합니다 (또는 --selftest)")

    crop_w = (a.crop[2] - a.crop[0]) if a.crop else a.viewport
    if a.scale:
        scale = a.scale
    elif a.slide_width:
        scale = scale_for(crop_w, a.slide_width, a.min_dpi)
        print("배율 x%d 로 정했습니다 — CSS %dpx 를 %.2fin 에 놓으면 %d DPI"
              % (scale, crop_w, a.slide_width, crop_w * scale / a.slide_width))
    else:
        scale = 2
        print("배율 x2 (기본). --slide-width 를 주면 필요한 배율을 계산합니다.")

    raw = a.keep_raw or os.path.join(tempfile.mkdtemp(), "raw.png")
    w, h = shoot(a.url, raw, viewport=(a.viewport, a.height), scale=scale,
                 chrome=a.chrome)
    print("원본 %dx%d" % (w, h))
    size = annotate(raw, a.out, scale, crop=a.crop, boxes=a.box)
    print("저장 %s  %dx%d" % (a.out, size[0], size[1]))
    if a.slide_width:
        print("→ %.2fin 에 놓으면 %d DPI" % (a.slide_width, size[0] / a.slide_width))
    print("⚠ 좌표의 유일한 검증은 «눈으로 보는 것»입니다. 결과 PNG 를 여세요.")
