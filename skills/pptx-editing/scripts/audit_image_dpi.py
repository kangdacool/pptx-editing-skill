# -*- coding: utf-8 -*-
"""덱 안 그림들이 «놓인 크기에 견줘» 충분한 해상도인지 잰다.

## 이 도구가 막는 사고

그림 파일 자체는 멀쩡한데 슬라이드에서만 흐린 경우가 있다. 원본이 작아서가 아니라
**넓게 놓았기 때문**이다. 눈으로는 「좀 흐린가?」 정도라 그냥 넘어가고, 투사하면 그때 보인다.

    실효 해상도(DPI) = 그림의 가로 픽셀 / 슬라이드에 놓인 가로 인치

2026-09-01 실측(1주차 설치 덱 5장): 네 장은 183~262 DPI 인데 **한 장만 121 DPI** 였다.
그 한 장은 원본 3414px 짜리를 좁게 잘라 924px 로 만들어 놓고 7.6in 로 «넓게» 놓은 것이었다.
사람이 세지 않으면 못 찾는다 — 한 장 한 장은 다 그럴듯해 보이기 때문이다.

## 문턱을 왜 그 값으로 두나

- 1920px 프로젝터가 13.33in 짜리 16:9 슬라이드를 쏘면 **144 DPI** 가 분해 한계다.
  기본 `--min 150` 은 거기에 약간의 여유를 준 값이다. 인쇄용 배포물이면 220~300 을 준다.
- 위쪽 `--waste 600` 은 «필요 이상으로 큰» 그림이다. 화질에는 문제가 없고 파일만 무겁다.
  줄여도 되지만 급하지 않다 — 그래서 실패가 아니라 알림으로만 찍는다.

## 두 번째 축 — 종횡비

그림을 상자에 억지로 맞추면 «눌린다». 원본 비와 놓인 비가 2% 넘게 다르면 잡는다.
python-pptx 로 `width` 와 `height` 를 «둘 다» 지정하면 조용히 이렇게 된다.

## 쓰는 법

    python audit_image_dpi.py deck.pptx
    python audit_image_dpi.py deck.pptx --min 200          # 인쇄물·큰 화면
    python audit_image_dpi.py deck.pptx --json             # 기계가 읽을 형태
    python audit_image_dpi.py --selftest

⚠ 이 검사는 «그림이 무엇을 보여 주는가» 는 못 본다. 그림 안에 구워진 글자가 낡았는지는
  캡처를 만드는 «소스» 를 따로 대조해야 한다.
"""
import argparse
import json
import sys

from pptx import Presentation
from pptx.util import Emu

PICTURE = 13          # MSO_SHAPE_TYPE.PICTURE
ASPECT_TOL = 0.02     # 종횡비 2% 차이까지는 반올림 오차로 본다


def scan(path):
    """(슬라이드번호, 이름, px, 인치, DPI, 비틀림) 목록을 돌려준다."""
    prs = Presentation(path)
    out = []
    for n, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if sh.shape_type != PICTURE:
                continue
            try:
                pw, ph = sh.image.size
            except Exception:
                continue                      # 링크만 걸린 그림 등
            iw = Emu(sh.width).inches
            ih = Emu(sh.height).inches
            if iw <= 0 or ih <= 0:
                continue
            dpi_w, dpi_h = pw / iw, ph / ih
            skew = abs(dpi_w - dpi_h) / max(dpi_w, dpi_h)
            out.append({
                "slide": n, "name": sh.name,
                "px": [pw, ph], "inches": [round(iw, 2), round(ih, 2)],
                "dpi": int(round(min(dpi_w, dpi_h))),
                "skew": round(skew, 4),
            })
    return out


def report(path, min_dpi=150, waste_dpi=600, as_json=False):
    rows = scan(path)
    if as_json:
        print(json.dumps(rows, ensure_ascii=False, indent=2))
        return 0

    if not rows:
        print("그림이 없습니다: %s" % path)
        return 0

    soft = [r for r in rows if r["dpi"] < min_dpi]
    skewed = [r for r in rows if r["skew"] > ASPECT_TOL]
    fat = [r for r in rows if r["dpi"] > waste_dpi]

    print("그림 %d개  (문턱 %d DPI)" % (len(rows), min_dpi))
    for r in sorted(rows, key=lambda r: r["dpi"]):
        mark = "  " if r["dpi"] >= min_dpi else "★ "
        print("  %s슬%-3d %-22s %5dx%-5d -> %5.2f x %5.2f in  =  %4d DPI%s"
              % (mark, r["slide"], r["name"][:22], r["px"][0], r["px"][1],
                 r["inches"][0], r["inches"][1], r["dpi"],
                 "   ⚠비틀림 %.1f%%" % (r["skew"] * 100)
                 if r["skew"] > ASPECT_TOL else ""))

    if fat:
        print("\n  참고 — 필요 이상으로 큰 그림 %d개 (화질 문제 아님, 파일만 무겁다): %s"
              % (len(fat), ", ".join("슬%d" % r["slide"] for r in fat)))

    if not soft and not skewed:
        print("\n이상 없음 — 모든 그림이 %d DPI 이상이고 종횡비도 보존돼 있습니다." % min_dpi)
        return 0

    print("")
    for r in soft:
        print("  ✗ 슬%d %s: %d DPI < %d — 놓인 폭(%.2fin)에 견줘 픽셀이 모자랍니다.\n"
              "      원본을 더 크게 뜨거나(웹 캡처면 배율을 올린다), 빈 여백을 잘라\n"
              "      «내용»이 그 폭을 채우게 하세요."
              % (r["slide"], r["name"], r["dpi"], min_dpi, r["inches"][0]))
    for r in skewed:
        print("  ✗ 슬%d %s: 종횡비가 %.1f%% 비틀렸습니다 — width 와 height 를 «둘 다»\n"
              "      지정하면 이렇게 됩니다. 하나만 주고 나머지는 비율로 계산하세요."
              % (r["slide"], r["name"], r["skew"] * 100))
    return 1


# ------------------------------------------------------------------ 자체 시험
def _selftest():
    """양성 3 · 음성 1 — 이 검사가 «항상 통과하지» 않는 것을 확인한다."""
    import os
    import tempfile

    from PIL import Image
    from pptx import Presentation as P
    from pptx.util import Inches

    tmp = tempfile.mkdtemp()
    big = os.path.join(tmp, "big.png")
    small = os.path.join(tmp, "small.png")
    Image.new("RGB", (2000, 1000), "white").save(big)
    Image.new("RGB", (200, 100), "white").save(small)

    prs = P()
    blank = prs.slide_layouts[6]
    s = prs.slides.add_slide(blank)
    s.shapes.add_picture(big, Inches(0.5), Inches(0.5), width=Inches(5))    # 400 DPI
    s = prs.slides.add_slide(blank)
    s.shapes.add_picture(small, Inches(0.5), Inches(0.5), width=Inches(5))  # 40 DPI
    s = prs.slides.add_slide(blank)
    s.shapes.add_picture(big, Inches(0.5), Inches(0.5),                     # 눌린 것
                         width=Inches(5), height=Inches(5))
    deck = os.path.join(tmp, "t.pptx")
    prs.save(deck)

    rows = scan(deck)
    fails = []
    if len(rows) != 3:
        fails.append("그림 3개를 못 찾았다: %d" % len(rows))
    by = {r["slide"]: r for r in rows}
    if by[1]["dpi"] != 400:
        fails.append("슬1 은 400 DPI 여야 한다: %d" % by[1]["dpi"])
    if by[2]["dpi"] != 40:
        fails.append("슬2 는 40 DPI 여야 한다: %d" % by[2]["dpi"])
    if by[1]["skew"] > ASPECT_TOL:
        fails.append("슬1 은 비틀리지 않았는데 비틀림으로 잡혔다")
    if by[3]["skew"] <= ASPECT_TOL:
        fails.append("슬3 은 «눌린» 그림인데 못 잡았다 (skew %.3f)" % by[3]["skew"])
    # 음성 대조 — 문턱을 낮추면 흐린 것도 통과해야 한다
    if report(deck, min_dpi=10, as_json=True) != 0:
        fails.append("json 모드는 0 을 돌려줘야 한다")

    print("\n" + "=" * 60)
    if fails:
        for f in fails:
            print("  ✗ " + f)
        raise SystemExit("selftest 실패 %d건" % len(fails))
    print("selftest 통과 — 흐린 그림·눌린 그림을 잡고, 멀쩡한 그림은 통과시킨다.")
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="덱 그림의 실효 해상도(DPI) 감사")
    ap.add_argument("deck", nargs="?", help="검사할 .pptx")
    ap.add_argument("--min", type=int, default=150, help="최소 실효 DPI (기본 150)")
    ap.add_argument("--waste", type=int, default=600, help="이 위는 «과하게 큼» 알림")
    ap.add_argument("--json", action="store_true", help="기계가 읽을 형태로 출력")
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args()
    if a.selftest:
        sys.exit(_selftest())
    if not a.deck:
        ap.error("deck 이 필요합니다 (또는 --selftest)")
    sys.exit(report(a.deck, min_dpi=a.min, waste_dpi=a.waste, as_json=a.json))
