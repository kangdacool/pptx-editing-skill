# -*- coding: utf-8 -*-
"""deck_mono 수정 자기시험 — 통과할 것과 실패할 것을 하나씩 넣어 본다."""
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.expanduser("~"), ".claude",
                                "skills", "pptx-editing", "scripts"))
import deck_mono as dm

fails = []

# ── 1. 기본 글꼴은 맑은 고딕이어야 한다 ──────────────────────────────
prs = dm.new_deck()
dm.set_running("smoke")
s = dm.content(prs, "제목", "절")
dm.bullets(s, ["한 줄"])
def fonts(pr):
    """tb 는 «문단» 글꼴을, bullets/card 는 «run» 글꼴을 쓴다 — 둘 다 봐야 한다."""
    out = set()
    for sl in pr.slides:
        for sh in sl.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                if p.font.name:
                    out.add(p.font.name)
                for r in p.runs:
                    if r.font.name:
                        out.add(r.font.name)
    return out


names = fonts(prs)
if names != {"맑은 고딕"}:
    fails.append(f"기본 글꼴이 맑은 고딕이 아니다: {names}")

# ── 2. set_font("Arial") 이 tb 가 만든 글자에도 먹어야 한다 ──────────
dm.set_font("Arial")
prs2 = dm.new_deck()
s2 = dm.content(prs2, "Title", "Section")       # content 는 tb 로 제목을 만든다
dm.table(s2, ["A", "B"], [["1", "2"]], col_x=[0.8, 5.0], col_w=[4.0, 4.0])
names2 = fonts(prs2)
if "맑은 고딕" in names2:
    fails.append(f"set_font 이후에도 맑은 고딕이 남는다: {names2}")
if "Arial" not in names2:
    fails.append(f"set_font 이 안 먹었다: {names2}")

# ── 3. 그림자가 꺼져 있어야 한다 (카드·괘선·배경) ────────────────────
s3 = dm.content(prs2, "Card", "Section")
dm.card(s3, 0.8, 1.7, 5.0, 1.5, "Header", ["line"])
# 레이아웃이 딸려 보내는 빈 placeholder 는 deck_mono 가 만든 것이 아니라 대상이 아니다.
on = []
for sl in prs2.slides:
    for sh in sl.shapes:
        if sh.is_placeholder:
            continue
        try:
            if sh.shadow.inherit:
                on.append(str(sh.shape_type))
        except (AttributeError, NotImplementedError):
            pass
if on:
    fails.append(f"그림자가 켜진 도형 {len(on)}개: {set(map(str, on))}")

# ── 4. font= 를 명시하면 그것이 이겨야 한다 ──────────────────────────
s4 = dm.content(prs2, "X", "")
dm.tb(s4, 1, 1, 3, 0.5, "explicit", font="Consolas")
box = s4.shapes[-1]
if box.text_frame.paragraphs[0].font.name != "Consolas":
    fails.append("명시한 font 인자가 무시된다")

# ── 5. 저장까지 되는가 ───────────────────────────────────────────────
out = os.path.join(tempfile.gettempdir(), "_deck_mono_smoke.pptx")
prs2.save(out)
os.remove(out)

dm.set_font("맑은 고딕")            # 원상복구

print("\n".join("FAIL " + f for f in fails) if fails else "PASS 5/5")
sys.exit(1 if fails else 0)
