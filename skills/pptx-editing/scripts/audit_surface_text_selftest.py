# -*- coding: utf-8 -*-
"""audit_surface_text_selftest.py — 양성·음성 대조.

**케이스의 절반이 «걸리면 안 되는» 것이다.** 표면 감사기의 진짜 실패 모드는 «못 잡는 것»이
아니라 «정상을 잡는 것»이다 -- 오탐을 내는 게이트는 곧 아무도 안 보는 게이트가 되고,
그러면 신호가 0이 된다(`precommit_scan`이 「암호」가 든 줄을 전부 잡아 항상 종료코드 1이었던
사고와 같은 부류).

    python audit_surface_text_selftest.py     # exit 1 if any case fails
"""
import re
import sys
import audit_surface_text as M

# (범주, 텍스트, 걸려야 하나)
CASES = [
    # ── rhetoric: 카드 제목이 «수사적 기능» ──────────────────────────────
    ("rhetoric", "The hook", True),
    ("rhetoric", "Key point", True),
    ("rhetoric", "Organizing principle", True),
    ("rhetoric", "The consequence:", True),
    ("rhetoric", "Why this matters", True),
    # 같은 낱말이 «문장 안»이면 정상 -- 제목일 때만 수사다
    ("rhetoric", "The hook of the sampling frame is described in Methods.", False),
    ("rhetoric", "Key point estimates are shown with 95% CI.", False),
    ("rhetoric", "Main findings", False),
    ("rhetoric", "Study design and population", False),
    # ── rhetoric: 기호로 쓴 말 ────────────────────────────────────────
    ("rhetoric", "Proteobacteria ↑ in cases", True),
    ("rhetoric", "risk ↓ after adjustment", True),
    # 일부러 뺀 것: 통계 덱에서 정당한 수학 기호
    ("rhetoric", "p ≠ 0.05 in the adjusted model", False),
    ("rhetoric", "OR ≈ 1.4 (95% CI 1.1-1.8)", False),
    # ── meta / nav / restate / provenance 회귀 ────────────────────────
    ("meta", "This is justified by the design.", True),
    ("nav", "see Table 3 for details", True),
    ("restate", "I.e., the effect is null.", True),
    ("provenance", "Source: Table2_main_260819", True),
    ("meta", "Adjusted for age, sex, and smoking.", False),
    ("nav", "Table 1. Baseline characteristics", False),
]

# caps_emphasis 는 «덱 전체» 문맥이라 따로 -- (paras, 걸려야 하는 낱말들)
DECKS = [
    # 강조: 같은 낱말이 대문자·소문자로 둘 다, 그리고 문장 «안»에 있다
    ([(1, "Proteobacteria DEPLETES in the exposed group overall"),
      (2, "the genus depletes across all strata")], ["DEPLETES"]),
    # 약어: 소문자로 등장하지 않는다 -> 걸리면 안 된다
    ([(1, "NHANES 2017-2018 cycle was used for the analysis"),
      (2, "We applied TMLE with SuperLearner to the pooled sample"),
      (3, "STROBE flowchart is shown in the appendix")], []),
    # 표 헤더가 설계상 ALLCAPS -- 본문에 소문자가 있어도 걸리면 안 된다
    ([(1, "MEAN"), (1, "SD"),
      (2, "the mean difference was 2.3 units across the two arms")], []),
]


def run():
    bad = []
    for cat, text, want in CASES:
        pat = re.compile(M.DEFAULT[cat], re.I)
        got = bool(pat.search(text))
        if got != want:
            bad.append("[%s] %r  기대=%s 실제=%s" % (cat, text[:60], want, got))

    for i, (paras, want_words) in enumerate(DECKS, 1):
        hits = M.caps_emphasis(paras)
        got = sorted(h[2] for h in hits)
        if got != sorted(want_words):
            bad.append("[caps_emphasis] deck%d  기대=%s 실제=%s" % (i, want_words, got))

    total = len(CASES) + len(DECKS)
    neg = sum(1 for _, _, w in CASES if not w) + sum(1 for _, w in DECKS if not w)
    print("%d cases (%d of them must NOT fire)" % (total, neg))
    if bad:
        print("FAIL %d:" % len(bad))
        for b in bad:
            print("  " + b)
        return 1
    print("  OK  all pass")
    return 0


if __name__ == "__main__":
    sys.exit(run())
