# -*- coding: utf-8 -*-
"""audit_surface_text.py FILE.pptx [--extra-pattern RE ...]

배포 직전 «표면 텍스트» 감사. audit_text_fit.py가 «글자가 넘치는가»를 본다면,
이 스크립트는 «있으면 안 되는 말이 남았는가»를 본다. 둘은 다른 실패를 잡으므로 둘 다 돌린다.

왜 필요한가 (2026-08-10, teacher_violence_qol):
  같은 세션에서 사용자가 "쓸데없는 서술이 많다"고 두 번 지적했고, 지운 뒤에도 즉석 grep을
  돌리자 4건이 더 나왔다. 사람 눈으로 훑는 방식은 반복해서 실패한다 -- 기계로 훑어야 한다.

무엇을 잡나 (기본 4종):
  meta        편집 해명·검정 의미 해설 ("justified", "not just", "are not interpreted")
  nav         내비게이션 안내 ("see Table 3", "column 2", "decomposed in")
  restate     재진술 신호 ("I.e.,", "in other words")
  provenance  내부 파일명·파이프라인 잔재 (".csv", "Source: TableX_...")

번호 drift(파이프라인 번호가 산출물에 남은 것)는 --max-fig/--max-tab로 잡는다:
  산출물에 Figure가 3개뿐인데 "Figure 9"가 캡션에 남아 있으면 청중에게 없는 그림을 가리킨다.

exit 1 on any hit -- 빌드 파이프라인에서 게이트로 쓸 수 있다.
"""
import re
import sys
import argparse

DEFAULT = {
    "meta": r"\bjustified\b|\bnot just\b|are not interpreted|separation artifact|"
            r"\brobustness\b|sensitivity analys|\bverified\b|directly testing|"
            r"\bnotably\b|\bimportantly\b|\bhonestly\b|it is worth noting",
    "nav": r"see (Table|Figure|column|panel)|decomposed in|\bcolumn \d|as shown (above|below)|"
           r"refer to (Table|Figure)",
    "restate": r"\bI\.e\.|\bi\.e\.,|in other words|that is to say|this means that",
    "provenance": r"\.csv\b|\.rds\b|\.xlsx\b|Source: [A-Z][A-Za-z0-9_]*_|output/tables",
}


def paragraphs(path):
    from pptx import Presentation
    prs = Presentation(path)
    for si, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if sh.has_text_frame:
                for para in sh.text_frame.paragraphs:
                    t = "".join(r.text for r in para.runs).strip()
                    if t:
                        yield si, t
            if getattr(sh, "has_table", False):
                for r in sh.table.rows:
                    for c in r.cells:
                        t = c.text.strip()
                        if t:
                            yield si, t


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--extra-pattern", action="append", default=[],
                    help="추가 정규식 (프로젝트별 금지어). 여러 번 지정 가능.")
    ap.add_argument("--max-fig", type=int, default=None,
                    help="산출물의 Figure 최대 번호. 초과 참조를 번호 drift로 신고.")
    ap.add_argument("--max-tab", type=int, default=None,
                    help="산출물의 Table 최대 번호.")
    a = ap.parse_args()

    pats = {k: re.compile(v, re.I) for k, v in DEFAULT.items()}
    for i, ex in enumerate(a.extra_pattern):
        pats["custom%d" % (i + 1)] = re.compile(ex, re.I)

    paras = list(paragraphs(a.file))
    hits = []
    for si, t in paras:
        for name, pat in pats.items():
            m = pat.search(t)
            if m:
                hits.append((name, si, m.group(0), t))

    def drift(kind, cap):
        out = []
        for si, t in paras:
            for m in re.finditer(r"\b%s\s*(\d+)" % kind, t, re.I):
                if int(m.group(1)) > cap:
                    out.append(("number-drift", si, m.group(0), t))
        return out

    if a.max_fig is not None:
        hits += drift("Figure", a.max_fig)
    if a.max_tab is not None:
        hits += drift("Table", a.max_tab)

    print("%s -- %d text blocks scanned" % (a.file, len(paras)))
    if not hits:
        print("  OK  surface-text audit clean")
        return 0
    print("  %d hit(s):" % len(hits))
    for name, si, frag, t in hits:
        print("\n  [%s] slide %d  matched %r" % (name, si, frag))
        print("      " + (t[:220] + ("..." if len(t) > 220 else "")))
    return 1


if __name__ == "__main__":
    sys.exit(main())
