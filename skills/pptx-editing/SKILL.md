---
name: pptx-editing
description: >-
  Build, edit, and QA PowerPoint (.pptx) decks with python-pptx without the
  common failures that make a deck look broken or lose work. Use this whenever a
  task involves a .pptx file, a PowerPoint / 파워포인트 / 발표자료 / 슬라이드 / 덱 /
  lab-meeting deck / research brief / conference slides — including creating a
  deck from a template, adding or editing slides, text, shapes, tables, figures,
  charts, section dividers, and especially SPEAKER NOTES (발표자 노트 / 슬라이드
  노트), fitting images without overflow, keeping numbers traceable to source
  files, and rendering the deck to PDF/PNG for a visual check. Trigger even if the
  user only says "이 슬라이드 고쳐줘" / "발표자료 만들어줘" / ".pptx" and doesn't
  mention the internals — naive python-pptx edits silently drop notes, overflow
  slide boundaries, or get wiped on the next rebuild.
license: MIT
---

# PPTX Editing

A `.pptx` is a **zip of XML (OOXML / PresentationML)**. `python-pptx` covers most
of it, but the parts that bite are the ones it *doesn't* surface cleanly. **Read
`references/pptx-guide.md` before editing**, and **use `scripts/pptx_kit.py` +
`scripts/inspect_pptx.py`** instead of hand-rolling the fiddly bits or eyeballing
correctness.

## The rule that would have saved the most time

**When something basic looks impossible, the bug is your assumption, not the
format.** The canonical example: a template whose notes master has no body
placeholder makes `slide.notes_slide.notes_text_frame` return `None`, so setting
`.text` raises. That is **not** "this template can't hold speaker notes" —
PowerPoint shows notes fine. The fix is to inject a body placeholder into the
notes slide (`pptx_kit.speaker_note` does this). Don't conclude "impossible";
check the assumption.

## Two rules that protect the user's work

1. **A rebuild regenerates the whole .pptx from your script — it WIPES anything
   added by hand in PowerPoint** (notes, a moved box, a typed number). So **put
   everything in the build script**, and tell the user: don't hand-edit a deck
   that will be rebuilt. Speaker notes especially go in the script
   (`pptx_kit.speaker_note`), never typed into PowerPoint.
2. **Guard overwrites.** Builders use `build_guard` (md5 fingerprint). If it
   blocks a save, the on-disk file changed since the last build — **someone may
   have hand-edited it.** Look before you clobber: back the file up, check what
   changed, and only then delete `<out>.build-md5` and rebuild.
3. **A "do not rebuild" ban is a last resort, not a destination.** If the reason a
   deck is unsafe to rebuild is that the script can't reproduce some hand-added
   thing (native equations, a real SVG), the right fix is usually to fold that
   *technique* into the script (see guide §10 for the equation case) so the
   generator becomes valid again — not to leave a ban in
   place forever. A ban that must live in someone's memory across sessions WILL
   get missed (a session that skips reading the project's SESSION_LOG.md will
   rebuild anyway and wipe the hand edits — this happened). If a ban is still
   necessary for some other unclosed gap, put it in **bold, at the very top** of
   the build script's own docstring — not mid-file near an unrelated slide's
   comment, where the next session won't see it before running the script.

## Workflow

1. **Inspect first.** `python scripts/inspect_pptx.py FILE.pptx` — per-slide dump
   of shapes, text, tables, pictures, **speaker notes**, and a **boundary-overflow
   flag** for any shape whose bottom/right leaves the slide. If you're editing an
   existing deck you didn't build, read it fully before changing it.
2. **Read the matching section of `references/pptx-guide.md`** (map at its top).
   Placeholder indices, layout names, and the notes master differ per template —
   read them from the actual file, never assume.
3. **Build with the helpers.** Keep project STYLE (palette, fonts, card/section
   components) in the project's own `kit_common.py`; import palette-agnostic
   MECHANICS from `pptx_kit` (`new_deck`, `blank_slide_layout`, `rect`,
   `slide_number`, `speaker_note`, `fit_picture`). Numbers on a slide must be read
   from a source file, never typed from memory (`ppt_rules` — data integrity).
4. **Never overflow.** Size images with `pptx_kit.fit_picture` (reads real pixel
   dims via PIL, fits inside a box preserving aspect). Textboxes don't autofit
   reliably — size them generously. **Footnote/table-note boxes are a common
   miss**: helpers like `footnote()`/table `note=` draw a FIXED-height box, so
   adding a citation or a longer caption can push text past the slide edge with
   no error — only a render catches it (happened 3x in one 2026-08-06 session:
   citation footnotes, an obesity-sensitivity table note). When you lengthen any
   note/footnote string, re-render and check that slide specifically, not just
   the ones you added.
5. **Render, then LOOK.** `python scripts/render_pptx.py FILE.pptx` (PowerPoint
   COM → PDF → PNG). python-pptx cannot tell you a card clipped its last line, a
   footnote bled off the slide, or a title overran. **Any layout-affecting change
   needs a rendered look** — steps 1–4 cannot see it.
6. **Audit numbers and type.** Cross-check every figure against its source file;
   `agent/tools/deck_audit.py`, `audit_font_sizes.py`, `audit_table_widths.py`,
   `audit_text_consistency.py` automate the mechanical passes.

## Scripts (`scripts/`) — import/run these, don't reinvent

| Script | Purpose |
|---|---|
| `pptx_kit.py` | Palette-agnostic **mechanics**: `new_deck`, `blank_slide_layout`, `rect`, `slide_number`, **`speaker_note`** (rebuild-proof notes; injects the missing placeholder), `fit_picture` (PIL-measured, overflow-safe image), `overflows` (boundary check). Also carries native-equation builders (`equation_slot`, `promote_equations`, `m_frac`/`m_sub`/`m_sup`/`m_nary`/`m_sqrt`/`m_acc`) — only relevant if a slide needs a real OOXML equation object; see guide §10 before using these. Import it; project style layers on top. |
| `inspect_pptx.py FILE` | Structure + notes + **overflow** dump. First thing to run on any deck. |
| `render_pptx.py FILE [--pdf OUT]` | Render to PDF (PowerPoint COM) then PNG per slide, for the §5 visual check. |
| `selftest.py` | Proves `speaker_note` round-trips (write → reopen → read) on a placeholder-less notes master, with no real template. |

**Shared lab tools this skill leans on** (in `agent/tools/`, already cross-project):
`build_guard.py` (md5 overwrite guard + stamp), `flowchart_generator.py`
(STROBE flowcharts), `manuscript_table.py` (`add_journal_table` — read tables from
a source CSV, never hand-type), `brief_builder.py`, and the `*_audit.py` passes.

## Where to read in the guide (`references/pptx-guide.md`)

Opens with a **"흔한 실패 TOP"**; skim that, then jump to the section you need:

- **§1** python-pptx model · EMU/Inches · templates, layouts, the blank-layout trap
- **§2** **speaker notes** — the placeholder-less-master fix; notes-in-script
- **§3** **rebuild wipes hand edits** · `build_guard` md5 · date-stamped outputs
- **§4** images — PIL-measured fit, overflow check, aspect ratio
- **§5** text — no reliable autofit, sizing boxes, font floors (ppt_rules)
- **§6** tables — `add_journal_table` from source CSV; the header_row/build_row pattern
- **§7** rendering & visual QA (COM/LibreOffice) · Windows console Unicode
- **§8** data integrity — every number traced to a source file
- **§9** teaching-deck (학습 PPT) norms — different rules than an academic deck: what belongs in
  the student file vs. the answer-key file, definitions the audience already knows but students don't
- **§10** native OOXML equations (Insert > Equation) — finding/replacing existing ones, building new
  ones from scratch (no python-pptx or COM API for this), and a PyMuPDF overlap-checker pitfall.
  **Skip this unless a slide needs a real equation object** — `formula()`-style monospace text boxes
  don't need any of it.

## Bash 툴 + 한글 경로: 간헐적 mojibake

Git Bash(Bash 툴)로 `python script.py`를 실행할 때, 스크립트 안에 한글 경로 리터럴이 있으면
**똑같은 스크립트가 방금 전엔 성공했는데 다음번엔 `PackageNotFoundError`로 실패**하는 일이
발생한다(트레이스백에 경로가 깨져 나옴) — 일회성이 아니라 반복 재현된다. 재시도로 안
고쳐지면 **PowerShell 툴로 같은 스크립트를 실행**해볼 것 — 스크립트 자체는 멀쩡하고 Bash
툴의 인코딩 처리가 문제인 경우가 많다.

## 손으로 수정한 파일 위에서 작업하기

빌드 스크립트로 만든 덱이라도 사용자가 PowerPoint에서 직접 수정했다면, 그 순간부터 **손으로
수정한 라이브 파일이 정본**이다(§3 규칙의 연장). 이번 세션에서 실제로 유효했던 절차:

1. 수정 전 반드시 날짜/목적이 담긴 이름으로 백업 (`BACKUP_<파일>_before_<의도>.pptx`).
2. python-pptx로 특정 run의 `.text`만 재할당 — 서식(`rPr`)은 그대로 두고 텍스트만 바꾼다.
   문단이 여러 run으로 쪼개져 있으면(PowerPoint 맞춤법 검사가 자동으로 쪼갠다) 각 run을
   개별적으로 확인하고 바꿀 것 — 통짜 문자열 치환은 run 경계와 안 맞아 실패하거나 서식을
   깨뜨린다.
3. python-pptx가 못 보는 요소(mc:AlternateContent 등, 가이드 §10)는 zip을 열어 해당 XML 엔트리만
   문자열/정규식으로 패치하고 다시 압축.
4. 편집 후 **다시 python-pptx로 열어서** 저장이 깨지지 않았는지, 그리고 raw XML로 다시
   검증할 것 — `prs.save()`는 전체 트리를 그대로 직렬화하므로 손대지 않은 mc:AlternateContent
   블록은 보존되지만, 매번 재확인하는 습관을 들일 것 (특히 사용자가 그 사이에 PowerPoint에서
   또 손을 댔을 수 있다 — zip 엔트리 개수가 예상과 다르면 그 신호일 수 있다).

## Guardrails

- Edits documents only; needs no credentials, fetches/executes nothing remote.
- Work on a **copy** / date-stamped output; never overwrite the only copy.
- **Never publish or hand off a deck that impersonates a real org/person or
  presents fabricated records as genuine.** Numbers come from real source files.
