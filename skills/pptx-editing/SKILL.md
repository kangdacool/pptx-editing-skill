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
   **Import it — never paste a copy into the project.** See the next section for
   how, and for what copying actually costs.
4. **Never overflow.** Size images with `pptx_kit.fit_picture` (reads real pixel
   dims via PIL, fits inside a box preserving aspect). Textboxes don't autofit
   reliably — size them generously. **Footnote/table-note boxes are a common
   miss**: helpers like `footnote()`/table `note=` draw a FIXED-height box, so
   adding a citation or a longer caption can push text past the slide edge with
   no error — only a render catches it (happened 3x in one 2026-08-06 session:
   citation footnotes, an obesity-sensitivity table note). When you lengthen any
   note/footnote string, re-render and check that slide specifically, not just
   the ones you added.
5. **Measure the fit, then render and LOOK.** Run
   `python scripts/audit_text_fit.py FILE.pptx` first: it asks PowerPoint for the size the
   text *actually* occupies (`TextFrame2.TextRange.BoundWidth/BoundHeight`) and reports only
   text that runs **off the slide** or lands **on another text/picture**. That replaces
   hunting for overflow by eye, which is what turns one tweak into a render-adjust cycle.
   Then `python scripts/render_pptx.py FILE.pptx` (PowerPoint COM → PDF → PNG) and look —
   the audit cannot judge crowding, an ugly mid-word break, or a wrong colour. **Any
   layout-affecting change still needs a rendered look**; the audit just means you are no
   longer looking *for overflow*.
6. **Audit the surface text.** `python scripts/audit_surface_text.py FILE.pptx --max-fig N --max-tab N`.
   배포 전 무조건 1회. 지웠다고 생각한 뒤에도 걸린다 — 실제로 2026-08-10 세션에서 "다 지웠다"고
   넘긴 직후 4건이 더 나왔다.
7. **Audit numbers and type.** Cross-check every figure against its source file. Mechanical passes
   live outside this skill, in the lab's shared `agent/tools/`: `deck_audit.py` (XML referential
   integrity — sections, zoom links, creationId dupes, TOC↔divider match, orphaned media),
   `deck_render_audit.py` (rendered-pixel check for table-row overflow that coordinates miss),
   `audit_font_sizes.py`, `audit_table_widths.py`, `audit_text_consistency.py`. Run what's relevant
   to the deck at hand — not every pass fires on every deck.
8. **Audit narrative order — a separate pass from step 7, but a cheap one.**
   Unlike step 7 (which opens every source file to cross-check each figure —
   real audit weight), this is a single read of just the titles+subtitles in
   sequence (`python-pptx`, one script call, then read the ~1-line-per-slide
   output once). No source-file lookups. For a normal-sized deck (15-20
   slides) it's well under a minute; do it every time a deck is finalized or
   revised, not just on request. Number/content accuracy audits (step 7) do
   not catch sequencing bugs, and neither does a render-and-look (step 5,
   which sees one slide at a time). Read the deck's
   titles/subtitles/bullets *in order* looking specifically for: (a) a term or
   abbreviation used substantively before the slide that defines it, (b) a
   slide that references "the finding we just saw" / "as shown above" for
   content that actually appears *later* in the deck. Both are invisible to
   grep-for-wrong-numbers and invisible to single-slide overflow checks — they
   only surface by reading slide N's text against what slide N-1 vs N+5
   actually established. Real example (blue_green_lone, 2026-08-07): a slide
   used "NDVI" as if already defined, 7 slides before the slide that defined
   it; another slide's punchline sentence presupposed a finding not revealed
   until 7 slides later. Both read fine in isolation — only the order was
   wrong. When the user says something like "the order matters" or "check
   this is accurate," run this pass explicitly; don't fold it into step 7 and
   call it done.

## Reordering slides in a script-generated deck — parse blocks, don't hand-edit

If slides are built in sequence by a script with per-slide comment markers
(e.g. `# ==== Slide N — title`), moving a slide (or a run of slides) by
cut-pasting the source lines by hand is exactly the kind of large multi-line
edit that risks slicing through a shared variable or leaving an orphaned
half-block. Safer: parse the file into blocks split on the marker regex, hold
them in an ordered list, splice/insert by label, and reassemble — the same
approach a mail-merge or template engine uses, applied to the script's own
source.
```python
import re
marker_re = re.compile(r'^# =+ (Slide[^\n]*)\n', re.MULTILINE)
markers = list(marker_re.finditer(text))
preamble = text[:markers[0].start()]
blocks, order = {}, []
for i, m in enumerate(markers):
    end = markers[i + 1].start() if i + 1 < len(markers) else len(text)
    label = m.group(1).strip()
    blocks[label] = text[m.start():end]
    order.append(label)

# move "Slide 6b" to sit right before "Slide 3":
order.remove("Slide 6b — title")
order.insert(order.index("Slide 3 — title"), "Slide 6b — title")

new_text = preamble + "".join(blocks[l] for l in order)
```
The point: parse to blocks first, reorder the *list of labels* with
`list.remove`/`list.insert`, reassemble last. Never find-and-replace across a
slide boundary by hand. This only works if each
block is self-contained (loads its own data, doesn't depend on a variable
defined by a block between its old and new position) — check for that first;
if two blocks share state, move them as a unit. After reordering, rebuild and
re-render the moved slides specifically (per step 7 above, not just step 5)
to confirm the surrounding narrative now reads correctly, not just that nothing overflows.

## Importing `pptx_kit` from a project — copy it and it WILL rot

"Import the mechanics, keep the style local" only works if the project actually
imports. The failure mode is quiet: the project has no obvious way to reach
`pptx_kit.py`, so someone pastes the functions into the project's own kit "just
for now". Nothing breaks that day — the copies drift later, and by then the two
files look independent.

Drop this in the project's kit module instead. No `pip install`, no `sys.path`
hardcoding, works on any machine where the skill is installed:

```python
def _load_pptx_kit():
    """Load pptx_kit.py from the installed pptx-editing skill (single source)."""
    import importlib.util, os, sys
    cands = [
        os.environ.get("PPTX_KIT"),   # escape hatch: skill installed elsewhere
        os.path.expanduser(os.path.join(
            "~", ".claude", "skills", "pptx-editing", "scripts", "pptx_kit.py")),
    ]
    for p in cands:
        if p and os.path.isfile(p):
            spec = importlib.util.spec_from_file_location("pptx_kit", p)
            mod = importlib.util.module_from_spec(spec)
            sys.modules["pptx_kit"] = mod
            spec.loader.exec_module(mod)
            return mod
    raise ImportError("pptx_kit.py not found — install the pptx-editing skill, "
                      "or set PPTX_KIT to its path.")

_kit = _load_pptx_kit()
m_frac = _kit.m_frac          # re-export what this project uses, by name
equation_slot = _kit.equation_slot
promote_equations = _kit.promote_equations
```

Re-export **explicitly by name** rather than `from pptx_kit import *` — it stays
obvious which names are the kit's and which are the project's.

**Resolve the path; never hardcode it.** `sys.path.insert(0, r"D:\...\skills\pptx-editing\scripts")`
works until the next machine, where the drive letter or install root differs. Two projects here had
drifted into two different mechanisms (one hardcoded absolute, one resolved) — unify on resolution:
`$PPTX_KIT` → `$CLAUDE_SKILLS_DIR/pptx-editing/scripts` → `~/.claude/skills/...`, and fail with a
message that names the directory it looked in.

**Do NOT install the skill as a symlink/junction inside a cloud-synced folder.** Cloud clients
(OneDrive/Dropbox) do not replicate reparse points, so the skill silently disappears on your other
machines — while looking perfectly fine on the one where you created it. Keep the installed copy a
**real directory**; if it and the repo ever disagree, the repo wins — copy it over.

**Keep local, on purpose**, anything whose *semantics* differ from the kit's —
don't swap it in just because the name matches. Real example: a project's
`check_overflow` used a **1-inch** tolerance while `pptx_kit.overflows` uses
**0.02 in**; silently switching would have blocked builds that were passing by
design. Same-name-different-meaning is fine as long as it's deliberate and
commented.

**What copying cost, concretely** (선택교과4, 2026-08-06): the project had pasted
12 kit functions into its own `ppt_common.py`, and that file itself existed in 5
copies (root + 4 per-day folders). A later session added the equation machinery to
**one** copy. Result: the root generator raised `NameError: equation_slot is not
defined` — a build script that simply could not run, while a sibling copy worked
fine. Consolidating to one import fixed it, and regenerating all 7 decks produced
**element-for-element identical output**, native equations included — proof the
copies had been pure redundancy all along. If you inherit a project in this state,
verify equivalence before deleting: run both implementations on the same inputs and
diff the returned XML, then re-render every deck and compare.

## Scripts (`scripts/`) — import/run these, don't reinvent

| Script | Purpose |
|---|---|
| `pptx_kit.py` | Palette-agnostic **mechanics**: `new_deck`, `blank_slide_layout`, `rect`, `slide_number`, **`speaker_note`** (rebuild-proof notes; injects the missing placeholder), `fit_picture` (PIL-measured, overflow-safe image), `overflows` (boundary check), `hang` (hanging indent — python-pptx has no property for it), `text_units`/`wrapped_row_count` (Hangul-aware wrap-length estimate, for pre-sizing a card before drawing it), `check_surface_leaks`/`save_and_check` (gate a save on caller-supplied banned-phrase hits + overflow — see §11 for why the phrase list is never a shared default). Also carries native-equation builders (`equation_slot`, `promote_equations`, `m_frac`/`m_sub`/`m_sup`/`m_nary`/`m_sqrt`/`m_acc`) — only relevant if a slide needs a real OOXML equation object; see guide §10 before using these. Import it; project style layers on top. |
| `inspect_pptx.py FILE` | Structure + notes + **overflow** dump. First thing to run on any deck. |
| `audit_text_fit.py FILE [--all]` | Asks PowerPoint how big the text really is and flags only text that runs **off-slide** or **onto another text/picture**. Exit 1 on a hit, so a build can gate. **A textbox does not clip — it spills**, and spilling over a background fill is normal layering; treating either as an error makes the check cry wolf (two earlier cuts of this script did exactly that, 3/3 false on a clean deck). |
| `audit_surface_text.py FILE [--max-fig N] [--max-tab N]` | «있으면 안 되는 말»이 남았는지 기계로 훑는다 — 편집 해명·내비게이션 안내·재진술 신호·내부 파일명, 그리고 **번호 drift**(산출물엔 Figure가 3개인데 캡션에 "Figure 9"가 남은 경우). `audit_text_fit.py`가 «글자가 넘치는가»를 본다면 이건 «내용이 표면에 남았는가»를 본다. exit 1이라 빌드 게이트로 쓸 수 있다. 사람 눈으로 훑는 방식은 반복해서 실패한다. |
| `render_pptx.py FILE [--pdf OUT]` | Render to PDF (PowerPoint COM) then PNG per slide, for the §5 visual check. |
| `selftest.py` | Proves `speaker_note` round-trips (write → reopen → read) on a placeholder-less notes master, with no real template. |
| `poster_kit.py` | Palette-agnostic mechanics for **large-format academic posters** (cm-scale canvas, not a 16:9 slide) — `two_col_grid`, `sectitle`, `bullets`/`caption`/`ptable` (all with an enforced minimum legible font size via `kf`), `pic_cm` (width-locked, no silent shrink-below-floor), `min_font_report` (catches text that bypassed `kf`). See guide §12 before building a poster — the 2-column-not-3, one-accent-color, narrative-caption habits it encodes. |
| `poster_kit_selftest.py` | Proves `poster_kit`'s font-floor enforcement and grid math without a real poster project. |

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
- **§11** surface-leak gating (`check_surface_leaks`/`save_and_check`) — why the banned-phrase list
  is always caller-supplied, never a shared default
- **§12** **academic posters** (`poster_kit.py`) — 2-column-not-3, enforced min font size, figures at
  layout size (never shrunk post-hoc), narrative captions, one accent color. Read before any poster build.

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
