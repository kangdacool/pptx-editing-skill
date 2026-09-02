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

## 빌드 루프 — 재빌드로 레이아웃을 «찾지» 마라 (덱·포스터 공통)

2026-08-26 에 포스터 하나를 **25회 넘게** 재빌드했다. 대부분이 「빌드 → 넘쳤네 → 줄이고 →
빌드 → 남네 → 키우고」였다. 원인은 판단력이 아니라 **상자 높이를 «추정»으로 잡았던 것**이다.

```
1. 숫자로 먼저 맞춘다   pptx_kit.text_width_in() / wrapped_lines()  — 그리기 «전에» 부른다
                        (포스터는 poster_kit.text_width_cm() / est_lines(), 같은 구현)
2. 빌드 1회             python <build>.py
3. 렌더 1회             render_pptx.py
4. 실측 1회             ink_extent.py <pptx>   — 실제 잉크가 어디서 끝나나
5. 눈으로 1회           크롭해서 본다
```

**그림이 있는 덱이면 여기에 한 줄 더 — `audit_image_dpi.py <pptx>`.** 흐린 그림은 눈으로
넘어가고 «투사했을 때» 보인다. 한 장씩 보면 다 그럴듯하고, 덱 전체를 «세워 놓고 견줘야»
어느 장이 뒤처지는지 나온다(실측: 183·189·210·243·262 사이에 혼자 **121**). 새로 찍을
때는 반대 방향으로 — `web_shot.py --slide-width <놓을 인치>` 가 배율을 «계산»해서 찍는다.

**빌드는 «검증»이지 «탐색»이 아니다.** 폭·줄 수가 궁금하면 `text_width_in()` 한 번이
빌드+렌더 한 바퀴보다 훨씬 싸다. 「대충 넣고 렌더해서 보자」를 반복하면 그날이 재현된다.

**⚠ 상자 높이를 «추정»하지 마라 — 폭은 잴 수 있다.** 두 킷 모두 그 죄가 있었다:
`poster_kit.est_lines` 는 평균 문자폭(0.52em), `pptx_kit.wrapped_row_count` 는 「한글=2,
나머지=1」에 **레이아웃마다 손으로 보정하는 `wrap_units`** 를 썼다. 그 보정을 «렌더를 보며»
하는 것이 곧 재빌드 시행착오다. 2026-08-26 에 둘 다 **PIL 폰트 메트릭 실측 + 단어 단위
줄바꿈 시뮬레이션**으로 바꿨다(정본은 `pptx_kit.text_width_in`, poster_kit 은 그걸 부른다).
포스터 여백 오차가 **4.1cm → 0.2cm** 가 됐다.
`wrapped_row_count(lines, width_in=..., size_pt=...)` 로 부르면 실측 경로다 — **그게 기본**이고,
`wrap_units` 만 주는 옛 호출은 하위호환일 뿐이다.

**⭐ 그래도 실측이 «근사»인 이유, 그리고 그걸 없애는 법.** PIL 로 폭을 재도 CJK 금칙처리·
커닝·자간까지는 재현할 수 없어서 긴 문단에서 한 줄이 갈린다(잔차 1.4cm). **PowerPoint 는
정답을 안다** — `TextFrame2.TextRange.BoundHeight` 가 실제로 그린 텍스트 높이다. 추정을 더
정교하게 만드는 대신 **되먹인다**:

```
python <build>.py                        # 1회차: 추정으로 그린다 (상자 이름에 키를 새긴다)
python measure_boxes.py <out.pptx>       # COM 1회로 BoundHeight 를 캐시에 적는다
python <build>.py                        # 2회차: 캐시를 써서 «정확»하다
python measure_boxes.py <out.pptx> --dry-run    # 검산: 「수렴」이 뜨는지 본다
```

**결정론적으로 2회에 수렴한다**(2026-08-26 teacher 포스터 실측: 1회차 최대 0.64cm·합계
+0.49cm → 2회차 **0.00cm**). 도형 이름이 `pptx_kit.box_key()` 해시라서 COM 이 텍스트를
다시 맞춰볼 필요가 없다 — **이름이 곧 키다**. 캐시는 프로젝트 밖
(`agent/cache/pptx_box_heights.json`)에 살아서 **다음 포스터·덱은 1회차부터 정확**하다.

**표도 캐시에 든다**(2026-08-27). `bullets`·`caption` 뿐 아니라 `ptable`·`dtable` 이 같은
루프를 탄다. 다만 저장되는 값의 «모양»이 다르다 — 상자는 **float(총높이 cm)**, 표는
**list(행별 높이 cm)** 다. 표는 총높이만으로는 행을 배치할 수 없기 때문이다.
검증 사례: brush_cog KSEPI2026 Table 2 가 예측 14.00cm / 실제 **14.73cm**(+0.73)로 캡션이
표의 아래 선을 밟고 있었다 — 되먹인 뒤 2회차 0.00cm. **`audit_text_fit` 은 이걸 못 본다**
(표에는 도형 수준 TextFrame 이 없다). 표 아래에 무언가를 이어 놓는다면 이 루프가 유일한 보증이다.

⚠ **COM 으로 표를 읽을 때 `Rows(r).Height` 가 「전부 0.000」 이면 속성이 아니라 «수명»을
의심하라.** 반복문 안에서 `app.Quit()` 을 부르면 다음 회차의 `Dispatch` 가 죽은 인스턴스를
잡고, 그 뒤 모든 읽기가 0 이나 「Object does not exist」 로 나온다. 앱은 **하나만 열고 맨
마지막에 닫는다**. 제대로 열면 늘어난 행도 늘어난 값으로 정확히 읽힌다(실측 0.600 → 1.778).

**빌드 끝에 `PK.cache_report()` 를 찍어라.** 「9/9 실측 (100%)」이 아니면 그 미스만큼 여백
보고가 «추정»이고, 그 오차가 열 끝 y 에 그대로 실린다. 이 한 줄이 없으면 다음 세션은 빌드가
찍은 여백을 참으로 믿는다 — 그게 4.3cm 틀렸던 그 사고다.

**`ink_extent.py` 는 여전히 마지막 관문이다.** 캐시가 100% 여도 상자 높이와 «잉크» 끝은
다르다(마지막 상자의 아래 여백만큼). 캐시는 «레이아웃 산술이 참인가», ink_extent 는
«눈에 보이는 여백이 얼마인가» — 다른 질문이다. 실측 예: 빌드 4.1cm vs 잉크 5.5cm.

**여백이 남으면 글을 늘려 메우지 않는다.** 내용을 더 넣거나 그림을 키운다. 포스터는
3cm 이상 남으면 「할 말이 없다」로 읽힌다(`poster_rules` §0·§1).

## 포스터 헤더 — 기관 로고 배지와 QR 묶음 (레시피)

«어디에 둘지»는 판단이라 `poster_rules` §13 에 있다. 여기는 **어떻게 그리는가**다.

**로고는 놓기 전에 대비를 «계산»한다.** 불투명 화소(alpha>200)의 최빈색과 헤더 배경색의
밝기(`0.299R+0.587G+0.114B`) 차가 **30 미만이면 묻힌다**(실측 11/255). 낮으면 로고 뒤에
흰 둥근 배지(`ROUNDED_RECTANGLE`, `adjustments[0]≈0.12`)를 깔고 로고 여백만큼 패딩 —
정사각이 아니면 배지도 «긴 변» 기준으로 정사각화한다.

**QR 은 «폭 고정 셀»에 담고 셀 단위로 정렬한다.** 개수가 바뀌는 묶음이기 때문이다.
```python
CELL_W = QR_SIZE + 2.4                      # QR 은 셀 가운데, 라벨은 셀 «전체 폭»에 가운데 정렬
x = (PW - MARGIN) - len(items) * CELL_W     # 셀 단위로 우측 여백선에 맞춘다
for path, label in items:
    pic(x + (CELL_W - QR_SIZE) / 2, y, QR_SIZE)
    tb(x, y + QR_SIZE + 0.15, CELL_W, lab_h)    # 라벨 상자 = 셀 폭 -> 여백선을 못 넘는다
    x += CELL_W
```
- 블록 폭을 **개수로 계산해 정렬하면** 하나가 빠졌을 때 남은 것이 구석으로 밀린다
  (2026-09-02 실측: 2개용 자리에 1개만 들어가 21cm 가 비었다).
- 라벨 상자를 `QR_SIZE + 여유` 로 잡고 QR 중심에 맞추면 **상자가 여백선을 넘는다.**
  셀 폭을 쓰면 그 문제가 «구조적으로» 사라진다.
- 라벨 폭은 어림하지 말고 `poster_kit.text_width_cm` 으로 재고, 셀을 넘으면 빌드를 세운다.
- **QR PNG 는 배치될 크기 그대로 만들고 `dpi=300` 을 파일에 기록**한다 — 안 그러면
  `audit_font_sizes` 의 [SCALE] 이 「native 보다 축소 배치」로 잡는다(그리고 그게 옳다).

## COM 렌더 뒤에는 산출물이 바뀌어 있을 수 있다 — md5 를 확인하라

`render_pptx.py` 는 PowerPoint 를 띄운다. 그 PowerPoint 가 파일을 저장해 버리는 사고가
**2026-08-26 하루에 두 번** 났다: 표·그림·푸터가 두 벌씩 복제된 pptx 가 만들어졌고
(도형 48개, 같은 좌표 중복 11쌍), **PDF·PNG 는 깨끗했으므로 렌더를 보는 검사로는 안 잡혔다.**
그대로 보냈으면 인쇄물이 겹쳐 나온다.

```
python agent/tools/build_guard.py verify <산출물.pptx>   # 렌더 «직후»에
```

빌드가 `guard()`/`stamp()` 를 쓰고 있으면 이 한 줄이면 된다. 어긋나면 **재빌드**가 정답이다
— 중복 도형을 손으로 지우지 마라(무엇이 지워졌는지 확인할 방법이 없다).

## 표는 «진짜 표»로 만든다 — 텍스트박스로 흉내내지 마라 (덱·포스터 공통)

랩의 덱 빌더들은 표를 셀마다 텍스트박스 + 가로줄 도형으로 그려 왔다(예: 선택교과4
`ppt_common.table_slide`). 이유가 셋인데 **둘은 정당했고 하나는 표류**다:

① **python-pptx 에는 셀 테두리 API 가 없다.** 채우기는 한 줄인데 선은 `a:lnL/R/T/B` XML 을
손으로 써야 한다. 그런데 학술 표의 표준은 **선만 있고 채우기가 없는 것**이라, 하필 못 하는 게
늘 원하는 서식이었다. (`poster_kit.ptable` 은 «테두리를 안 쓰고 줄무늬 채우기»로 피해 갔다 —
포스터엔 통하지만 학술 표에는 안 통한다.)
② 진짜 표는 템플릿 표 스타일을 상속한다 — 띠 색·테마 폰트·테두리가 딸려 온다.
③ **그리고 진짜 이유: 헬퍼가 없었다.** 그래서 덱마다 각자 만들었고, 통제가 쉬운 쪽이
텍스트박스였다.

**`pptx_kit.dtable()` 이 ①을 한 번만 제대로 풀어 놓았으므로 이제 흉내낼 이유가 없다.**
진짜 `add_table` + 학술 서식(가로 3선, 세로선 없음, 채우기 없음) + 행 높이 실측.

**흉내낸 표가 치르는 대가:**
- ⚠ **셀이 줄바꿈되면 아래 행을 «밀지 않고 겹친다».** 모든 행이 고정 높이라서다. 진짜
  표는 PowerPoint 가 그 행을 늘려 밀어낸다(음성대조 실측: 지정 1.20 → 실제 2.54cm).
- **감사가 못 본다.** `shape.has_table` 을 도는 수치 검증기는 가짜 표를 통째로 통과시킨다 —
  가짜 표를 쓰면 «어떤 감사가 적용되는지»가 조용히 바뀐다.
- 손으로 행을 하나 넣으면 아래 전부를 다시 배치해야 한다. 진짜 표는 흐른다.

**표가 늘어났는지는 `measure_boxes.py` 가 잡는다.** ⚠ 비교 대상을 틀리기 쉽다 — PowerPoint 는
행을 늘리면서 «행 높이도 같이» 갱신하므로, 열어 본 뒤에는 행합 == 전체높이라 **언제나
일치한다**. 참인 비교는 **파일에 적힌 값(python-pptx) vs 렌더된 값(COM)** 이다.
⚠ 그리고 두 판본을 맞출 때 **`shape_id` 는 슬라이드 안에서만 고유하다** — 키에 슬라이드
번호를 안 넣으면 덱에서 서로 다른 도형이 같은 칸을 덮어써 거짓 양성이 쏟아진다. 1슬라이드
포스터에서는 절대 안 드러나고 덱에서만 나온다(2026-08-27, 실제로 11건을 오보고했다).

**XML 을 직접 쓰는 함수라 조용히 깨진다.** `a:tcPr` 자식의 스키마 순서를 어기면 PowerPoint 가
복구 모드로 열고, 원치 않는 변에 `noFill` 을 안 쓰면 템플릿 스타일의 세로선이 비쳐 나온다.
둘 다 `selftest.py::test_dtable` 이 지킨다 — `dtable` 을 고치면 그걸 돌려라.

## 값 대조는 «누락»을 못 본다 — 수치가 실린 모든 pptx 산출물

포스터 수치를 원본 CSV와 1:1 대조하는 스크립트를 만들어 **32/32 통과**를 받고 "보내도
되나"라고 물었다. 그런데 심사자 감사가 잡은 결함은 하나도 그 검증에 안 걸렸다:

- CSV 에 있는데 표에서 **빠진 행**(탈락자 수, 역치 초과 %) — 값 대조는 «있는 값»만 본다
- **분모가 사라진 것**(`% (n)` 로 바꾸며 파고별 N 을 날림) — 남은 값들은 다 맞았다
- 한 열에 대해 **거짓인 행 라벨**(비교군의 가해자가 「학생·보호자」일 리 없다)
- **자료 창을 넘는 주장**(마지막 파고가 2023 인데 결론이 "after 2023")

**값이 맞는 것과 표가 참인 것은 다른 명제다.** 값 대조 뒤에 반드시 물을 것:
① CSV 에 있는데 안 실은 행이 있나 ② 각 열 라벨이 «그 열»에 대해 참인가
③ 각 주장이 자료가 덮는 기간·집단 안에 있나 ④ 분모가 보이나.

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
   **`python scripts/diff_pptx.py OLD.pptx NEW.pptx` is that "check what changed" step**
   — it also takes `shipped.zip::inner/deck.pptx` when the only baseline you have is
   inside a circulated archive. Run it a second time *after* folding their edits into
   the script: the only differences left should be the ones you meant to introduce.
   ⚠ **Compare paragraphs, not runs.** PowerPoint re-splits every paragraph into
   word-level runs on save, so a run-level diff buries four real edits under ~600
   spurious rows (measured). `diff_pptx.py` joins runs for you — do not hand-roll a
   run-level comparison. Text-box *height* differences after someone merely opened the
   file are usually PowerPoint shrink-wrapping the box, not an edit; the tool flags those.
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

   **The mirror image: a helper that EXPANDS to fill the slide.** A whole-slide
   table helper typically derives its row height from all the space left below
   the header, so it always reaches the bottom. Put a card or a second block
   under it and the two silently overlap — no error, and every coordinate you
   passed is individually correct. Fix the helper, don't nudge the card: add a
   variant that takes an explicit bottom edge (`table_block(y0, y1)` beside
   `table_slide`) so a slide sharing space with a table cannot be built wrong.
   A 2026-08-19 deck hit this on four slides at once.
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
7. **Audit numbers and type.** Cross-check every figure against its source file. The mechanical
   passes live outside this skill, in the lab's shared `agent/tools/`. **Don't run them one by one
   from memory — that is how they get forgotten** (these scripts were missing from this file
   entirely until 2026-08-10). One entry point runs every check that applies to the artifact and
   reports what it skipped and why:

       python agent/tools/audit.py FILE.pptx      # also .docx .hwpx .md .tex
       python agent/tools/audit.py --list         # the whole genre x check map

   It dispatches on **genre, not file extension** — a conference poster, a lab-meeting deck and a
   talk deck are all `.pptx` but want different font floors and different surface rules (a meeting
   deck *should* print `출처: x.csv`; a poster should not). It prints the genre it inferred and
   why; override with `--genre poster|slide|meeting`.

   It dispatches to `audit_text_fit.py` and `audit_surface_text.py` from this skill plus
   `deck_audit.py` (XML referential integrity), `deck_render_audit.py` (rendered-pixel table
   overflow), `audit_font_sizes.py`, and `audit_table_widths.py`. Exit 1 if any check fails, so a
   build can gate on it. If you are working from the standalone public skill without the lab's
   `agent/tools/`, run this skill's two scripts directly and know that the rest are not covered.
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
| `pptx_kit.py` | Palette-agnostic **mechanics**: `new_deck`, `blank_slide_layout`, `rect`, `slide_number`, **`speaker_note`** (rebuild-proof notes; injects the missing placeholder), `fit_picture` (PIL-measured, overflow-safe image), `overflows` (boundary check), `hang` (hanging indent — python-pptx has no property for it), `text_units`/`wrapped_row_count` (Hangul-aware wrap-length estimate, for pre-sizing a card before drawing it), **`dtable`** (학술 서식의 **진짜 표** — 가로 3선·세로선 없음·채우기 없음, 행 높이는 `wrapped_lines` 로 실측. 표를 텍스트박스로 흉내내지 않게 해주는 함수이니 표가 필요하면 여기부터), `check_surface_leaks`/`save_and_check` (gate a save on caller-supplied banned-phrase hits + overflow — see §11 for why the phrase list is never a shared default). Also carries native-equation builders (`equation_slot`, `promote_equations`, `m_frac`/`m_sub`/`m_sup`/`m_nary`/`m_sqrt`/`m_acc`) — only relevant if a slide needs a real OOXML equation object; see guide §10 before using these. Import it; project style layers on top. |
| `inspect_pptx.py FILE` | Structure + notes + **overflow** dump. First thing to run on any deck. |
| `audit_text_fit.py FILE [--all]` | Asks PowerPoint how big the text really is and flags only text that runs **off-slide** or **onto another text/picture**. Exit 1 on a hit, so a build can gate. **A textbox does not clip — it spills**, and spilling over a background fill is normal layering; treating either as an error makes the check cry wolf (two earlier cuts of this script did exactly that, 3/3 false on a clean deck). |
| `audit_surface_text.py FILE [--max-fig N] [--max-tab N]` | «있으면 안 되는 말»이 남았는지 기계로 훑는다 — 편집 해명·내비게이션 안내·재진술 신호·내부 파일명, 그리고 **번호 drift**(산출물엔 Figure가 3개인데 캡션에 "Figure 9"가 남은 경우). `audit_text_fit.py`가 «글자가 넘치는가»를 본다면 이건 «내용이 표면에 남았는가»를 본다. exit 1이라 빌드 게이트로 쓸 수 있다. 사람 눈으로 훑는 방식은 반복해서 실패한다. |
| `render_pptx.py FILE [--pdf OUT]` | Render to PDF (PowerPoint COM) then PNG per slide, for the §5 visual check. |
| `selftest.py` | Proves `speaker_note` round-trips (write → reopen → read) on a placeholder-less notes master, with no real template. |
| `poster_kit.py` | Palette-agnostic mechanics for **large-format academic posters** (cm-scale canvas, not a 16:9 slide) — `two_col_grid`, `sectitle`, `bullets`/`caption`/`ptable` (all with an enforced minimum legible font size via `kf`), `pic_cm` (width-locked, no silent shrink-below-floor), `min_font_report` (catches text that bypassed `kf`). See guide §12 before building a poster — the 2-column-not-3, one-accent-color, narrative-caption habits it encodes. |
| `poster_kit_selftest.py` | Proves `poster_kit`'s font-floor enforcement and grid math without a real poster project. |
| `audit_image_dpi.py FILE [--min N]` | 그림이 **놓인 크기에 견줘** 충분한 해상도인지 잰다 — `실효 DPI = 가로 픽셀 / 놓인 인치`. 파일 자체는 멀쩡한데 «넓게 놓아서» 흐린 경우를 잡는다(실측: 한 덱에서 네 장은 183~262 인데 한 장만 **121**). 1920px 프로젝터가 13.33in 슬라이드를 쏘면 144 DPI 가 한계라 기본 문턱은 150, 인쇄물이면 `--min 220`. **종횡비 비틀림**도 같이 본다(`width` 와 `height` 를 둘 다 주면 그림이 눌린다). ⚠ 그림 «안»에 구워진 글자가 낡았는지는 못 본다 — 그건 캡처 소스를 따로 대조한다. |
| `web_shot.py URL OUT` | 웹 페이지를 **슬라이드에 놓을 크기에 맞춰** 찍는다(headless Chrome). `--slide-width 7.6 --min-dpi 200` 을 주면 **필요한 device scale factor 를 계산**해서 찍는다 — 눈대중으로 배율을 정하면 그림마다 DPI 가 제각각이 된다. 좌표(`--crop`, `--box`)는 **CSS px** 이라 배율을 바꿔도 그대로 유효하다(뷰포트가 고정이므로 레이아웃이 안 변한다). 찍힌 폭이 기대와 다르면 «멈춘다» — 배율만 올리고 원본을 다시 안 찍으면 엉뚱한 데를 조용히 자른다. ⚠ 페이지가 개편되면 좌표가 어긋나므로 **결과 PNG 를 반드시 눈으로** 볼 것. |
| `ink_extent.py FILE [--cols ...]` | 렌더에서 **열별 «실제» 잉크 끝**을 잰다. 빌드가 찍는 여백은 상자 기준이라 실제와 다를 수 있다(2026-08-26: 0.2 vs 4.3cm). `.pptx` 를 주면 캔버스를 읽고 렌더를 찾거나 만들고, 머리글·푸터 «띠»를 자동 배제한다. 「여백이 남았나/빡빡한가/두 열이 균형인가」는 이걸로 판정한다 — `audit_text_fit.py`(겹침·이탈)와 축이 다르다. |
| `measure_boxes.py FILE [--dry-run]` | **PowerPoint 가 실제로 그린 텍스트 높이**(`BoundHeight`)를 COM 1회로 걷어 캐시(`agent/cache/pptx_box_heights.json`)에 적는다. 킷이 도형 이름에 새긴 `pk:<해시>` 가 키라 텍스트를 다시 맞출 필요가 없다. 빌드→측정→빌드로 **오차 0 에 결정론적 2회 수렴**. `--dry-run` 은 캐시를 안 쓰고 오차만 본다(검산용). **표는 «행별» 높이를 캐시에 담고**(값이 리스트), 별도로 «행 자동확장»도 검사한다 — 늘어난 표가 있으면 exit 1(빌더가 받은 높이가 거짓이고 그 아래가 이미 겹쳤다는 뜻). 읽기 전용(`ReadOnly=True`)이지만 뒤에 `build_guard.py verify` 를 권한다. |

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
