<div align="center">

# 📊 PPTX Editing Skill

**LLM 에이전트가 PowerPoint(`.pptx`) 덱을 안 깨고, 손편집을 안 지우고 편집하게 해주는 Agent Skill.**

재빌드가 손편집·발표자 노트를 지우는 문제, 이미지/텍스트 오버플로, 신뢰 안 되는 autofit 같은 실전
함정을 검증된 규칙과 실제로 돌아가는 Python 도구로 묶었습니다. Claude Code · Codex · Cursor ·
Gemini CLI 등에서 그대로 씁니다.

_A portable Agent Skill that teaches AI coding agents to build and edit PowerPoint (`.pptx`) decks
without losing hand-edited work on rebuild, overflowing slide boundaries, or silently dropping
speaker notes._

[![CI](https://github.com/kangdacool/pptx-editing-skill/actions/workflows/ci.yml/badge.svg)](https://github.com/kangdacool/pptx-editing-skill/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue)
![Agent Skill](https://img.shields.io/badge/format-SKILL.md-8A2BE2)
![Works with](https://img.shields.io/badge/agents-Claude%20Code%20%C2%B7%20Codex%20%C2%B7%20Cursor%20%C2%B7%20Gemini-orange)

**[한국어](#한국어)** · **[English](#english)**

</div>

---

## 한국어

### 왜 필요한가

`python-pptx`는 대부분을 커버하지만, 물리지 않는 부분들이 있습니다:

- **재빌드는 손편집을 지운다.** 빌더가 `.pptx`를 통째로 새로 쓰면 PowerPoint에서 직접 넣은 노트·
  이동한 상자·타이핑한 숫자가 전부 사라집니다.
- **발표자 노트가 조용히 실패한다.** 템플릿의 notes master에 body placeholder가 없으면
  `notes_text_frame`이 `None`이라 `.text` 대입이 에러 — "이 템플릿은 노트를 못 쓴다"로 오판하기 쉽지만,
  실제로는 PowerPoint에서 정상적으로 노트를 씁니다. placeholder를 주입하면 됩니다.
- **넘침을 안 알려준다.** 이미지·텍스트가 슬라이드 경계를 넘어도 python-pptx는 에러를 안 냅니다.
- **autofit을 못 믿는다.** 텍스트 상자 자동축소는 렌더러마다 다르게 동작합니다 — 렌더링해서 눈으로
  확인하는 것 외엔 방법이 없습니다.

이 스킬은 이 함정들을 막는 재사용 가능한 메커닉(`pptx_kit.py`)과, 왜/어떻게를 설명하는 가이드로
구성됩니다.

### 빠른 시작

Claude Code 등 Agent Skill을 지원하는 도구에서 `skills/pptx-editing/`를 스킬 디렉토리로 등록하면
자동으로 트리거됩니다 — `.pptx` 파일이나 "파워포인트", "발표자료", "슬라이드 고쳐줘" 같은 요청에
반응합니다.

```bash
# 어떤 덱이든 먼저 구조를 확인
python skills/pptx-editing/scripts/inspect_pptx.py deck.pptx

# 만든 뒤 렌더링해서 시각 확인 (PowerPoint COM 필요, Windows)
python skills/pptx-editing/scripts/render_pptx.py deck.pptx
```

### 안에 뭐가 있나

- **`SKILL.md`** — 에이전트가 매번 읽는 핵심 규칙(항상 로드되므로 의도적으로 짧게 유지: 재빌드
  손편집 보호, 노트 함정, 오버플로 방지, 워크플로 8단계 — 숫자·서식 감사에 더해 슬라이드 순서
  감사까지).
- **`references/pptx-guide.md`** — §1–§12으로 나뉜 상세 가이드(좌표계, 노트, 재빌드/`build_guard`,
  이미지, 텍스트, 표, 렌더링, 데이터 무결성, 학습용 PPT, 네이티브 OOXML 수식, surface-leak 게이트,
  **학술 포스터**) — 필요한 절만 그때 읽음.
- **`scripts/pptx_kit.py`** — 팔레트 무관 메커닉: `new_deck`, `speaker_note`(재빌드에도 안전한 노트),
  `fit_picture`(PIL 실측 기반 오버플로 안전 이미지), `overflows`(경계 검사), 네이티브 수식 빌더,
  `check_surface_leaks`/`save_and_check`(금칙어·오버플로 게이트).
- **`scripts/poster_kit.py`** — **대형 캔버스 학술 포스터** 전용 메커닉(2단 그리드, 최소 글자 크기
  강제, 폭 고정 이미지 배치, `min_font_report`) — 16:9 슬라이드가 아닌 cm 단위 벽보용.
  `ptable`이 돌려주는 높이는 `행높이 × 행수`가 아니라 **셀이 실제로 줄바꿈된 것을 반영한 참값**이라,
  표 바로 밑에 놓은 캡션이 렌더에서 마지막 행에 겹치지 않는다.
- **`scripts/measure_boxes.py`** — **PowerPoint가 실제로 그린 텍스트 높이를 걷어 캐시에 적는다.**
  상자 높이를 «추정»하면 그 오차가 여백에 실려 「빌드→넘침→줄임→빌드」 루프가 시작된다. PIL 실측도
  근사다(CJK 금칙처리·커닝을 재현할 수 없다). COM의 `TextFrame2.TextRange.BoundHeight`를 되먹이면
  2회차 빌드부터 오차 0으로 수렴하고, 캐시는 프로젝트를 넘어 남는다. 늘어난 표가 있으면 exit 1.
- **`scripts/ink_extent.py`** — 렌더된 PNG에서 «잉크가 실제로 닿은 범위»를 잰다. 좌표·XML 검사가
  통과시킨 넘침을 픽셀에서 잡는 마지막 관문.
- **`scripts/diff_pptx.py`** — **덱 두 개가 무엇이 다른가.** 누군가 손으로 고친 덱을 덮어쓰기
  «전에» 무엇이 바뀌었는지 보고, 그 수정을 빌드 스크립트에 옮긴 «뒤에» 제대로 옮겨졌는지
  다시 확인한다. 기준본이 회람 zip 안에 있으면 `shipped.zip::inner/deck.pptx` 로 준다.
  ⚠ **문단 단위로 비교한다 — run 단위로 하면 안 된다.** PowerPoint 는 저장할 때 문단을
  낱말 단위 run 으로 다시 쪼개므로, 아무도 안 건드린 문단이 수십 건의 «변경»으로 나온다
  (실측: 진짜 수정 4건이 헛것 600여 건에 묻혔다).
- **`scripts/inspect_pptx.py`** — 슬라이드별 구조·텍스트·노트·오버플로 덤프.
- **`scripts/render_pptx.py`** — PowerPoint COM으로 PDF→PNG 렌더링(시각 QA용).
  ⚠ `--pdf`만 주면 PNG는 만들지 않는다. 예전에는 원본 옆에 `<덱>_png/`를 «늘» 만들어서,
  압축해 배포하는 폴더에서 부르면 그 PNG가 그대로 딸려 나갔다.
- **`scripts/audit_text_fit.py`** — PowerPoint COM으로 실측한 텍스트 크기 기준 오버플로/겹침 감사.
- **`scripts/audit_surface_text.py`** — 편집 해명·내비게이션 안내·재진술·Figure/Table 번호 drift를
  기계로 훑는다. 빌드 게이트로 쓸 수 있게 exit 1.
- **`scripts/audit_image_dpi.py`** — **그림이 «놓인 크기»에 견줘 충분히 선명한가.**
  `실효 DPI = 가로 픽셀 / 놓인 인치`. 파일은 멀쩡한데 넓게 놓아서 흐린 경우를 잡는다 —
  한 장씩 보면 다 그럴듯하고 덱 전체를 세워 놓고 견줘야 뒤처지는 장이 나온다
  (실측: 183·189·210·243·262 사이에 혼자 121). 종횡비 비틀림도 함께 본다.
- **`scripts/web_shot.py`** — 웹 페이지를 **슬라이드에 놓을 크기에 맞춰** 찍는다(headless Chrome).
  `--slide-width 7.6 --min-dpi 200` 을 주면 필요한 device scale factor 를 «계산»한다.
  좌표는 CSS px 이라 배율을 바꿔도 유효하고, 찍힌 폭이 기대와 다르면 멈춘다 — 배율만 올리고
  원본을 다시 안 찍으면 엉뚱한 데를 조용히 자르기 때문이다.
- **`scripts/deck_mono.py`** — 발표덱용 «흑백» 스타일 한 벌(`cover`·`section`·`content`·
  `bullets`·`card`·`table`·`code`·`picture`). 기계 규격은 `pptx_kit` 이 지고 여기엔 색과 문법만
  둔다. 덱에 색을 쓰지 않는 이유는 **논문 원본 그림이 원색으로 들어오기 때문**이다 — 덱에도
  색이 있으면 충돌하고 청중이 색을 «의미»로 읽는다. 강조는 한 색뿐이고 덱당 한둘만 쓴다.
  영문 덱이면 `set_font("Arial")` — 모듈 전역을 직접 바꾸면 일부 헬퍼만 따라와 더 나빠진다.
- **`scripts/selftest.py`** / **`scripts/poster_kit_selftest.py`** / **`scripts/deck_mono_selftest.py`**
  / **`scripts/audit_surface_text_selftest.py`** — 실제 템플릿 없이 각각 `speaker_note` 의
  왕복(쓰기→저장→재열기→읽기), `poster_kit` 의 폰트 최솟값·그리드 계산, `deck_mono` 의 글꼴
  전파·그림자 제거, 표면 문구 검사기의 정밀도를 증명.

### 라이선스

MIT — [LICENSE](LICENSE) 참고.

---

## English

### Why this exists

`python-pptx` covers most of the format, but the parts that bite are the ones it doesn't surface
cleanly:

- **A rebuild wipes hand edits.** If a builder script regenerates the whole `.pptx`, anything added
  by hand in PowerPoint — notes, a moved box, a typed number — disappears.
- **Speaker notes fail silently.** If a template's notes master has no body placeholder,
  `notes_text_frame` is `None` and `.text` raises — easy to misread as "this template can't hold
  notes," when PowerPoint handles notes on it just fine. The fix is injecting the missing
  placeholder.
- **No overflow warning.** Images or text can extend past the slide boundary with no error from
  python-pptx.
- **No reliable autofit.** Text-box auto-shrink behaves differently across renderers — rendering to
  an image and looking is the only reliable check.

This skill packages reusable mechanics (`pptx_kit.py`) that guard against these failure modes, plus
a guide explaining the why/how.

### Quickstart

Point any Agent Skill-aware tool (Claude Code, etc.) at `skills/pptx-editing/` and it triggers
automatically on `.pptx` files or requests like "edit this deck" / "build a PowerPoint."

```bash
# Inspect any deck's structure first
python skills/pptx-editing/scripts/inspect_pptx.py deck.pptx

# After building, render for a visual check (needs PowerPoint COM, Windows)
python skills/pptx-editing/scripts/render_pptx.py deck.pptx
```

### What's inside

- **`SKILL.md`** — the core rules an agent reads every time (deliberately kept short since it's
  always loaded: rebuild-safety, the notes gotcha, overflow prevention, an 8-step workflow — a
  narrative-order audit alongside the numbers/formatting audits).
- **`references/pptx-guide.md`** — detailed guide in §1–§12 (coordinates, notes, rebuild/
  `build_guard`, images, text, tables, rendering, data integrity, teaching-deck norms, native OOXML
  equations, surface-leak gating, **academic posters**) — read only the section that's relevant.
- **`scripts/pptx_kit.py`** — palette-agnostic mechanics: `new_deck`, `speaker_note` (rebuild-proof
  notes), `fit_picture` (PIL-measured, overflow-safe image placement), `overflows` (boundary check),
  native-equation builders, `check_surface_leaks`/`save_and_check` (banned-phrase + overflow gate).
- **`scripts/poster_kit.py`** — mechanics for **large-format academic posters** — 2-column grid,
  enforced minimum legible font size, width-locked image placement, `min_font_report` — a cm-scale
  wall poster, not a 16:9 slide. The height `ptable` returns is the true sum over rows whose cells
  actually wrap, not `row_height × n_rows`, so a caption placed under the table does not end up
  overlapping its last row at render time.
- **`scripts/inspect_pptx.py`** — per-slide structure/text/notes/overflow dump.
- **`scripts/render_pptx.py`** — PDF→PNG rendering via PowerPoint COM, for visual QA.
- **`scripts/audit_text_fit.py`** — overflow/overlap audit using PowerPoint-measured text size.
- **`scripts/audit_surface_text.py`** — machine sweep for editing-process leakage, navigation asides,
  restated text, and figure/table number drift. Exits 1, so it can gate a build.
- **`scripts/audit_image_dpi.py`** — **is each picture sharp enough for the size it is placed at?**
  `effective DPI = pixel width / placed inches`. Catches images that are fine as files but soft on
  the slide because they were placed wide. One at a time they all look plausible; only lining the
  whole deck up shows the outlier (measured: 121 among 183·189·210·243·262). Also flags squashed
  aspect ratios (what you get by setting both `width` and `height`).
- **`scripts/web_shot.py`** — capture a web page **at the size it will occupy on the slide**
  (headless Chrome). Give it `--slide-width 7.6 --min-dpi 200` and it *computes* the device scale
  factor. Crop/box coordinates are in CSS px, so they stay valid when the scale changes; if the shot
  does not come out at the expected width it stops, because raising the scale without re-shooting
  silently crops the wrong region.
- **`scripts/deck_mono.py`** — one black-and-white deck style (`cover`, `section`, `content`,
  `bullets`, `card`, `table`, `code`, `picture`). `pptx_kit` carries the mechanics; this file holds
  only colour and grammar. The deck stays monochrome because **the figures you paste into a talk
  arrive in full colour** — a coloured deck fights them, and the audience starts reading the deck's
  colour as meaning. One accent colour, once or twice per deck. For an English deck call
  `set_font("Arial")`; rebinding the module global directly leaves half the helpers behind.
- **`scripts/selftest.py`** / **`scripts/poster_kit_selftest.py`** / **`scripts/deck_mono_selftest.py`**
  / **`scripts/audit_surface_text_selftest.py`** — prove `speaker_note`'s round-trip (write → save →
  reopen → read), `poster_kit`'s font-floor/grid math, `deck_mono`'s font propagation and shadow
  removal, and the surface-text auditor's precision, with no real template or project needed.

### License

MIT — see [LICENSE](LICENSE).
