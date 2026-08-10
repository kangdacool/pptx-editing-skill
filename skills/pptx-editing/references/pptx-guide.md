# PPTX Editing Guide (python-pptx)

## 흔한 실패 TOP (skim first)

1. **"이 템플릿은 발표자 노트를 못 쓴다"** — 거짓. notes master에 body placeholder가
   없어 `notes_text_frame`이 `None`일 뿐. placeholder를 주입하면 됨 (§2). 기본이 안 되는
   것처럼 보이면 **포맷이 아니라 내 가정을 의심**한다.
2. **손으로 넣은 노트/편집이 재빌드에 사라진다** — 빌더는 pptx를 통째로 새로 만든다.
   노트·수치·미세조정은 **스크립트에** 넣는다 (§2, §3).
3. **이미지가 슬라이드 밖으로 삐져나온다** — 종횡비 무시하고 크기 지정. PIL로 실제 픽셀을
   재고 박스에 맞춘 뒤 **경계 초과를 검사**한다 (§4).
4. **카드/텍스트가 잘린다** — python-pptx엔 신뢰할 autofit이 없다. 박스를 넉넉히 잡고
   **렌더해서 눈으로 본다** (§5, §7). 구조검사는 잘림을 못 본다.
5. **build_guard가 저장을 막는다** — on-disk 파일이 최종 빌드 이후 바뀐 것. 손편집일 수
   있으니 **백업하고 확인한 뒤** `.build-md5` 삭제 (§3).
6. **Windows 콘솔에서 `print`가 UnicodeEncodeError** — em-dash·한글 등. stdout을 utf-8로
   감싼다 (§7).
7. **수치를 기억으로 타이핑** — 금지. 소스 파일(csv/xlsx/md)에서 읽는다 (§8).

---

## §1 python-pptx 모델 · 좌표 · 템플릿/레이아웃

- 단위는 EMU. `from pptx.util import Inches, Pt` — 좌표·크기는 `Inches(...)`, 폰트는 `Pt(...)`.
- 슬라이드는 **레이아웃**에서 생성: `prs.slides.add_slide(layout)`. 템플릿의 테마/마스터를
  물려받으려면 그 템플릿을 열어서 쓴다.
- **빈 레이아웃 함정**: 표준 인덱스 6이 아닌 템플릿이 많다(예: 'DEFAULT' 하나만 노출).
  placeholder가 0개인 첫 레이아웃을 찾는다 → `pptx_kit.blank_slide_layout`.
- **슬라이드 삭제 API가 없다**: 템플릿의 기존 슬라이드를 지우려면 `sldIdLst`에서 요소를
  제거하고 관계(rId)를 drop한다 → `pptx_kit.new_deck`이 처리(테마는 유지, 슬라이드만 비움).

```python
from pptx_kit import new_deck, blank_slide_layout
prs = new_deck("template.pptx")          # 테마 상속 + 슬라이드 비움
s = prs.slides.add_slide(blank_slide_layout(prs))
```

## §2 발표자 노트 (가장 자주 틀리는 곳)

`slide.notes_slide.notes_text_frame.text = "..."`는 **notes master에 body placeholder가
있을 때만** 동작한다. 없으면 `notes_text_frame`이 `None`이라 `.text`가 AttributeError.
이때 **"불가능"이 아니라** notes 슬라이드 spTree에 placeholder를 직접 주입한다.

```python
from pptx_kit import speaker_note
speaker_note(slide, "청중에게 말할 내용.\n두 번째 문단.")   # 줄바꿈 = 문단
```

- **노트는 빌드 스크립트에 둔다.** 스크립트에서 쓴 노트는 재빌드마다 재생성된다.
  PowerPoint에 손으로 타이핑한 노트는 **다음 재빌드가 지운다**(§3). 사용자에게 그렇게 안내.
- 검사: `inspect_pptx.py`가 슬라이드별 노트 유무·길이를 찍는다. 왕복 증명은 `selftest.py`.

## §3 재빌드가 손편집을 지운다 · build_guard · 날짜 스탬프

- 빌더는 `prs.save(out)`으로 **파일 전체를 새로 쓴다** → PowerPoint에서 손으로 넣은 노트·
  이동한 상자·타이핑한 수치는 전부 사라진다. **모든 것을 스크립트에** 넣는 게 유일한 방어.
- `build_guard` (agent/tools): 저장 직전 md5 지문을 기록(`<out>.build-md5`). 다음 빌드에서
  on-disk 파일이 그 지문과 다르면 저장을 막는다 = **누군가 손댔을 수 있다는 신호**.
  - 복구: 파일을 백업 → 무엇이 바뀌었는지 확인 → 정말 덮어써도 되면 `.build-md5` 삭제 후 재빌드.
  - on-disk mtime이 `.build-md5`보다 **나중**이면 특히 조심(PowerPoint에서 열어 저장했을 수 있음).
- 산출물은 `YYMMDD` 스탬프. `OP_STAMP`/`OUT_STAMP` 같은 env로 기존 빈티지를 덮어쓸 수 있게 한다.

## §4 이미지 — 넘침 없는 배치

python-pptx는 넣은 그림이 슬라이드를 벗어나는지 **말해주지 않는다**. 항상:

```python
from pptx_kit import fit_picture, overflows
fit_picture(slide, "fig.png", left=0.5, top=1.6, max_w=7.5, max_h=4.6)  # PIL로 실측·종횡비 유지
bad = overflows(slide, sw=13.333, sh=7.5)   # 경계 넘친 shape 목록 → 배포 전 0이어야
```

- 실제 픽셀은 `PIL.Image.open(path).size`로 잰다(파일이 존재하는지 먼저 확인).
- 박스에 맞추되 종횡비 유지: 폭 기준 축소 후 높이가 넘치면 높이 기준으로 다시.
- 배치 후 `bottom = top + h < slide_height`, `right = left + w < slide_width` 확인.

## §5 텍스트 — 신뢰할 autofit이 없다

- `text_frame`의 자동 축소(MSO auto-size)는 렌더러/버전마다 다르게 동작 → **의존하지 않는다.**
- 상자를 **넉넉히** 잡고, 줄 수·한글 폭(한글 글리프≈라틴의 2배)을 대략 계산해 높이를 준다 —
  손으로 어림하지 말고 `pptx_kit.text_units`/`wrapped_row_count`로 계산할 것(§11).
- 두 줄 이상 wrap되는 불릿은 `pptx_kit.hang(paragraph, width)`로 hanging indent를 줘야 wrap된
  줄이 마커 밑이 아니라 왼쪽 여백으로 도로 빠지는 걸 막는다(python-pptx엔 이 속성이 없음).
- 잘림은 오직 **렌더로만** 보인다(§7). 카드·각주·긴 제목은 렌더 확인 필수.
- **폰트 하한(ppt_rules)**: 한글 제목바 30–32pt, 표 헤더 17pt, 본문 16pt, 보조 15pt,
  **13pt 미만 금지**. `audit_font_sizes.py`로 점검.

## §6 표 — 소스에서 채운다, 손타이핑 금지

- 원고/보고서 표는 **`agent/tools/manuscript_table.py::add_journal_table()`**로 소스 CSV에서
  읽어 만든다. 값을 손으로 옮기지 않는다(재실행 한 번이 전수 갱신이자 검증).
- 간단 회귀표는 `header_row()`+`build_row()` 패턴(CLAUDE.md)으로 조립.
- 셀 폭 합 = 표 폭. 넘치면 `audit_table_widths.py`가 잡는다.

## §7 렌더링 · 시각 QA · 콘솔 Unicode

- **렌더 없이는 신뢰하지 않는다.** `render_pptx.py`가 PowerPoint COM으로 PDF를 뽑고
  슬라이드별 PNG로 변환한다(win32com). PNG를 실제로 **열어서 본다** — 잘림·넘침·정렬.
- COM이 없으면 LibreOffice `soffice --headless --convert-to pdf`가 대안(레이아웃은 근사).
- **Windows 콘솔 Unicode**: `print`에 em-dash·한글이 있으면 cp949 인코딩 에러. 스크립트
  상단에서 stdout을 감싼다:
  ```python
  import sys, io
  if hasattr(sys.stdout, "buffer"):
      sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
  ```

## §8 데이터 무결성 (NON-NEGOTIABLE)

- **슬라이드의 모든 수치는 소스 파일(csv/xlsx/md)에서 읽은 값**이어야 한다. 기억으로
  "대략 1.38" 같은 건 날조. 소스를 못 찾으면 사용자에게 알리고 멈춘다.
- 슬라이드 코드를 쓰기 **전에** 그 슬라이드가 참조할 표를 읽는다. 값 추출 → 코드 작성 순서.
- 생성 후 **수치 감사**: 슬라이드별로 소스와 1:1 대조(별도 단계, 생략 금지).
- 서사가 숫자와 일치하는지도 감사한다("여성에서 더 강함"은 OR이 같으면 틀림).

## §9 학습(교육용) PPT ≠ 학술발표 PPT

이 스킬 문서는 원래 학술발표/연구brief 덱을 염두에 두고 쓰였다. 강의·실습용(학습) PPT는
성격이 달라서 같은 규칙을 그대로 적용하면 안 되는 지점이 있다.

- **모든 내용이 "다 보여주기 위한" 게 아니다.** 학술발표 덱은 전부 청중에게 보여줄 내용이지만,
  학습 PPT는 흔히 "과제 설명" 파일과 "정답 공개" 파일이 분리돼 있다. 같은 비교/기준값이라도
  어느 파일에 넣을지가 다르다:
  - 과제 **정의 자체**에 필요한 비교(예: "두 후보 도구 중 뭐가 나은지 겨뤄라"는 게 과제 그 자체)는
    학생용 슬라이드에 명시해야 완결된 지시가 된다.
  - 사후 해석·보너스 통찰(예: "네가 계산한 값이 데모 값보다 높더라")은 정답 공개 파일에서만
    드러나야 한다 — 과제 설명에 미리 넣으면 답을 흘리는 셈이다.
  슬라이드 문구 하나를 고칠 때마다 "이건 과제 정의인가, 정답 힌트인가"부터 구분해서 넣을 파일을
  판단할 것. (2026-08-05 세션에서 실제로 한 번 정답 힌트를 과제 슬라이드에 잘못 넣었다가 되돌린 사례 있음.)
- **개념 설명은 청중이 아니라 학생 기준으로 자명성을 낮춰야 한다.** 학술발표라면 생략할 정의
  ("r이 정확히 뭘 뜻하는지", "CVI가 무슨 약자인지")를 학습 PPT에서는 명시해야 한다. 공식은
  네이티브 수식(OOXML Math, §10)으로 엄밀하게 보여주되, 값의 범위·용어 설명을 바로 옆에 나란히 붙인다.
- **연습문제 설계 자체가 콘텐츠다.** 학술 덱은 이미 끝난 분석 결과를 보여주지만, 학습 PPT
  뒤의 데이터/코드 설계는 그 자체로 학생에게 가르치는 방법론이다 — 비교 기준이 통계적으로
  부적절하면(검증 안 된 잣대를 "정답" 자리에 쓰거나, 문항 수가 다른 도구를 대등하게 겨루게
  하는 등) 단순 오탈자보다 심각한 문제다. 문구를 다듬기 전에 "이 비교·설계가 방법론적으로
  정당한가"부터 검증할 것 — 필요하면 실제 데이터로 다시 계산해서 확인하고, 절대 숫자를 추측해서
  채우지 말 것.

## §10 네이티브 수식(OOXML Math) — 다루기, 처음부터 만들기, 검사 시 함정

**Only read this section when a slide actually needs a real PowerPoint equation object**
(Insert > Equation) — plain formula-look text in a monospace textbox (`kit_common.formula()`)
does not need any of this.

### §10a 기존 네이티브 수식을 다룰 때 (mc:AlternateContent)

python-pptx로 PowerPoint 네이티브 수식(Insert > Equation)을 삽입하면
`<mc:AlternateContent><mc:Choice Requires="a14">...<m:oMath>...</m:oMath>...</mc:Choice>
<mc:Fallback>...</mc:Fallback></mc:AlternateContent>` 구조가 생기는데, 이게 여러 함정을 만든다:

- **`slide.shapes` 순회로는 안 보인다.** python-pptx의 Shapes 컬렉션은
  `p:sp`/`p:pic`/`p:graphicFrame`/`p:grpSp`/`p:cxnSp`만 인식하고 `mc:AlternateContent`로
  감싸인 도형은 건너뛴다. `shape.text_frame` 기반 진단(예: "oMath 태그가 몇 개인지 세기")은
  **수식 도형에 대해 항상 0을 반환한다** — 실제로는 존재해도. 수식 도형을 찾거나 셀 때는
  반드시 **저장된 zip의 raw XML을 정규식으로 검색**할 것
  (`re.findall(r"<p:sp>.*?</p:sp>", xml, re.DOTALL)` 후 각 블록에서 `oMath` 포함 여부 확인).
- **이 때문에 "기존 도형 정리 후 새로 추가" 패턴이 깨진다.** `clear_slide_extra_shapes()`류
  헬퍼가 `slide.shapes`를 순회해서 지우는 방식이면 수식 도형은 안 보이니 안 지워지고, 재실행할
  때마다 같은 슬라이드에 수식이 중복으로 쌓인다(실제로 2026-08-05 세션에서 발생 — 두 번의 수식
  패치가 겹쳐 같은 슬라이드에 동일한 수식이 두 개 존재했다). 수식을 교체할 때는 raw XML에서 대상
  도형의 `id=` 속성으로 옛 `mc:AlternateContent` 블록을 찾아 직접 제거해야 한다.
- **`<p:sld>` 루트에 `mc:Ignorable="a14"`가 없으면** 일부 뷰어에서 Choice와 Fallback이 동시에
  렌더링될 수 있다. lxml 객체 모델로는 네임스페이스 맵을 사후에 못 바꾸므로, 이것도 저장된
  zip의 `<p:sld ...>` 시작 태그에 `xmlns:mc`/`xmlns:a14`/`mc:Ignorable="a14"`를 문자열 치환으로
  주입해야 한다.
- **수식 삽입/교체 후에는 반드시 raw XML로 재검증**: 슬라이드마다 `AlternateContent` 개수, 각
  블록의 `id=`, 그리고 `m:t` 텍스트 런을 이어붙여 실제 수식 내용을 재확인할 것 —
  `slide.shapes` 기반 검증은 이 부분에서 신뢰할 수 없다.

### §10b 네이티브 수식을 처음부터 만들 때 (기존 수식이 아직 없을 때)

위 절은 **이미 있는** 수식을 찾거나 교체하는 이야기다. 슬라이드에 수식이 **아예 처음** 들어가야
하면(예: 학습 PPT에 새 개념 슬라이드를 추가할 때) 완전히 다른 문제다 — python-pptx에는 수식을
만드는 API가 없다.

**"PowerPoint COM엔 있겠지"는 검증 없이 믿지 말 것 (2026-08-05 확인).** VBA 기억으로
`TextRange2.OMaths.Add(...)` 나 `MathZones.Add(...)` 같은 게 있을 거라 짐작하고 세 번 시도해 매번
실패했다. `win32com.client.gencache.EnsureModule()`로 실제 typelib을 생성해 까보니 —
`TextRange2.MathZones(Start, Length)`는 있지만 이건 **기존 수식 구간을 읽는 메서드**였다
(`Method 'MathZones' returns object of type 'TextRange2'`, 인자가 Start/Length). **새 수식을
만드는 멤버는 이 typelib에 아예 없다.** COM으로 수식을 넣는 유일한 길은 SendKeys로 UI를 흉내내는
것 뿐이고, 그건 이 문제에 안 맞는다. → **"이 API가 있을 것"이라는 기억은 확인하기 전엔 설계에
반영하지 말 것.** `gencache.EnsureModule` + 생성된 `.py` 모듈을 grep해서 실제 멤버 목록을 보는
게 30초면 된다(SKILL.md "The rule that would have saved the most time"의 반대쪽 함정 — 여기선
"안 될 것"이 아니라 "될 것"이라는 가정이 틀렸다).

**실제로 되는 방법 — placeholder + 저장 후 문자열 치환:**
1. 덱을 만드는 동안, 수식이 들어갈 자리에 **빈 textbox**를 이름표(`EQN::marker`)와 함께 둔다
   (`pptx_kit.equation_slot`).
2. `prs.save(OUTPUT)` 다음, 저장된 파일을 zip으로 다시 열어 그 이름표를 가진 `<p:sp>`를
   실제 `<mc:AlternateContent>`(Choice=진짜 `m:oMath`, Fallback=평문)로 문자열 치환한다
   (`pptx_kit.promote_equations`).
3. **Fallback은 렌더된 이미지가 아니어도 된다** — 평문 한 줄로 충분하다. Fallback은 a14
   확장을 지원 안 하는 옛 뷰어에게만 보이므로, "수식을 이미지로 렌더해서 넣어야 하나"라는
   훨씬 어려운 문제를 통째로 피해간다.
4. 분수·아래첨자·위첨자·Σ·제곱근·바 액센트(x̄)는 `pptx_kit.m_frac`/`m_sub`/`m_sup`/`m_nary`/
   `m_sqrt`/`m_acc`로 조립한다(기존 손편집 수식의 실제 구조를 그대로 본뜬 빌더 — Pearson r의
   전체 공식(합·괄호·바 액센트·제곱·제곱근 중첩)을 이 6개만으로 재조립해 렌더까지 확인함,
   2026-08-05). 행렬·다단 정렬(`m:eqArr`)처럼 이 6개로 안 되는 구조가 필요하면 실제 파일에서
   해당 구조를 raw XML로 추출해 붙여넣을 것.
5. **기존 수식을 재사용하려고 raw XML을 그대로 뽑아 쓸 때 함정:** 추출 정규식
   `<m:oMath[^>]*>(.*)</m:oMath>` 는 `[^>]*`가 태그 이름 경계를 안 봐서 **`<m:oMathPara ...>`
   (바깥 wrapper)까지 매칭**해버린다 — 그 결과로 만든 수식은 안에 `oMathPara`가 한 겹 더 중첩된
   채 저장된다(겉보기엔 슬라이드가 만들어지고 오류도 안 나서, 렌더해서 직접 봐야 드러난다).
   `<m:oMath(?=[\s>])[^>]*>` 처럼 lookahead로 "다음 글자가 공백 또는 `>`"를 강제해야 안전하다.
6. 만든 뒤 반드시 **렌더링해서 눈으로 확인** — raw XML 개수 대조(`mc:AlternateContent`·
   `oMathPara`·`oMath` 각각 몇 개여야 하는지 미리 계산해두고 대조)까지 하면 위 5번류 함정을
   렌더 전에 잡을 수 있다.

이 패턴 전체(빌더 함수 5개 + `equation_slot` + `promote_equations`)는 `pptx_kit.py`에
있다 — 새로 만들지 말고 import.

### §10c 자체 겹침 검사기를 PyMuPDF로 짤 때 — 블록이 잘못 합쳐진다

`render_pptx.py`로 뽑은 PDF를 `fitz`(PyMuPDF)의 `page.get_text("blocks")`로 열어 텍스트
블록 좌표를 서로 비교해 "겹침"을 자동 검사하는 스크립트를 직접 짠다면, PyMuPDF의 블록
클러스터링이 **물리적으로 떨어진 두 텍스트를 하나의 블록으로 합칠 수 있다** — 예: 오른쪽
패널의 마지막 불릿과, 패널 밖 하단에 있는 결론 문구가 위치상 가깝다는 이유로 한 블록으로
묶이면, 그 (부풀려진) 가짜 블록이 다른 열의 텍스트와 겹친다고 오판정한다. 실제로 한
프로젝트에서 문단 하나를 수정한 뒤 이 검사가 "겹침"을 보고했는데, **같은 PDF**를
`page.get_pixmap()`으로 직접 렌더링해 눈으로 보면 완전히 깨끗했다.

**교훈**: 이런 자동 겹침/오버플로 검사가 뭔가를 flag하면, 그 검사가 열었던 것과 **동일한
PDF**를 `get_pixmap()`으로 직접 이미지로 뽑아보기 전엔 진짜 결함으로 단정하지 말 것 —
PowerPoint COM으로 별도 캡처해도 되지만 렌더 엔진이 다르면(soffice vs PowerPoint) 결과가
달라질 수 있으니, 가장 확실한 증거는 **검사기가 실제로 검사한 그 PDF**를 그대로 이미지로
보는 것이다. 이 실수를 피하고 싶으면 애초에 PyMuPDF 블록 비교 대신 `deck_render_audit.py`의
가장자리 픽셀-밀도 방식(`agent/tools/`)을 쓰는 것도 대안이다 — 그건 블록 병합 문제 자체가 없다.

## §11 surface-leak 게이트(`check_surface_leaks`/`save_and_check`) — 금칙어 목록은 항상 호출자가 준다

편집 흔적("이전 버전"), 교차문서 참조, AI 티(수사적 메타발언)를 슬라이드 텍스트에서 검출하는
`check_surface_leaks(prs, terms)`는 **`terms`에 기본값이 없다.** 프로젝트마다 하나의 공유 금칙어
목록을 두지 않는다 — 무엇이 "누출"인지는 산출물의 성격(학술발표 vs 학습 PPT vs 교육자료)마다
다르고, 넓은 공유 목록은 **오탐을 만든다**(실사고: 정상적인 문맥의 "다음과 같다"가 금칙어 목록에
걸려 잘못 flag된 적이 있음). 각 프로젝트는 `output_surface.md`의 "편집자 vs 청중" 레지스터
테스트로 **자기 목록을 직접 만들어** 넘긴다:

```python
from pptx_kit import save_and_check
MY_LEAK_TERMS = ["이전 버전", "수정했습니다", "다음 슬라이드에서"]  # 이 프로젝트만의 목록
save_and_check(prs, OUT, leak_terms=MY_LEAK_TERMS)   # 저장 + 검사, 걸리면 SystemExit
```

`save_and_check`는 오버플로도 같이 검사한다 — 이때 `overflows()`를 슬라이드별로 재사용하므로(§4),
"오버플로"의 정의가 검사기마다 다르게 갈릴 위험이 없다(과거 프로젝트별 사본은 여유 마진이 1인치
vs 0.02인치로 서로 달랐다 — 하나로 합침).

`hang(paragraph, width)`(§5의 hanging indent)와 `text_units`/`wrapped_row_count`(한글 2폭 가중
줄바꿈 추정, 카드 높이를 그리기 전에 미리 계산할 때 씀)도 이 파일에 있다 — 프로젝트마다 따로
구현하지 말고 import.

## §12 학술 포스터(대형 캔버스) — `poster_kit.py`

포스터는 슬라이드덱과 물리적으로 다른 매체다: cm 단위의 대형 캔버스(보통 90×120~140cm), 3피트
밖에서 읽는다는 최소 가독 크기 제약, 국영문 혼용 줄바꿈 추정. `pptx_kit`이 아니라
`poster_kit.py`(같은 스킬 `scripts/`)를 쓸 것 — 2026-08-10에 `KWCS/teacher_violence_qol` 프로젝트의
2단 영문 포스터가 같은 학회 제출용 다른 프로젝트의 3단 포스터보다 명백히 나은 품질로 판정된 뒤,
그 스크립트의 메커닉을 일반화해 뽑아냈다. 한 가지 트릭이 아니라 다섯 가지 습관의 차이였다:

1. **3단이 아니라 2단.** 3단은 그림을 작게 만들고 줄 길이를 짧게 강제해 멀리서 보면 빽빽해
   보인다. `two_col_grid(page_w, margin, gutter)`를 기본으로 쓰고, 3단은 내용이 정말 2단에 안
   들어갈 때만(그리고 그건 보통 "그림이 너무 많다"는 신호이지 레이아웃 문제가 아니다) 예외로.
2. **강제되는 최소 글자 크기 하나, 예외 없이.** `kf()`가 모든 run에서 assert한다(기본 24pt) —
   포스터는 1.5m 밖에서 읽으므로 "이 캡션만 좀 작게"가 실제로 실측 시 20pt 이하로 굳어지는
   사고를 만든다. 장식용 텍스트(저작권 표시 등)만 `min_pt=None`으로 명시적으로 예외 처리.
3. **그림은 배치 크기 그대로 저장, 사후 축소 금지.** `pic_cm(sl, path, l, t, w_cm)`는 열 너비에
   폭을 고정하고 실제 종횡비(PIL)로 높이를 계산한다 — **`max_h`로 축소하지 않는다.** 그림 안에
   박힌 축 라벨·범례의 pt 크기가 축소와 함께 눈에 안 띄게 최소 크기 밑으로 내려간다. 그림이 남는
   공간보다 크면 배치 스크립트에서 그림 자체를 다른 종횡비로 다시 저장하거나 절 순서를
   바꿀 것 — 여기서 눌러 줄이지 말 것.
4. **서술형 캡션, 라벨 캡션 아님.** `caption()`은 2-3문장을 넉넉히 받는다. "Figure 1. X vs Y"는
   포스터에서 본문을 안 읽는 행인이 유일하게 얻어가는 자리를 낭비하는 것 — 그림이 보여주는
   **의미**를 초록 마지막 문장 쓰듯 쓸 것.
5. **강조색 하나, 아주 드물게.** 섹션색(BAR)·본문색(INK)·캡션색(GRAY) 외에는 절마다 가장 중요한
   콜아웃 하나에만 강조색을 쓴다 — 캔버스 하나에 4가지 다른 톤의 카드 배경색이 경쟁하게 하지
   말 것. `bullets()`/`ptable()`의 `accent` 파라미터는 정확히 이 용도 — 절마다 한 줄/한 행에만.

```python
def _load_kit(name):
    import importlib.util, os, sys
    p = os.path.expanduser(os.path.join("~", ".claude", "skills", "pptx-editing", "scripts", name))
    spec = importlib.util.spec_from_file_location(name[:-3], p)
    mod = importlib.util.module_from_spec(spec); sys.modules[name[:-3]] = mod
    spec.loader.exec_module(mod); return mod

KIT = _load_kit("pptx_kit.py")      # overflows() 등 슬라이드 공통 메커닉
PK = _load_kit("poster_kit.py")     # sectitle/bullets/caption/ptable/pic_cm/two_col_grid/kf
```

**출고 전 체크리스트** (§7 렌더링 QA에 추가로):
- 캔버스에 "TBD"/"to be confirmed"류 자리표시자 문구가 하나도 없을 것 — 포스터는 초안이 아니라
  완성물이다. 실제로 프로젝트 다른 곳에서 이미 확정된 정보(공저자 소속)가 포스터에는 "확인
  필요" 문구로 두 번의 빌드 주기 동안 그대로 남아 있던 사고가 있었다 — 빌드 스크립트를 고칠 때
  프로젝트의 최신 확정 정보를 다시 대조할 것.
- `PK.min_font_report(slide, min_pt=24)`가 빈 리스트를 반환할 것 — `kf()`를 거치지 않고 만들어진
  run(표 셀을 손으로 스타일링한 경우 등)까지 잡아내는 안전망.
- `KIT.overflows(slide, sw_in, sh_in)` — **포스터 실제 크기(예: 90×120cm → 35.4×47.2in)를 넘겨야
  한다.** 기본값(13.333×7.5, 16:9 슬라이드)을 그대로 두면 무조건 통과해 버려 아무것도 못 잡는다.
- `render_pptx.py`로 렌더한 뒤 화면 꽉 채우기가 아니라 **실제 크기 비율(축소 배율을 계산해서
  "3피트 거리에서 이 정도로 보인다")로 확대/축소해 눈으로 볼 것** — 풀스크린으로 보면 이 킷이
  막으려는 "가까이서만 보이는 작은 글자" 문제 자체가 안 보인다.

## 구조 원칙 (요약)

- **메커닉 ↔ 스타일 분리**: 팔레트-무관 메커닉은 `pptx_kit`(이 스킬), 프로젝트 팔레트·
  컴포넌트(cover/section/content/card/table)는 프로젝트 `kit_common.py`. 중복 재구현 금지 —
  이게 노트 함정이 프로젝트마다 재발한 원인이었다.
- 빌드 스크립트는 자기 위치에서 `ROOT`를 계산(`Path(__file__).resolve().parents[N]`), 절대경로
  하드코딩 금지. 재빌드로 이식성 검증.
