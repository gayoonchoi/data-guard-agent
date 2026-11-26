# DataGuard Agent — Rule Knowledge Base v0.1

## 무엇을 만들었나

기존 4개 + 신규 019번 **검사결과서**에서 `구축기업 구문규칙`과 `검사 구문규칙`을 **셀 값 단위로 비교**하고, 일치하는 규칙을 프로젝트·데이터 그룹 단위의 **외부 JSON 지식 사전**으로 추출했습니다. 원본 데이터를 변경하지 않습니다. 규칙은 최초 사용에 앞서 담당자가 재확인해야 합니다.

| 프로젝트 | 규칙 그룹 | 규칙 항목 행 수 | 실제 그룹명 |
|---|---:|---:|---|
| 007 | 5 | 261 | D1_개인, D2_시군당해, D3_시군차년, D4_2024추가, D5_생산비 |
| 013 | 1 | 29 | 카자흐스탄 모국어 사용자의 한국어 음성데이터 |
| 017 | 1 | 45 | 터널 결함 이미지 데이터셋 |
| 022 | 9 | 18 | 9개 LLM 학습데이터 그룹(각 2개 규칙 항목) |
| 019 | 2 | 21 | JSON 16개 항목, TXT 5개 항목 (JSON 규칙 K15 불일치) |
| **총계** | **18** | **374** | **019 JSON 구축기업/검사 규칙의 K15에 차이 존재** |

**374는 검사 항목 정의가 들어 있는 엑셀 행 수**이며, 추가 서술형·교차 컬럼 조건까지 모두 센 '검증 로직 수'는 아닙니다.

## 파일 구조

```text
DataGuard_RuleKB_v01/
  catalog.json             # 프로젝트/그룹명/파일 위치/원본 SHA-256
  rule_dictionary.json     # 프로젝트·그룹·항목경로별 단일 Python dict 형태
  audit.json               # 그룹별 비교 결과, 수동 해석 필요 항목
  rule_index.csv           # 353개 항목의 Excel용 색인 (UTF-8 BOM)
  folder_mapping.template.json # 실제 데이터 폴더 → 그룹 매핑의 빈 템플릿
  rules/
    007/group_01.json ~ group_05.json
    013/group_01.json
    017/group_01.json
    022/group_01.json ~ group_09.json
  extract_rules.py         # 검사결과서 XLSX -> 지식 사전 재생성
  validate_dataset.py      # 1차 읽기 전용 Rule Engine PoC
  test_rule_kb.py          # 가상 데이터로 검증하는 유닛 테스트
```

## 각 규칙의 구조

원본의 `타입`, `필수여부`, `유효값`, `최소값`, `최대값`, `패턴`, `기타`, `null 허용여부`, `설명`, `비고`를 **별도 필드와 `original_settings`에 보존**했습니다.

```json
{
  "rule_id": "007-G1-003",
  "field": "Age",
  "path": "$.Age",
  "data_type": "integer",
  "required": false,
  "nullable": false,
  "minimum": 0,
  "maximum": 99
}
```

중첩 JSON은 `$.Segmentation.eojeols[*].text`, `$.annotations[*].category_id`와 같은 경로로 저장합니다. `source`에는 원본 시트명, 행 번호, 셀 좌표가 남아 있습니다. 값이 없는 '필수여부'나 'null 허용 여부'는 안전하게 `null`(미정)로 처리합니다.

## v0.2 — 신규 검사결과서 한 개 추가하기 (019 사례)

`extract_rules.py`의 새 버전에는 `--input-file` 옵션과 **기존 사전 병합**이 추가되었습니다. 기존 사업 007/013/017/022는 유지하고 신규 사업 019의 JSON/TXT 규칙 2그룹만 추가합니다. `catalog.json`, `audit.json`, `rule_dictionary.json`, `rule_index.csv`, `folder_mapping.template.json`은 갱신되며, 수정 전 파일은 `_backups/`에 보존됩니다. 입력 검사결과서 원본은 변경하지 않습니다.

```powershell
Set-Location "C:\Users\USER\Desktop\최가윤\data-guard-agent\DataGuard_RuleKB_v01"
python .\extract_rules.py --input-file "C:\Users\USER\Desktop\최가윤\019. 국토안전관리원 스마트 안전장비 실증현장 이동형 CCTV 기반 AI 이벤트 데이터\0.보완자료\검사결과서(구문)\(019)구문적정확성_검사결과서_국토안전관리원 스마트 안전장비 실증현장 이동형 CCTV 기반 AI 이벤트_261011.xlsx" --output-dir "."
python .\test_rule_kb.py
```

`--input-dir`도 사용 가능하지만 **경로가 아니라 폴더 전체**를 의미합니다. `--output-dir "."`은 현재 위치한 `DataGuard_RuleKB_v01`에 저장하는 뜻입니다. 다른 경로에서 실행하려면 `--output-dir`에 해당 폴더 전체 경로를 기입하세요. 기존 4개는 병합 모드에서 자동 보존되며, `--rebuild`는 필요한 경우에만 선택하세요.

**주의 — 019 JSON 규칙 K15 불일치:** `imagePath`의 구축기업 규칙은 `$d{4}.jpg`, 검사 규칙은 `^\d{5}.jpg$`입니다. 원문을 모두 보존했고, 충돌하는 정규식 검사는 담당자 확인 전까지 비활성화했습니다. TXT 그룹은 서로 일치합니다. JSON `shapes[*].text`는 규칙에 있으나 최상위 `$.text`는 규칙에 없습니다. **현재 읽기 전용 검증기는 규칙 외 추가 키를 검출하는 금지 키 검사까지 구현하지 않았습니다.** TXT 행 단위 검사 및 `w/h > 0`, 좌표축별 상한 검증도 별도 구현이 필요합니다.

## 1. 사전 생성 또는 갱신 방법

Python 3.10+ 권장. **외부 패키지가 필요하지 않습니다.** 여러 개 `.xlsx`를 한 폴더에 둔 다음 아래 명령으로 추출합니다.

```powershell
python extract_rules.py --input-dir "C:\검사결과서" --output-dir "C:\DataGuard\rule_kb"
```

이 스크립트는 **구문 규칙/검사 종합 현황 시트만 읽습니다**. 수십만 행의 오류 목록을 전수 로드하지 않으며, 엑셀 원본을 수정하지 않습니다. 새 규칙이 입력되면 재실행하여 통합 사전과 `audit.json`, `rule_index.csv` 등을 갱신할 수 있습니다.

## 2. 규칙 조회

전체 사전을 단일 파일로 불러오려면 다음과 같이 조회합니다.

```python
import json
from pathlib import Path

kb = json.loads(Path('rule_dictionary.json').read_text(encoding='utf-8'))
age = kb['projects']['007']['groups']['1']['rules_by_path']['$.Age']
print(age['minimum'], age['maximum'])  # 0 99
```


```python
import json
from pathlib import Path

rule_set = json.loads(Path('rules/007/group_01.json').read_text(encoding='utf-8'))
print(rule_set['dataset_group_name'])
for rule in rule_set['rules'][:5]:
    print(rule['path'], rule['data_type'], rule['required'], rule['allowed_values'])
```

**규칙 자동 적용 시 폴더명으로 D1/D2 등을 추측하지 않습니다.** 사용자가 원본 데이터 폴더와 규칙 그룹의 매핑을 직접 정의해야 합니다. `folder_mapping.template.json`에 실제 프로젝트의 상대 경로를 기록한 후 연결합니다. 파일 내부의 품목별 규칙 적용 범위는 비고에 보존하되, 아직 자동 필터링하지 않습니다.

## 3. 읽기 전용 Rule Engine PoC 실행

**기존 데이터는 절대 자동 수정되지 않습니다.** TXT 파일 검사는 아직 지원하지 않으며 JSON/CSV/JSONL의 규칙 해석 중 일부는 별도 확인이 필요합니다. CSV/JSON/JSONL의 자료형, 존재 여부, null 여부(명시된 경우), 유효값, 수치 범위, 정규식을 검사합니다. 배열의 길이/문자열 길이는 자료형에 따라 해석합니다. Python 표준 라이브러리만 필요합니다.

```powershell
python validate_dataset.py --kb "." --project 007 --group 1 --input-path "C:\검증데이터\D1_개인" --output-dir "C:\검증결과"
```

출력은 `issues_007_G1.jsonl`과 `summary.json`입니다. 구문 규칙에 기재되지 않은 관계 규칙은 임의로 추가하지 않습니다. 022 메시지 배열은 **최상위 JSON 배열** 형식일 때에만 명시된 메시지 수·역할 순서·공백 본문 조건을 1차 검사합니다. 실제 JSON의 외부 래핑 구조가 확인되면 입력 어댑터를 조정해야 합니다.

```powershell
python test_rule_kb.py
```

위 테스트는 가상 데이터만 사용하므로 회사의 실제 데이터를 변경하거나 외부 서비스로 전송하지 않습니다.

## 주의: 아직 보완결과서 생성 단계는 아님

1. **007 기후 데이터:** 그룹(D1~D5)별 다른 규칙을 유지했습니다. `Sex`, `Age`, `Year`처럼 같은 항목이라도 그룹마다 다릅니다. 품목별 적용 범위는 `비고`에 있으며, 실제 품목별 파일명/폴더 매핑이 필요합니다. 일부 `코드값 시트 참조`는 현재 검사결과서에서 원본 코드표를 확인할 수 없어 별도 참조표가 필요합니다. `Career ≤ Age` 같은 **교차 컬럼 조건은 구문규칙 표에 명시된 조건인지 확인한 뒤 별도 승인된 규칙으로 등록**해야 합니다.
2. **013 음성 JSON:** 배열/객체 계층, 어절의 `text`, `start`, `end` 등 구조화했습니다. `text` 끝의 불필요한 공백을 실제로 수정하려면 별도 세부 보완 규칙/승인이 필요합니다.
3. **017 터널 JSON:** `category_id`와 `category_name` 각각의 유효값 및 비고에 적힌 ID-명칭 대응관계를 보존했습니다. **ID↔이름 매핑의 교차 필드 검증 및 자동 보완은 아직 구현하지 않았습니다.** ID 14는 비고에 `유사 백태`로 설명되어 있습니다.
4. **022 LLM 데이터:** 9개 그룹별 `role`과 `content` 외에도 `role` 비고의 **메시지 수·순서·본문 공백 금지**를 `group_level_rules`로 추출했습니다. `content` 정규식은 역할별(`user.content`, `assistant.content`)로 제한하며, 원문에 포함된 따옴표처럼 해석이 필요한 패턴은 자동 수정/강제 검증하지 않습니다.
5. 4개 검사결과서는 **규칙을 정의하는 자료**이지 **원본 데이터 파일 자체가 아닙니다**. 원본 CSV/JSON 파일 및 기존 보완결과서 양식이 있어야 `보완 전/후` 생성과 정확도 평가가 가능합니다.

**구현 순서 권장:** ① 규칙 사전 확정 → ② 그룹/품목별 매핑 등록 → ③ 읽기 전용 검증 → ④ 오류 유형별 승인 가능한 보완 함수 → ⑤ Human Approval + 재검증 → ⑥ 기존 TTA 양식 Excel Writer → ⑦ 예외사항 GraphRAG/Agent 연결.
