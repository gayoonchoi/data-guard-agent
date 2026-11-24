# DataGuard (단계별 클릭 기반 자동화 파이프라인)

**AI 학습데이터 구문 품질 검증 및 보완결과서 작성 지원 PoC**  
`Rule Knowledge Base` · `Deterministic Validation` · `Gradio` · `Excel Reporting`

> **상태:** 로컬 PoC (v1.3 워크플로우 기준). 규칙 등록, 검사결과서 기반 초안 생성, 실제 파일 대조, 보완 전후 검토 기록 반영까지 구현한 단계입니다. **원본 라벨링데이터를 자동 수정하거나 보완 완료를 보증하는 시스템은 아닙니다.**

<img width="991" height="503" alt="image" src="https://github.com/user-attachments/assets/81d93af9-61e5-4873-828f-c83b06784947" />

## 1. 프로젝트 소개

DataGuard Agent는 데이터셋마다 다른 **구문 규칙**, 검사결과서의 **오류 지적 사항**, 실제 **라벨링데이터**를 연결하여 구문 품질 검토와 보완결과서 작성 과정을 지원하는 로컬 도구입니다.

반복적인 검사결과서 확인과 오류 위치 조회를 규칙 기반으로 처리하고, 원본 규칙과 검토 이력을 추적할 수 있도록 구성했습니다. 명확하게 정의된 조건은 Python으로 처리하고, 근거가 부족한 수정은 담당자 검토 대상으로 남깁니다.

## 2. 구현 기능

| Gradio 단계 | 기능 | 산출물 / 동작 |
| --- | --- | --- |
| **① 검사결과서 등록** | Excel의 사업별 구문 규칙을 추출하여 지식 사전에 등록·갱신 | 규칙 사전, 불일치 목록 |
| **② 규칙 승인** | 구축기업 규칙과 검사 규칙의 불일치 확인 및 승인/철회 이력 관리 | 승인 상태 및 이력 |
| **③ 보완결과서 생성** | 검사결과서와 해당 사업의 구문 규칙을 이용해 **독립적인 보완결과서 초안** 생성 | 표준양식 기반 `.xlsx` |
| **④ 실제 데이터 연결** | 로컬 폴더 또는 ZIP에서 지적 파일 매칭, 실제 값 조회, 보완 전후 검토 초안 작성 | 매칭 CSV, 검토 CSV |
| **⑤ 보완 전후 Excel 반영** | **③ 초안 + ④ 검토 CSV**를 결합하여 기존 Excel의 보완 전/후 항목 갱신 | 별도 `.xlsx` (③ 원본 초안 유지) |

### 처리 흐름

```text
검사결과서 Excel ──> ① 규칙 사전 등록 ──> ② 규칙 불일치 검토
       │
       ├──────────> ③ 보완결과서 초안 생성 ────────────┐
       │                                              │
       └──> ④ 로컬 폴더/ZIP 파일 매칭 및 값 검토 ───────┤
                                                      ▼
                                        ⑤ 보완 전후 Excel 반영
                                        (새 파일로 저장)
```

**③은 ④의 결과가 없어도 실행할 수 있습니다.** ⑤는 ③과 ④의 결과를 합치는 별도 작업입니다. 전체 과정이 버튼 하나로 실행되는 구조는 아닙니다.

## 3. 기술 스택

- **언어:** Python
- **화면:** Gradio
- **데이터 처리:** Pandas, JSON/CSV/JSONL 및 ZIP 처리
- **Excel:** openpyxl 및 보고서 생성 모듈, Windows Excel 연동이 필요한 경우 pywin32
- **검증:** 외부 구문 규칙 사전 및 결정론적 검사 로직

LangGraph Multi-Agent와 GraphRAG는 초기 기획에 포함된 확장 후보이며, **현재 ①~⑤ 파이프라인의 필수 구성 요소가 아닙니다.**

## 4. 프로젝트 구조

```text
data-guard-agent/
├── dataguard_gradio_ui.py          # ①~⑤ Gradio 화면
├── DataGuard_RuleKB_v01/          # 규칙 추출·조회·승인·테스트 코드
│   ├── extract_rules.py
│   ├── validate_dataset.py
│   ├── rule_approval.py
│   └── test_*.py
├── DataGuard_RuleKB_UserBuilt/    # 실행 중 생성되는 업무별 사전 (업로드 제외)
├── 1. RuleEngine/
│   ├── data_loader_matcher.py     # ④ 파일 매칭
│   ├── rule_evidence_inspector.py # ④ 보완 전후 검토 근거 추출
│   └── test_*.py
├── 5. ExcelReport/
│   ├── report_generator.py        # ③ 보완결과서 초안
│   ├── rule_sheet_copy_exact.py   # 원본 구문 규칙 시트 처리
│   └── evidence_report_bridge.py  # ⑤ 검토 기록 Excel 반영
├── 2. LangGraph_MultiAgent/       # 향후 검토 영역
├── 3. GraphRAG_Error_Pattern/     # 향후 검토 영역
├── 4. HumanApproval/             # 향후 검토 영역
└── README.md
```

> 파일 구성은 현재 로컬 PoC를 기준으로 합니다. 일부 모듈과 실행 중 생성되는 폴더는 버전에 따라 달라질 수 있습니다. 업무에서 추출한 규칙 JSON과 검사결과서, 표준 Excel 템플릿, 출력 파일은 이 저장소에 포함하지 않습니다.

## 5. 로컬 실행 (Windows 명령 프롬프트)

프로젝트 루트 폴더에서 실행합니다.

```cmd
cd /d "C:\path\to\data-guard-agent"
python -m pip install "gradio>=6,<7" pandas openpyxl
python dataguard_gradio_ui.py
```

브라우저에서 `http://127.0.0.1:7860`에 접속합니다. 실제 Excel 시트의 원본 서식 복사가 필요한 Windows 환경에서는 Microsoft Excel과 `pywin32`가 필요할 수 있습니다.

```cmd
python -m pip install pywin32
```

일부 대용량 JSON/JSONL 처리 경로에서 스트리밍 파서를 사용할 경우 추가로 `ijson`을 설치합니다.

```cmd
python -m pip install ijson
```

**실행 전 준비:** 승인받은 로컬 검사결과서, 해당 사업의 규칙 사전, 라벨링데이터 폴더/ZIP, 보완결과서 표준 Excel 양식을 각 기능에서 요구하는 경로에 준비해야 합니다. 이 저장소는 업무용 원본 데이터와 양식을 배포하지 않습니다.

## 6. 테스트

```cmd
python "DataGuard_RuleKB_v01\test_rule_kb.py"
python "DataGuard_RuleKB_v01\test_rule_approval.py"
python "1. RuleEngine\test_data_loader_matcher.py"
python "1. RuleEngine\test_rule_evidence_inspector.py"
```

테스트 파일이 있는 경우 실행합니다. 테스트 통과는 특정 코드 경로의 동작을 확인한 것이며, 새로운 데이터셋에 대한 전체 정확도를 보장하지는 않습니다.

## 7. 안전성 및 현재 한계

- **원본 보호:** ③과 ⑤는 결과물을 별도 파일로 저장합니다. ④는 파일을 읽고 확인하는 단계이며 원본 라벨링데이터를 수정하지 않습니다.
- **수정 지시와 완료의 구분:** 보완 후에는 규칙에 따른 *수정 요구사항*이 기재될 수 있습니다. 실제 수정 및 재검증 전에는 **보완 완료**로 간주하지 않습니다.
- **규칙 불일치:** 구축기업·검사 규칙이 서로 다르면 승인 전까지 관련 자동 판정을 제한해야 합니다.
- **수동 검토:** 원본에서 값의 정답을 확정할 수 없는 오류는 담당자 판단이 필요합니다.
- **운영 범위:** 로컬 PoC로서 사용자 인증, 다중 사용자 운영, 전사 배포, 모든 포맷의 완전한 자동 보완을 제공하지 않습니다.

## 8. 정보보안 및 공개 범위

이 저장소는 **프로그램 소스코드와 설명 문서**만 공개·공유하기 위한 용도입니다. 다음 자료는 커밋하지 않습니다.

- 고객사·제공기관의 검사결과서, 구문 규칙 원문, 라벨링데이터, 보완결과서
- 사업별로 생성된 Rule Knowledge Base JSON/CSV 및 승인 이력
- API 키, 계정정보, 사내 경로 또는 민감한 로그가 포함된 파일
- 사업 자료를 포함하는 화면 캡처·노트북·압축파일

회사 데이터와 업무용 구문 규칙을 공개/외부 저장소에 반출하려면 사전 승인이 필요합니다. **`.gitignore`는 이미 Git이 추적 중인 파일이나 과거 커밋의 내용을 자동 제거하지 않으므로, `git diff --cached`와 `git ls-files`를 별도로 확인해야 합니다.**

---

**개발 목적:** 데이터 구문 품질 검증과 보완결과서 작성의 반복 작업을 줄이고, 규칙·오류 근거·검토 이력의 추적 가능성을 높이기 위한 개인 기술 PoC입니다.
