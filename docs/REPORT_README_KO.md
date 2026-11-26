# DataGuard Agent | 보완결과서 자동 생성 v0.6 (019 시범)

## 설치 (기존 DataGuard 프로젝트 보존)
1. Gradio를 Ctrl+C로 종료합니다.
2. 이 ZIP의 `data-guard-agent` 안 파일을 현재 프로젝트의 같은 폴더 위치에 복사합니다.
   - `dataguard_gradio_ui.py` : 기존 파일 교체
   - `5. ExcelReport/report_generator.py` : 신규
   - `5. ExcelReport/report_generator_local.py` : 신규
   - `5. ExcelReport/보완결과서_표준양식.xlsx` : 신규 (제공하신 원본 양식)
   - `DataGuard_RuleKB_v01`, `DataGuard_RuleKB_UserBuilt` 폴더는 변경·삭제하지 않습니다.
3. Windows PowerShell:
   ```powershell
   cd "C:\Users\USER\Desktop\최가윤\data-guard-agent"
   python -m pip install openpyxl
   python .\dataguard_gradio_ui.py
   ```
4. `③ 보완결과서 생성` 탭에서 실제 (019) 검사결과서 업로드 → `보완 차수` 선택 → `보완결과서 초안 생성` 클릭 → 파일 다운로드.

## 데이터 처리 원칙
- Excel의 각 그룹 `구축기업 구문규칙`, `검사 구문규칙` 시트를 그대로 읽어 해당 사업의 규칙 시트에 넣습니다.
- 사용자 제공 보완결과서 표준 양식의 예시 규칙, 예시 오류는 남기지 않습니다.
- 국토안전관리원 019는 JSON·TXT 2그룹으로 구성하여 13개 시트가 생성됩니다.
- 검사결과서의 상세 오류 목록을 가져와 `보완 전(검사 결과)` / `보완 후(필요한 작업 지시)`를 기록합니다.
- 원본 라벨링 JSON/TXT와 실제 재검증 결과가 없으므로 **보완 완료 0건**, **검토 초안**으로 표시합니다.
- 기업/검사 규칙이 불일치하면 양쪽 값을 모두 보존합니다. 019 JSON의 K15는 불일치하여 검토 필요입니다.
- 생성된 파일은 `5. ExcelReport/outputs/`에 저장되고 Gradio에서 다운로드할 수 있습니다.
- 외부 LLM API를 호출하지 않고 로컬에서 처리합니다.

## 구현 범위
- (019) 샘플 실제 검사결과서로 생성·대조 테스트 완료: 2개 그룹, 11,860개 오류 행, 13개 시트.
- 동일한 TTA 시트 구조인 다른 사업도 자동으로 그룹을 찾지만, 실제 업무별 보완 조치·다른 양식은 추가 검증 필요.
- 실물 JSON/TXT 내용 수정, 수정 완료 판정, 원본 기준 재검사는 아직 구현하지 않았습니다.
