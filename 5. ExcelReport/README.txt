DataGuard v0.9.3 - apply_template_style 인자 불일치 해결

원인:
report_generator.py가 apply_template_style(report, template, inspection)으로 3개 인자를 보내지만,
설치된 template_style_engine.py의 구버전은 2개만 허용해 TypeError가 발생했습니다.

설치:
1. Gradio 종료(Ctrl+C), 생성된 보완결과서 Excel도 닫기
2. 폴더 내 Python 파일 3개를
   C:\Users\USER\Desktop\최가윤\data-guard-agent\5. ExcelReport\
   위치에 덮어쓰기 (기존 파일 백업 권장)
3. 현재 규칙 JSON, 검사결과서, 템플릿, 보고서 데이터는 건드리지 않습니다.
4. 프로젝트 루트에서 python .\dataguard_gradio_ui.py 실행

포함:
 - report_generator.py: 2개 또는 3개 인자 양쪽 스타일 엔진에 호환되도록 수정
 - template_style_engine.py: 현재 3개 인자 버전
 - rule_sheet_copy_exact.py: 검사결과서에서 규칙 시트 전체 복사(Windows Excel 및 pywin32 필요)

Windows + Microsoft Excel, python -m pip install pywin32가 필요합니다.
본 버전은 스타일 함수 호출의 TypeError를 수정했으며,
022번 검사결과서에 대한 전체 Excel COM 생성 과정은 Windows에서 최종 확인해야 합니다.
