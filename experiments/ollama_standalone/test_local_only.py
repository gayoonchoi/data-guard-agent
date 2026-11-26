import unittest
from unittest.mock import patch
import ollama_test_ui as app

class LocalOnlySmokeTests(unittest.TestCase):
    def test_sample_csv_has_no_duplicates(self):
        cases = app.load_cases()
        self.assertEqual(len(cases), len({x['case_id'] for x in cases}))
        self.assertEqual(len(cases), 6)

    def test_expected_answer_not_passed_to_model(self):
        for case in app.load_cases():
            self.assertNotIn(case['expected_after'], app.payload_text(case))

    def test_ollama_call_is_local_and_records_real_response(self):
        case = app.load_cases()[0]
        captured=[]
        def pretend(url, data=None, timeout=8):
            captured.append((url, data))
            return {'model':'qwen2.5:7b-instruct', 'message':{'content':'{"suggestion":"유사 백태로 수정 필요","reason":"id 매핑 규칙 확인","human_review_required":true,"missing_evidence":""}'},'eval_count':34}
        with patch.object(app,'request_json',side_effect=pretend):
            res=app.call_ollama(case,'qwen2.5:7b-instruct')
        self.assertEqual(captured[0][0], 'http://127.0.0.1:11434/api/chat')
        self.assertFalse(captured[0][1]['think'])
        self.assertFalse(captured[0][1]['stream'])
        self.assertEqual(res['suggestion'],'유사 백태로 수정 필요')
        self.assertNotIn(case['expected_after'], captured[0][1]['messages'][1]['content'])

    def test_unknown_source_not_marked_deleted(self):
        case=app.load_cases()[2]
        self.assertIn('삭제 승인 기록은 제공되지 않음',app.payload_text(case))
        self.assertIn('단정하지',case['expected_after'])


class HumanEvaluationTests(unittest.TestCase):
    def test_new_results_default_to_review_needed(self):
        case = app.load_cases()[0]
        record = app.record_row(case, {"suggestion": "유사 백태로 수정", "reason": "규칙", "model_response": "qwen2.5:7b-instruct"})
        self.assertEqual(record[app.EVAL_COL], '추가 검토 필요')
        self.assertEqual(record[app.EVAL_REVIEWER_COL], '미평가')

    def test_manual_review_persists_without_mutating_original(self):
        import tempfile
        import csv
        from pathlib import Path
        case = app.load_cases()[0]
        old_dir = app.LOG_DIR
        try:
            with tempfile.TemporaryDirectory() as folder:
                app.LOG_DIR = Path(folder)
                original = app.save_results([app.record_row(case, {"suggestion": "유사 백태", "reason": "id=14", "model_response": "qwen2.5:7b-instruct"})])
                original_bytes = Path(original).read_bytes()
                status, output, _ = app.save_human_review(original, 'S01', '적절', 'category_id=14 규칙 일치')
                self.assertIn('평가 저장', status)
                self.assertEqual(Path(original).read_bytes(), original_bytes)
                _, data = app.read_review_records(output)
                self.assertEqual(data[0][app.EVAL_COL], '적절')
                self.assertEqual(data[0][app.EVAL_REASON_COL], 'category_id=14 규칙 일치')
                self.assertEqual(data[0][app.EVAL_REVIEWER_COL], '담당자 수동 평가')
        finally:
            app.LOG_DIR = old_dir

    def test_cannot_grade_without_reason(self):
        import tempfile
        from pathlib import Path
        case = app.load_cases()[0]
        old_dir = app.LOG_DIR
        try:
            with tempfile.TemporaryDirectory() as folder:
                app.LOG_DIR = Path(folder)
                original = app.save_results([app.record_row(case, {"suggestion": "값 수정", "model_response": "qwen2.5:7b-instruct"})])
                status, output, _ = app.save_human_review(original, 'S01', '부적절', '')
                self.assertIn('실패', status)
                self.assertIsNone(output)
        finally:
            app.LOG_DIR = old_dir



class ExcelExplanationTests(unittest.TestCase):
    def test_workbook_has_results_and_notes_and_ground_truth(self):
        import csv
        import tempfile
        from pathlib import Path
        from zipfile import ZipFile
        from excel_export import export_csv_to_xlsx, GROUND_TRUTH_COL, LEGACY_NAMES
        with tempfile.TemporaryDirectory() as folder:
            csv_file = Path(folder) / 'legacy.csv'
            with csv_file.open('w', encoding='utf-8-sig', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=['case_id', '모범답안_비교용', 'Ollama 제안', '오류'])
                writer.writeheader()
                writer.writerow({'case_id':'S02', '모범답안_비교용':'최상위 text 삭제', 'Ollama 제안':'최상위 text 추가', '오류':''})
            original_bytes = csv_file.read_bytes()
            result = Path(export_csv_to_xlsx(csv_file))
            self.assertEqual(csv_file.read_bytes(), original_bytes)
            with ZipFile(result) as z:
                self.assertIsNone(z.testzip())
                wb = z.read('xl/workbook.xml').decode('utf-8')
                data = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
                notes = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
                self.assertIn('테스트 결과', wb)
                self.assertIn('비고', wb)
                self.assertIn(GROUND_TRUTH_COL, data)
                self.assertNotIn('모범답안_비교용</t>', data)
                self.assertIn('평가 주체', notes)
                self.assertIn('참고용 잠정 내용 평가', notes)
                self.assertIn('모델을 학습시키지도 않음', notes)

    def test_legacy_csv_review_uses_new_column_names(self):
        import csv
        import tempfile
        from pathlib import Path
        from excel_export import GROUND_TRUTH_COL, KEYWORD_COL, RUN_ERROR_COL
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'old.csv'
            with path.open('w', encoding='utf-8-sig', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=['case_id', 'Ollama 제안', '모범답안_비교용', '자동키워드확인', '오류'])
                writer.writeheader()
                writer.writerow({'case_id':'S01', 'Ollama 제안':'확인 필요', '모범답안_비교용':'정답', '자동키워드확인':'충족', '오류':''})
            headers, records = app.read_review_records(path)
            for header in [GROUND_TRUTH_COL, KEYWORD_COL, RUN_ERROR_COL]:
                self.assertIn(header, headers)
                self.assertIn(header, records[0])


if __name__ == '__main__':
    unittest.main(verbosity=2)
