"""No Ollama/server needed. Run: python -m unittest -v test_case_rag.py"""
import csv
import tempfile
import unittest
from pathlib import Path

from case_rag_engine import (
    bootstrap, connect, load_tests as get_samples, retrieve, make_prompt, register_verified_case,
    run_comparison, status_counts, save_csv,
)
from excel_export import export_csv_to_xlsx

class CaseRAGTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.db = Path(self.temp.name) / "lab.sqlite"
        self.conn = connect(self.db)
        bootstrap(self.conn)
        self.cases = get_samples()
        self.s02 = next(x for x in self.cases if x["case_id"] == "S02")

    def tearDown(self):
        self.conn.close()
        self.temp.cleanup()

    def test_unverified_six_never_retrieved_in_real_mode(self):
        counts = status_counts(self.conn)
        self.assertEqual(counts["전체"], 7)
        self.assertEqual(counts["실제 승인 검색 사례"], 0)
        self.assertEqual(retrieve(self.conn, self.s02, demo=False), [])

    def test_synthetic_demo_not_real_approval(self):
        matches = retrieve(self.conn, self.s02, demo=True)
        self.assertTrue(matches)
        self.assertEqual(matches[0]["source_kind"], "가상 데모")
        self.assertEqual(matches[0]["case_id"], "DEMO-TOPLEVEL-KEY")

    def test_no_ground_truth_leak(self):
        matches = retrieve(self.conn, self.s02, demo=True)
        prompt = make_prompt(self.s02, matches, demo=True)
        self.assertNotIn(self.s02["expected_after"], prompt)
        self.assertIn("현재 적용 구문규칙", prompt)
        self.assertIn("가상 데모", prompt)
        self.assertIn("$.note", prompt)

    def test_manual_approval_requires_all_evidence(self):
        values = {
            "project_id":"019", "filename":"test.json", "error_type":"정의되지 않은 Key",
            "target_path":"$.foo", "reported_error":"최상위 foo 키 존재",
            "rule_text":"최상위 foo는 미정의", "approved_after":"최상위 foo 제거 필요",
            "evidence_reference":"검사결과서 규칙 9행 + 원본 JSON + 승인문서 번호", "reviewer":"품질담당A",
        }
        with self.assertRaises(ValueError):
            register_verified_case(self.conn, values, {})
        ident = register_verified_case(self.conn, values, {
            "source_checked":True, "rule_checked":True, "after_approved":True,
        })
        self.assertTrue(ident.startswith("REAL-"))
        self.assertEqual(status_counts(self.conn)["실제 승인 검색 사례"], 1)
        self.assertTrue(retrieve(self.conn, self.s02, demo=False))

    def test_only_baseline_when_no_approved_cases(self):
        sent = []
        def fake(prompt, model):
            sent.append(prompt)
            return {"model":model, "suggestion":"담당자 확인 필요", "reason":"규칙 확인",
                    "missing_evidence":"원본", "elapsed_sec":1.0}
        result = run_comparison(self.conn, self.s02, "fake-model", demo=False, infer=fake)
        self.assertEqual(len(sent), 1)
        self.assertFalse(result["rag"])
        self.assertEqual(result["record"]["B · Ollama 제안(Case RAG)"], "")
        self.assertNotIn(self.s02["expected_after"], sent[0])

    def test_two_real_calls_with_demo_retrieval(self):
        sent = []
        def fake(prompt, model):
            sent.append(prompt)
            return {"model":model, "suggestion":"최상위 미정의 키를 제거하도록 수정 필요",
                    "reason":"구문규칙 우선", "elapsed_sec":0.1}
        result = run_comparison(self.conn, self.s02, "fake-model", demo=True, infer=fake)
        self.assertEqual(len(sent), 2)
        self.assertIn("가상 데모", sent[1])
        self.assertNotIn(self.s02["expected_after"], sent[1])
        output = save_csv(result["record"], Path(self.temp.name))
        with output.open(encoding="utf-8-sig", newline="") as f:
            row = next(csv.DictReader(f))
        self.assertEqual(row["RAG 검색 모드"], "가상 데모")
        self.assertEqual(row["B · Ollama 제안(Case RAG)"], "최상위 미정의 키를 제거하도록 수정 필요")
        xlsx = export_csv_to_xlsx(output)
        self.assertTrue(Path(xlsx).is_file())
        import zipfile
        with zipfile.ZipFile(xlsx) as z:
            self.assertIsNone(z.testzip())
            self.assertIn("xl/worksheets/sheet2.xml", z.namelist())

    def test_database_is_idempotent(self):
        bootstrap(self.conn)
        self.assertEqual(status_counts(self.conn)["전체"], 7)

if __name__ == '__main__':
    unittest.main()
