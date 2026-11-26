import csv
import tempfile
import json
from unittest.mock import patch
from io import BytesIO
import unittest
from pathlib import Path
from local_correction_agent import enhance_csv, eligible, AFTER, STATUS, ask_ollama

class CorrectionAgentTests(unittest.TestCase):
    def test_only_verified_errors(self):
        row = {"사업ID":"022", "규칙그룹":"7", "매칭상태":"일치",
               STATUS:"현재 규칙 위반", "검사 규칙":r"유효패턴: \S"}
        self.assertTrue(eligible(row)[0])
        row[STATUS] = "현재 규칙 통과"
        self.assertFalse(eligible(row)[0])

    def test_blocks_known_conflict(self):
        row = {"사업ID":"019", "규칙그룹":"1", "검사 프로퍼티":"$.imagePath",
               "매칭상태":"일치", STATUS:"현재 규칙 위반", "검사 규칙":"regex"}
        self.assertFalse(eligible(row)[0])

    def test_csv_output_only_not_source(self):
        headers = ["사업ID", "규칙그룹", "검사 규칙", "매칭상태", STATUS,
                   "검사 프로퍼티", AFTER, "보완 완료 여부"]
        row = {"사업ID":"022", "규칙그룹":"7", "검사 규칙":r"유효패턴: \S",
               "매칭상태":"일치", STATUS:"현재 규칙 위반",
               "검사 프로퍼티":"content", AFTER:"기존 지시", "보완 완료 여부":"미완료"}
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/"evidence.csv"
            with path.open("w", encoding="utf-8-sig", newline="") as f:
                writer=csv.DictWriter(f,fieldnames=headers); writer.writeheader(); writer.writerow(row)
            before=path.read_bytes()
            called=[]
            def fake(row, model, include_before=False):
                called.append(row)
                return {"suggestion":"공백이 아닌 실제 내용을 채우도록 확인", "reason":"유효패턴"}
            result=enhance_csv(path, limit=5, infer=fake)
            self.assertEqual(before,path.read_bytes())
            self.assertEqual(result["generated"],1)
            self.assertEqual(len(called),1)
            with open(result["output"], encoding="utf-8-sig", newline="") as f:
                out=list(csv.DictReader(f))[0]
            self.assertIn("수정 제안(로컬 LLM)",out[AFTER])
            self.assertEqual(out["보완 완료 여부"],"미완료")

    def test_ollama_api_uses_local_korean_structured_output(self):
        row = {"검사 프로퍼티": "annotations[0].category_name",
               "검사 규칙": "category_id 14 = 유사 백태",
               "보완 전(실제값/미리보기)": '"유사 백화"',
               STATUS: "현재 규칙 위반", AFTER: "수정 지시 초안"}
        class Response(BytesIO):
            def __enter__(self): return self
            def __exit__(self, *_): self.close()
        packet = {"message": {"content": json.dumps({
            "suggestion": "유사 백화를 유사 백태로 수정 필요",
            "reason": "유효값과의 불일치", "human_review_required": True
        }, ensure_ascii=False)}}
        with patch('local_correction_agent.urlopen', return_value=Response(
                json.dumps(packet, ensure_ascii=False).encode('utf-8'))) as mock:
            answer = ask_ollama(row, "qwen2.5:7b-instruct", include_before=True)
            request = mock.call_args.args[0]
            sent = json.loads(request.data.decode('utf-8'))
            self.assertEqual(sent['model'], 'qwen2.5:7b-instruct')
            self.assertEqual(sent['think'], False)
            self.assertIn('유사 백화', sent['messages'][1]['content'])
            self.assertEqual(request.full_url, 'http://127.0.0.1:11434/api/chat')
            self.assertIn('유사 백태', answer['suggestion'])

    def test_skipped_and_does_not_call_model(self):
        headers=["사업ID","규칙그룹","검사 규칙","매칭상태",STATUS,AFTER]
        with tempfile.TemporaryDirectory() as folder:
            src=Path(folder)/"a.csv"
            with src.open("w", encoding="utf-8-sig", newline="") as f:
                writer=csv.DictWriter(f, fieldnames=headers);writer.writeheader()
                writer.writerow({"사업ID":"022","규칙그룹":"7","검사 규칙":"regex",
                                 "매칭상태":"일치",STATUS:"현재 규칙 통과", AFTER:"old"})
            def fail(*a,**k): raise AssertionError("must skip")
            result=enhance_csv(src, infer=fail)
            self.assertEqual(result["generated"],0)
            self.assertEqual(result["skipped"],1)

if __name__ == '__main__': unittest.main()
