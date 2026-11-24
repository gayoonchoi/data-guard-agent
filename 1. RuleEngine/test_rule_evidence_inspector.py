"""Run: python test_rule_evidence_inspector.py"""
import json
import tempfile
import unittest
from pathlib import Path
import zipfile

from data_loader_matcher import parse_file_reference, normalize_filename, scan_folder, match_inspection
from rule_evidence_inspector import read_value, determine_issue, analyze


def make_inspection(path):
    import xml.sax.saxutils as xs
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rows = [
      {'A': '오류 위치는 파일명(행번호)로 작성하였습니다.'},
      {'A':'파일명','B':'검사 항목','C':'검사 프로퍼티','D':'검사 규칙','E':'오류 메시지'},
      {'A':'moel_counsel_mrc_v.01/dialog.json (3)','B':'pattern 오류',
       'C':'dialogue[1].content','D':'유효패턴: \\S','E':'""'},
      {'A':'administrative_interpretation_dialog/제4장(2).jsonl (2)','B':'pattern 오류',
       'C':'dialogue[0].content','D':'유효패턴: \\S','E':'""'},
    ]
    sheet = [f'<worksheet xmlns="{ns}"><sheetData>']
    for n,r in enumerate(rows,1):
        sheet.append(f'<row r="{n}">')
        for col,val in r.items():
            sheet.append(f'<c r="{col}{n}" t="inlineStr"><is><t>{xs.escape(val)}</t></is></c>')
        sheet.append('</row>')
    sheet.append('</sheetData></worksheet>')
    with zipfile.ZipFile(path,'w') as z:
        z.writestr('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="7. 형식 오류 목록" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr('xl/worksheets/sheet1.xml',''.join(sheet))

class EvidenceTests(unittest.TestCase):
    def test_filename_forms(self):
        self.assertEqual(parse_file_reference('moel_counsel_mrc_v.01/dialog.json (128)'),('moel_counsel_mrc_v.01/dialog.json','128'))
        self.assertEqual(parse_file_reference('administrative_interpretation_dialog/제4장(2).jsonl (31)'),('administrative_interpretation_dialog/제4장(2).jsonl','31'))
        self.assertEqual(normalize_filename('moel_counsel_mrc_v.01/dialog.json (128)'),('dialog.json','128'))
        self.assertEqual(normalize_filename('08655.txt(2)'),('08655.txt','2'))

    def test_read_json_record_and_jsonl(self):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td)
            (root/'d').mkdir()
            (root/'d'/'dialog.json').write_text(json.dumps([{'dialogue':[{'content':'ok'}, {'content':'x'}]}, {'dialogue':[{'content':'ok'}, {'content':'y'}]}, {'dialogue':[{'content':'ok'}, {'content':''}]}]), encoding='utf8')
            (root/'d'/'sample.jsonl').write_text('\n'.join(json.dumps({'dialogue':[{'content':text}]}) for text in ('a',''))+'\n',encoding='utf8')
            value, desc=read_value('folder','d/dialog.json',root,'3','dialogue[1].content')
            self.assertEqual(value, '')
            self.assertIn('3번째',desc)
            val,des=read_value('folder','d/sample.jsonl',root,'2','dialogue[0].content')
            self.assertEqual(val, '')
            self.assertIn('2번째 줄',des)
            self.assertEqual(determine_issue('', r'유효패턴: \S')[0],'현재 규칙 위반')

    def test_zip_evidence_read(self):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td)
            zfile=root/'labels.zip'
            with zipfile.ZipFile(zfile,'w') as z:
                z.writestr('moel_counsel_mrc_v.01/dialog.json', json.dumps([{'dialogue':[{'content':'a'},{'content':'x'}]}]*2+[{'dialogue':[{'content':'a'},{'content':''}]}]))
                z.writestr('administrative_interpretation_dialog/제4장(2).jsonl', '\n'.join(json.dumps({'dialogue':[{'content':text}]}) for text in ('a',''))+'\n')
            value,desc=read_value('zip','administrative_interpretation_dialog/제4장(2).jsonl',zfile,'2','dialogue[0].content')
            self.assertEqual(value,'')
            self.assertEqual(read_value('zip','moel_counsel_mrc_v.01/dialog.json',zfile,'3','dialogue[1].content')[0],'')

    def test_different_folder_same_basename_is_not_selected(self):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td)
            inspect=root/'(022)검사결과서.xlsx'; make_inspection(inspect)
            base=root/'label'; (base/'moel_counsel_mrc_v.01').mkdir(parents=True)
            (base/'other').mkdir()
            (base/'other'/'dialog.json').write_text('[]')
            (base/'moel_counsel_mrc_v.01'/'dialog.json').write_text('[]')
            rows,meta=match_inspection(inspect,folder=base)
            self.assertEqual(rows[0]['매칭상태'],'일치')
            self.assertEqual(rows[0]['행표기'],'3')
            self.assertEqual(rows[0]['실제 데이터 위치'],'moel_counsel_mrc_v.01/dialog.json')
            self.assertEqual(rows[1]['매칭상태'],'파일 없음')

    def test_end_to_end_draft_does_not_change_original(self):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td)
            inspect=root/'(022)검사결과서.xlsx'; make_inspection(inspect)
            base=root/'label'; (base/'moel_counsel_mrc_v.01').mkdir(parents=True)
            (base/'administrative_interpretation_dialog').mkdir(parents=True)
            data=[{'dialogue':[{'content':'hello'}, {'content':'hello'}]}]*2 + [{'dialogue':[{'content':'hello'}, {'content':''}]}]
            file1=base/'moel_counsel_mrc_v.01'/'dialog.json'
            file1.write_text(json.dumps(data),encoding='utf8')
            file2=base/'administrative_interpretation_dialog'/'제4장(2).jsonl'
            file2.write_text('\n'.join(json.dumps({'dialogue':[{'content':text}]}) for text in ('ok',''))+'\n',encoding='utf8')
            before=(file1.read_bytes(),file2.read_bytes())
            res=analyze(inspect,base,output_dir=root/'reports')
            self.assertEqual(res['summary']['issues'],2)
            self.assertEqual(res['summary']['evidence_status_counts'],{'현재 규칙 위반':2})
            self.assertEqual(before,(file1.read_bytes(),file2.read_bytes()))
            self.assertTrue(Path(res['csv']).exists())
            self.assertTrue(all(row['보완 완료 여부']=='미완료' for row in res['rows']))

if __name__=='__main__':
    unittest.main(verbosity=2)
