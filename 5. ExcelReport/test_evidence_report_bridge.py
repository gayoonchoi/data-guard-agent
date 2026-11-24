# -*- coding: utf-8 -*-
import tempfile
import unittest
import csv
from pathlib import Path
from evidence_report_bridge import _patch_sheet_xml, _evidence_rows

class EvidenceBridgeTests(unittest.TestCase):
    def test_only_value_changes_style_intact(self):
        src=(b'<?xml version="1.0" encoding="utf-8"?><worksheet><sheetData>'
             b'<row r="6"><c r="A6" s="2" t="inlineStr"><is><t>file.json</t></is></c>'
             b'<c r="D6" s="18" t="inlineStr"><is><t>old</t></is></c>'
             b'<c r="E6" s="20" t="inlineStr"><is><t>old instruction</t></is></c>'
             b'</row></sheetData><mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells></worksheet>')
        out=_patch_sheet_xml(src,{6:{'D':'현재 값: <text>&', 'E':'수정 지시'}}).decode()
        self.assertIn('r="D6" s="18" t="inlineStr"',out)
        self.assertIn('r="E6" s="20" t="inlineStr"',out)
        self.assertIn('현재 값: &lt;text&gt;&amp;',out)
        self.assertIn('ref="A1:C1"',out)
        self.assertIn('file.json',out)

    def test_duplicate_keys_are_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d)/'e.csv'
            fields=['사업ID','규칙그룹','검사결과서 시트','검사결과서 행','파일명(원문)',
                    '매칭상태','보완 전(실제값/미리보기)','보완 후(규칙상 요구사항/초안)', '현재 검증 상태']
            rec=dict(zip(fields,['019','1','1. 구조 오류 목록','6','07075.json',
                                 '일치','"foo"','top-level text 제거','현재 규칙 위반']))
            with p.open('w',encoding='utf-8-sig',newline='') as f:
                w=csv.DictWriter(f,fieldnames=fields);w.writeheader();w.writerow(rec);w.writerow(rec)
            with self.assertRaisesRegex(ValueError,'중복'):
                _evidence_rows(p,'019')

if __name__=='__main__':unittest.main()
