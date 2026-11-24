"""DataGuard step4 offline tests. Run: python test_data_loader_matcher.py"""
from __future__ import annotations
import csv
import json
import tempfile
import unittest
from pathlib import Path
import zipfile
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from data_loader_matcher import normalize_filename, parse_inspection, match_inspection, run, scan_zip


def mock_inspection(path):
    ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    def make_sheet(rows):
        lines = [f'<worksheet xmlns="{ns}"><sheetData>']
        for r, entries in enumerate(rows, 1):
            lines.append(f'<row r="{r}">')
            for col, value in entries.items():
                from xml.sax.saxutils import escape
                lines.append(f'<c r="{col}{r}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>')
            lines.append('</row>')
        lines.append('</sheetData></worksheet>')
        return ''.join(lines).encode('utf-8')
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('xl/workbook.xml', '''<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="1. 구조 오류 목록" sheetId="1" r:id="rId1"/><sheet name="2. 형식 오류 목록" sheetId="2" r:id="rId2"/></sheets></workbook>''')
        z.writestr('xl/_rels/workbook.xml.rels', '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>''')
        z.writestr('xl/worksheets/sheet1.xml', make_sheet([
            {'A':'검사 항목별 건수'}, {'A':'검사 항목', 'B':'오류 수'},
            {'A':'정의되지 않은 Key', 'B':'2'}, {},
            {'A':'파일명', 'B':'검사 항목', 'C':'오류 메시지'},
            {'A':'07075.json', 'B':'정의되지 않은 Key', 'C':'text'},
            {'A':'01234.json', 'B':'정의되지 않은 Key', 'C':'text'},
        ]))
        z.writestr('xl/worksheets/sheet2.xml', make_sheet([
            {'A':'검사 항목별 건수'}, {'A':'검사 항목', 'B':'오류 수'},
            {'A':'value 유효범위 오류', 'B':'2'}, {},
            {'A':'파일명', 'B':'검사 항목', 'C':'검사 규칙', 'D':'검사 프로퍼티', 'E':'오류 메시지'},
            {'A':'08655.txt(2)', 'B':'value 유효범위 오류', 'C':'최댓값: 1', 'D':'h', 'E':'1.0012'},
            {'A':'00743.txt(3)', 'B':'value 유효범위 오류', 'C':'최솟값: 0(초과)', 'D':'h', 'E':'0'},
        ]))


class MatcherTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.inspection = self.root / '(019)검사결과서_샘플.xlsx'
        mock_inspection(self.inspection)
        self.files = self.root / 'LabelingData'
        (self.files / 'json').mkdir(parents=True)
        (self.files / 'txt').mkdir()
        (self.files / 'json' / '07075.json').write_text('{"text":"bad"}', encoding='utf-8')
        (self.files / 'txt' / '08655.txt').write_text('h:1.0012', encoding='utf-8')
        (self.files / 'txt' / '00743.txt').write_text('h:0', encoding='utf-8')

    def tearDown(self):
        self.temp.cleanup()

    def test_normalize_suffix(self):
        self.assertEqual(normalize_filename('08655.txt(2)'), ('08655.txt', '2'))
        self.assertEqual(normalize_filename('sub\\07075.json'), ('07075.json', ''))

    def test_inspection_sheet_rows(self):
        rows, meta = parse_inspection(self.inspection)
        self.assertEqual(meta['issues'], 4)
        self.assertEqual([r['규칙그룹'] for r in rows], ['1', '1', '2', '2'])
        self.assertEqual(rows[-1]['행표기'], '3')
        self.assertEqual(rows[-1]['검사결과서 행'], 7)

    def test_folder_matching_read_only(self):
        before = {p.relative_to(self.files).as_posix(): p.read_bytes() for p in self.files.rglob('*') if p.is_file()}
        rows, meta = match_inspection(self.inspection, folder=self.files)
        self.assertEqual(meta['status_counts']['일치'], 3)
        self.assertEqual(meta['status_counts']['파일 없음'], 1)
        self.assertEqual(meta['unique_matched_files'], 3)
        after = {p.relative_to(self.files).as_posix(): p.read_bytes() for p in self.files.rglob('*') if p.is_file()}
        self.assertEqual(before, after)

    def test_zip_matching_without_extracting(self):
        zpath = self.root / 'labels.zip'
        with zipfile.ZipFile(zpath, 'w') as z:
            for p in self.files.rglob('*'):
                if p.is_file():
                    z.write(p, p.relative_to(self.files).as_posix())
        rows, meta = match_inspection(self.inspection, archive=zpath)
        self.assertEqual(meta['status_counts']['일치'], 3)
        self.assertEqual(meta['status_counts']['파일 없음'], 1)
        self.assertFalse((self.root / 'json').exists())

    def test_duplicate_not_arbitrarily_matched(self):
        (self.files / 'another').mkdir()
        (self.files / 'another' / '07075.json').write_text('{}')
        rows, meta = match_inspection(self.inspection, folder=self.files)
        self.assertEqual(meta['status_counts']['동일 파일명 중복'], 1)
        self.assertEqual(rows[0]['실제 데이터 위치'], '')

    def test_zip_path_traversal_rejected(self):
        zpath = self.root / 'unsafe.zip'
        with zipfile.ZipFile(zpath, 'w') as z:
            z.writestr('../07075.json', '{}')
        with self.assertRaises(ValueError):
            scan_zip(zpath)

    def test_exports_utf8_csv_json(self):
        result = run(self.inspection, self.files, mode='folder', output_dir=self.root / 'results')
        with open(result['csv'], encoding='utf-8-sig', newline='') as f:
            data = list(csv.DictReader(f))
        with open(result['json'], encoding='utf-8') as f:
            meta = json.load(f)
        self.assertEqual(len(data), 4)
        self.assertEqual(meta['issues'], 4)
        self.assertEqual(data[-1]['매칭파일명'], '00743.txt')


if __name__ == '__main__':
    unittest.main(verbosity=2)
