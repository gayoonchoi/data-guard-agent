# -*- coding: utf-8 -*-
"""DataGuard - TTA report draft with the project's own syntax rule sheets.

읽기 전용 입력: 검사결과서 Excel. 수정 완료로 주장하지 않음.
산출: 원본 템플릿 레이아웃 기반 N차 보완결과서 검토초안.

Python dependencies: artifact_tool (xlsx writer); standard-library (xlsx reader).
"""
from __future__ import annotations

import argparse
from collections import Counter
from datetime import datetime
from pathlib import Path
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

try:
    from artifact_tool import Blob, SpreadsheetFile
except ImportError:
    Blob = SpreadsheetFile = None  # Windows fallback is project-specific

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "보완결과서_표준양식.xlsx"
MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _col_number(addr: str) -> int:
    n = 0
    for ch in re.match(r"[A-Z]+", addr).group():
        n = n * 26 + ord(ch) - 64
    return n


def _excel_col(num: int) -> str:
    chars = ""
    while num:
        num, x = divmod(num - 1, 26)
        chars = chr(65 + x) + chars
    return chars


class WorkbookReader:
    """Streaming, read-only parser: no source data modification."""
    def __init__(self, path):
        self.path = Path(path)
        self.zip = zipfile.ZipFile(self.path)
        wb = ET.fromstring(self.zip.read("xl/workbook.xml"))
        rels = ET.fromstring(self.zip.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in rels}
        self.sheets = {}
        for s in wb.find(MAIN_NS + "sheets"):
            rel_path = targets[s.get(REL_NS + "id")]
            member = (rel_path.lstrip("/") if rel_path.startswith("/")
                      else posixpath.normpath("xl/" + rel_path))
            self.sheets[s.get("name")] = member
        self.shared = []
        if "xl/sharedStrings.xml" in self.zip.namelist():
            strings = ET.fromstring(self.zip.read("xl/sharedStrings.xml"))
            self.shared = ["".join(s.itertext()) for s in strings.findall(MAIN_NS + "si")]

    def rows(self, name: str):
        if name not in self.sheets:
            return
        member = self.sheets[name]
        with self.zip.open(member) as source:
            for _, row in ET.iterparse(source, events=("end",)):
                if row.tag != MAIN_NS + "row":
                    continue
                values = {}
                for c in row.findall(MAIN_NS + "c"):
                    ref = c.get("r", "")
                    if not re.match(r"^[A-Z]+\d+$", ref):
                        continue
                    element = c.find(MAIN_NS + "v")
                    inline = c.find(MAIN_NS + "is")
                    if element is not None and element.text is not None:
                        v = element.text
                        if c.get("t") == "s":
                            v = self.shared[int(v)]
                        elif c.get("t") not in ("str", "inlineStr", "e"):
                            try:
                                v = float(v) if "." in v or "E" in v else int(v)
                            except (ValueError, TypeError):
                                pass
                    elif inline is not None:
                        v = "".join(inline.itertext())
                    else:
                        continue
                    values[_col_number(ref)] = v
                yield int(row.get("r", "0")), values
                row.clear()

    def cell(self, sheet, row_num, col_num):
        for idx, vals in self.rows(sheet):
            if idx == row_num:
                return vals.get(col_num)
            if idx > row_num:
                break
        return None

    def close(self):
        self.zip.close()


def project_id_from_inspection(path):
    """Use the same identity logic as RuleKB registration (including (8) -> 007)."""
    import importlib.util
    f=ROOT.parent / 'DataGuard_RuleKB_v01' / 'extract_rules.py'
    if f.is_file():
        spec=importlib.util.spec_from_file_location('dataguard_rule_extractor_for_report',f)
        module=importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        try:
            return module.identify_project_id(Path(path), {'projects':{}})
        except ValueError as error:
            if '식별하지 못했습니다' not in str(error):
                raise
    m=re.match(r"^\((\d{1,4})\)",Path(path).name)
    if not m:
        raise ValueError("사업 ID를 식별할 수 없습니다. 파일명 앞에 (019)처럼 번호를 붙여 주세요.")
    return m.group(1).zfill(3)


def extract_all(reader: WorkbookReader):
    groups = []
    for name in reader.sheets:
        m = re.fullmatch(r"(\d+)\. 검사 종합 현황", name)
        if not m:
            continue
        g = int(m.group(1))
        company_name, inspection_name = f"{g}. 구축기업 구문규칙", f"{g}. 검사 구문규칙"
        if company_name not in reader.sheets or inspection_name not in reader.sheets:
            raise ValueError(f"{g}번 그룹의 구축기업/검사 구문규칙 시트가 누락됐습니다.")
        overview = {r: vals for r, vals in reader.rows(name) if r in (5, 11, 12, 13)}
        group_label = str(overview.get(5, {}).get(2) or f"그룹 {g}")
        counts = {
            'file': (int(overview.get(11, {}).get(6) or 0), int(overview.get(11, {}).get(4) or 0)),
            'struct': (int(overview.get(12, {}).get(6) or 0), int(overview.get(12, {}).get(4) or 0)),
            'format': (int(overview.get(13, {}).get(6) or 0), int(overview.get(13, {}).get(4) or 0)),
        }
        errors = {'file': [], 'struct': [], 'format': []}
        for kind, src, first in [
            ('file', f'{g}. 파일 오류 목록', 2),
            ('struct', f'{g}. 구조 오류 목록', 6),
            ('format', f'{g}. 형식 오류 목록', 6),
        ]:
            if src not in reader.sheets:
                raise ValueError(f"필수 시트 누락: {src}")
            for rowidx, cells in reader.rows(src):
                if rowidx < first:
                    continue
                a, b = cells.get(1), cells.get(2)
                if a in (None, '') or b in (None, ''):
                    continue
                errors[kind].append((rowidx, cells))
        # All original rule cells, including nesting and conflicting patterns.
        rules = {}
        for which, sheet_name in [('company',company_name),('inspection',inspection_name)]:
            rows = [(r, vals) for r, vals in reader.rows(sheet_name) if r <= 1000]
            if not rows:
                raise ValueError(f"구문규칙 시트가 비었습니다: {sheet_name}")
            height = max(r for r,_ in rows)
            width = max(max(v) for _,v in rows if v)
            if height > 999 or width > 50:
                raise ValueError('예상보다 큰 규칙 표입니다. 시트 형식을 확인해 주세요.')
            matrix = [[None]*width for _ in range(height)]
            for r, vals in rows:
                for col, value in vals.items():
                    matrix[r-1][col-1] = value
            rules[which] = matrix
        gaps=[]
        a,b = rules['company'],rules['inspection']
        for r in range(max(len(a),len(b))):
            for col in range(max(len(a[r]) if r<len(a) else 0,len(b[r]) if r<len(b) else 0)):
                va=a[r][col] if r<len(a) and col<len(a[r]) else None
                vb=b[r][col] if r<len(b) and col<len(b[r]) else None
                if (str(va) if va is not None else '') != (str(vb) if vb is not None else ''):
                    gaps.append(f"{_excel_col(col+1)}{r+1}")
        for kind, (_, stated) in counts.items():
            if stated != len(errors[kind]):
                raise ValueError(f"{g}번 그룹 {kind} 집계 불일치: 종합현황 {stated}, 상세 목록 {len(errors[kind])}. 검토가 필요합니다.")
        groups.append({'id':g,'name':group_label,'counts':counts,'errors':errors,'rules':rules,'differences':gaps})
    groups.sort(key=lambda g:g['id'])
    if not groups:
        raise ValueError('검사결과서에서 「1. 검사 종합 현황」 형태의 시트를 찾지 못했습니다.')
    return groups


def _fmt_value(value):
    if value is None:
        return '(미기재)'
    return str(value)


def _suggest(kind, fields, rule_matrix):
    """Never infer corrected literal values when original labeling data is absent."""
    if kind == 'file':
        return ('검사결과서: ' + _fmt_value(fields.get(4)),
                '파일 복구 또는 경로 수정 후 재검사 필요 (원본 파일 미확인)', '미보완 · 원본 확인 필요')
    if kind == 'struct':
        prop = str(fields.get(3) or '')
        if prop.strip() == 'text':
            return (f"검사결과서: 정의되지 않은 Key / {prop}",
                    'JSON 최상위 $.text가 실제로 존재하는 경우 해당 키만 제거. '
                    'shapes[*].text는 규칙에 존재하므로 유지 (원본 확인 필요)',
                    '미보완 · 실제 위치 확인 필요')
        return (f"검사결과서: {_fmt_value(fields.get(2))} / {prop}",
                '해당 사업의 구문규칙에서 허용하지 않는 항목인지 확인한 뒤 수정',
                '미보완 · 규칙/원본 확인 필요')
    rule = _fmt_value(fields.get(3))
    prop = _fmt_value(fields.get(4))
    observed = _fmt_value(fields.get(5))
    before = f'검사 당시 {prop} = {observed} / 위반 규칙: {rule}'
    if prop == 'shapes.points':
        after = '좌표값과 이미지 크기, 축별 허용범위를 확인한 뒤 실제 값을 보완. 규칙 충돌 시 임의 숫자 대입 금지'
        state = '미보완 · 좌표 기준 검토'
    elif prop == 'h' and ('최솟값' in rule or '최댓값' in rule):
        after = 'TXT의 h는 0 < h ≤ 1. 원천 이미지에서 높이를 확인한 뒤 수정 (임의값 대입 금지)'
        state = '미보완 · 원본 영상 확인'
    else:
        after = f'{prop} 항목을 해당 구문 규칙({rule})에 맞춰 수정 후 재검증 (원본 값 확인 필요)'
        state = '미보완 · 원본 확인 필요'
    return (before,after,state)


def _set(sheet, address, val):
    sheet.get_range(address).values = [[val]]


def _zero_template_rows(sheet, endrow, endcol='F'):
    # Retain header styles and merged cells, but erase all sample values.
    # Do not write to merged placeholders (only clear detailed data rows).
    sheet.get_range(f'A8:{endcol}{endrow}').values = [[None]*(ord(endcol)-64) for _ in range(endrow-7)]


def _write_table(sheet, header_row, rows, issue_counts, kind):
    # The original template uses row 6 for file issues and row 5 for other issues.
    sheet.get_range('A3:C4').values = [[None]*3,[None]*3]
    by_kind=Counter((row[2] if kind == 'file' else row[1]) for row in rows)
    count_line=3
    for label,count in sorted(by_kind.items()):
        sheet.get_range(f'A{count_line}:C{count_line}').values = [[label,count,0]]
        count_line+=1
    if count_line>=header_row:
        # Keep at least one free row between counts and table header.
        header_row = count_line+1
    if kind == 'file':
        header=['파일경로','파일명','보완 항목','보완 전','보완 후','비고']
    else:
        header=['파일명','오류 유형','프로퍼티','보완 전','보완 후','비고']
    sheet.get_range(f'A{header_row}:F{header_row}').values=[header]
    if rows:
        for start in range(0,len(rows),700):
            batch = rows[start:start+700]
            first = header_row+start+1
            last = first+len(batch)-1
            sheet.get_range(f'A{first}:F{last}').values=batch
    sheet.get_range(f'A{header_row}:F{header_row}').format={
        'fill':'#DCE8F0','font':{'bold':True},'row_height':23,
        'vertical_alignment':'center'
    }
    # Set practical column widths, preserving 019 template general appearance.
    for letter,width in zip('ABCDEF',[29,23,24,56,70,35]):
        sheet.get_range(f'{letter}:{letter}').format.column_width=width
    sheet.freeze_panes.freeze_rows(header_row)


def _copy_template_sheets(wb, groups):
    base = {name:wb.worksheets.get_item(name) for name in [
        '보완 종합 현황','파일 보완 목록','구조 보완 목록',
        '형식 보완 목록','구축기업 구문규칙','검사 구문규칙'
    ]}
    ranges={
        '보완 종합 현황':'A1:G13', '파일 보완 목록':'A1:F7',
        '구조 보완 목록':'A1:F6','형식 보완 목록':'A1:F7',
        '구축기업 구문규칙':'A1:O137','검사 구문규칙':'A1:O137'
    }
    # Clone from untouched template before renaming original group-1 sheets.
    for i,g in enumerate(groups[1:],start=2):
        for name,s in base.items():
            new=wb.worksheets.add(f'{i}. {name}')
            new.get_range(ranges[name]).copy_from(s.get_range(ranges[name]),'all')
    for name,s in base.items():
        s.name=f'1. {name}'


def _sheet(wb, group_index, kind):
    return wb.worksheets.get_item(f'{group_index}. {kind}')


def _group_overview(summary, group, idx, when, round_no):
    counts=group['counts']
    summary.get_range('B1:F1').values=[[None]*5]
    _set(summary, 'A5', '데이터명')
    _set(summary, 'B5', group['name'])
    _set(summary, 'B6', f'{when} / {round_no}차 보완(검토 초안)')
    for row,kind in [(11,'file'),(12,'struct'),(13,'format')]:
        total,errors=counts[kind]
        _set(summary,f'B{row}',total)
        _set(summary,f'D{row}',errors)
        _set(summary,f'F{row}',0)
    diff=group['differences']
    _set(summary,'G11', f'구문 규칙 불일치 {len(diff)}셀 — 확인 필요' if diff else '구문 규칙 비교 일치')
    _set(summary,'G12','검사결과서 기준 · 실제 보완 완료 미확인')
    _set(summary,'G13','수정 전후 원본 검증 후 완료 수량 반영')


def _write_rules(sheet, matrix):
    # Remove unrelated sample template rules (all 137 original example rows).
    h=max(137,len(matrix))
    w=max(15,max(map(len,matrix)))
    for start in range(0,h,120):
        n=min(120,h-start)
        sheet.get_range(f'A{start+1}:{_excel_col(w)}{start+n}').values=[[None]*w for _ in range(n)]
    for start in range(0,len(matrix),200):
        batch=matrix[start:start+200]
        sheet.get_range(f'A{start+1}:{_excel_col(len(matrix[0]))}{start+len(batch)}').values=batch
    sheet.get_range(f'A1:{_excel_col(w)}2').format={'fill':'#DCE8F0','font':{'bold':True}}
    for col in range(1,w+1):
        letter=_excel_col(col)
        width=20 if col < 6 else (32 if col in (14,15) else 16)
        sheet.get_range(f'{letter}:{letter}').format.column_width=width
    sheet.freeze_panes.freeze_rows(2)


def _populate_group_list(sheet, project_name, groups, when, round_no):
    _set(sheet,'A3',None)
    _set(sheet,'A4',project_name+' 종합')
    _set(sheet,'C3',f'보완기준일: {when} / {round_no}차 보완(검토 초안)')
    totals={k:[0,0] for k in ('file','struct','format')}
    for g in groups:
        for kind,(t,e) in g['counts'].items():
            totals[kind][0]+=t;totals[kind][1]+=e
    # Set formulas explicitly: imported template formulas may still reference only 1 group.
    for row,offset in [(6,2),(7,3),(8,4)]:
        for col in ('B','D','F'):
            refs=','.join(f'{col}{10+i*6+offset}' for i in range(len(groups)))
            sheet.get_range(f'{col}{row}').formulas=[[f'=SUM({refs})']]
    for idx,g in enumerate(groups):
        row=10+idx*6
        if idx>0:
            # Reuse the original summary section's headings and styling.
            sheet.get_range(f'A{row}:F{row+4}').copy_from(sheet.get_range('A10:F14'),'all')
        _set(sheet,f'A{row}',g['name'])
        for offset,kind in [(2,'file'),(3,'struct'),(4,'format')]:
            r=row+offset
            total,errors=g['counts'][kind]
            _set(sheet,f'A{r}',{'file':'구조 오류(파일)','struct':'구조 오류(항목)','format':'형식 오류'}[kind])
            _set(sheet,f'B{r}',total)
            _set(sheet,f'D{r}',errors)
            _set(sheet,f'F{r}',0)
        sheet.get_range(f'A{row}:F{row}').format={'fill':'#DCE8F0','font':{'bold':True},'row_height':24}
    for letter,width in zip('ABCDEF',[70,22,17,20,17,20]):
        sheet.get_range(f'{letter}:{letter}').format.column_width=width


def _issue_row(kind, source, cells, rule_matrix):
    filename = str(cells.get(2) if kind=='file' else cells.get(1) or '')
    if kind=='file':
        before,after,note = _suggest(kind,cells,rule_matrix)
        return [str(cells.get(1) or ''),filename,str(cells.get(3) or ''),before,after,note]
    issue = str(cells.get(2) or '')
    prop = str(cells.get(3) or '') if kind=='struct' else str(cells.get(4) or '')
    before,after,note=_suggest(kind,cells,rule_matrix)
    if kind=='format' and re.fullmatch(r'.+\.txt\(\d+\)',filename):
        basename,linenumber=filename.rsplit('(',1)
        filename=f'{basename[:-4]}_{linenumber[:-1]}.txt'
        note+=f' / 원검사 표기: {cells.get(1)}'
    return [filename,issue,prop,before,after,note]



def _style_like_uploaded_sample(report_path, template_path, inspection_path):
    """First apply the report template, then EXACTLY copy original rule worksheets.

    The Excel worksheet copy MUST run last: later style engines could overwrite
    the original style again. Errors are raised, never silently ignored.
    """
    import importlib.util
    import inspect

    style_path = ROOT / "template_style_engine.py"
    spec = importlib.util.spec_from_file_location("dataguard_template_style", style_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    apply_fn = module.apply_template_style
    parameters = inspect.signature(apply_fn).parameters
    if "inspection_path" in parameters:
        styled = apply_fn(report_path, template_path, inspection_path=inspection_path)
    elif len(parameters) >= 3:
        styled = apply_fn(report_path, template_path, inspection_path)
    else:
        styled = apply_fn(report_path, template_path)

    # No other workbook formatting is allowed AFTER this call.
    copier_path = ROOT / "rule_sheet_copy_exact.py"
    copier_spec = importlib.util.spec_from_file_location("dataguard_rule_copy_exact", copier_path)
    copier = importlib.util.module_from_spec(copier_spec)
    copier_spec.loader.exec_module(copier)
    return copier.restore_original_rule_sheets(inspection_path, styled)


def generate_report(inspection_file, output_path=None, template_path=None,
                    round_no=1, project_id=None):
    inputfile=Path(inspection_file)
    if not inputfile.is_file():raise FileNotFoundError(inputfile)
    if int(round_no)<1 or int(round_no)>99:raise ValueError('보완 단계는 1~99차만 허용합니다.')
    known_id=project_id_from_inspection(inputfile)
    if project_id and str(project_id).zfill(3)!=known_id:
        raise ValueError(f'사업 ID 불일치: 파일명 {known_id}, 지정값 {project_id}')
    template=Path(template_path) if template_path else TEMPLATE
    if not template.is_file():raise FileNotFoundError(f'보완결과서 표준양식이 없습니다: {template}')
    reader=WorkbookReader(inputfile)
    try:
        groups=extract_all(reader)
        project_name=re.sub(r'\s*\([^()]*\)\s*$', '',groups[0]['name']).strip()
        when=datetime.now().strftime('%Y-%m-%d')
        stamp=datetime.now().strftime('%y%m%d')
        safe_name=re.sub(r'[\\/:*?"<>|\[\]]','_',project_name).strip()[:60]
        dest=Path(output_path) if output_path else ROOT / 'outputs' / f'({known_id})구문정확성_보완결과서_{safe_name}_{int(round_no)}차_{stamp}.xlsx'
        dest.parent.mkdir(parents=True,exist_ok=True)
        # 이미 같은 이름의 파일이 있으면 덮어쓰지 않습니다.
        # 출력 파일명은 유지하고 실행 폴더만 새로 만들어 Windows 파일 잠금을 피합니다.
        if dest.exists():
            import uuid
            subdir = '재생성_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + uuid.uuid4().hex[:6]
            dest = dest.parent / subdir / dest.name
            dest.parent.mkdir(parents=True, exist_ok=False)
        if SpreadsheetFile is None:
            import importlib.util
            local_path = ROOT / 'report_generator_local.py'
            spec = importlib.util.spec_from_file_location('dg_report_local', local_path)
            local_module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(local_module)
            build_local = local_module.build_local
            build_local(template, dest, project_name, groups, when, round_no, _issue_row)
            dest = _style_like_uploaded_sample(dest, template, inputfile)
            total=sum(sum(len(v) for v in g['errors'].values()) for g in groups)
            conflicts=sum(len(g['differences']) for g in groups)
            return {'path':str(dest),'project_id':known_id,'project_name':project_name,
                    'groups':len(groups),'issues':total,'done':0,'mismatch_cells':conflicts,
                    'group_details':[(g['id'],g['name'],sum(len(v) for v in g['errors'].values()),g['differences']) for g in groups]}
        # v0.7: construct output in a NEW workbook, copying only needed template
        # ranges. This prevents stale notes, example styles and wrong merged cells.
        import importlib.util
        builder_path = ROOT / 'report_builder_v07.py'
        spec = importlib.util.spec_from_file_location('dg_report_builder_v07', builder_path)
        builder = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(builder)
        builder.create_report(template, dest, project_name, groups, when, round_no, _issue_row)
        dest = _style_like_uploaded_sample(dest, template, inputfile)
        total=sum(sum(len(v) for v in g['errors'].values()) for g in groups)
        conflicts=sum(len(g['differences']) for g in groups)
        return {'path':str(dest),'project_id':known_id,'project_name':project_name,
                'groups':len(groups),'issues':total,'done':0,'mismatch_cells':conflicts,
                'group_details':[(g['id'],g['name'],sum(len(v) for v in g['errors'].values()),g['differences']) for g in groups]}
    finally:
        reader.close()


def main():
    parser=argparse.ArgumentParser(description='TTA 검사결과서 → 표준양식 보완결과서 검토 초안')
    parser.add_argument('--inspection',required=True)
    parser.add_argument('--template',default=str(TEMPLATE))
    parser.add_argument('--output',default=None)
    parser.add_argument('--round',type=int,default=1)
    args=parser.parse_args()
    result=generate_report(args.inspection,args.output,args.template,args.round)
    print(result)

if __name__=='__main__':main()
