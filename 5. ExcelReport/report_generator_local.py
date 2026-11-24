# -*- coding: utf-8 -*-
"""Windows/local Excel writer fallback. The report generator's parser is shared."""
from collections import Counter
from pathlib import Path
from copy import copy
from openpyxl.styles import PatternFill, Font, Alignment


def build_local(template, destination, project_name, groups, when, round_no, row_builder):
    """Use the user's standard Excel template, never the sample rule data."""
    from openpyxl import load_workbook

    wb = load_workbook(str(template))
    types = ['보완 종합 현황', '파일 보완 목록', '구조 보완 목록',
             '형식 보완 목록', '구축기업 구문규칙', '검사 구문규칙']
    base = {t:wb[t] for t in types}
    for idx in range(2, len(groups)+1):
        for t in types:
            ws = wb.copy_worksheet(base[t])
            ws.title = f'{idx}. {t}'
    for t,ws in base.items():
        ws.title = f'1. {t}'

    summary=wb['데이터 그룹 보완 종합 현황']
    summary['A3'] = None
    summary['A4'] = project_name + ' 종합'
    summary['C3'] = f'보완기준일: {when} / {round_no}차 보완(검토 초안)'
    group_totals = {k:[0,0] for k in ('file','struct','format')}
    for g in groups:
        for k,(total,errors) in g['counts'].items():
            group_totals[k][0]+=total
            group_totals[k][1]+=errors
    for row,k in [(6,'file'),(7,'struct'),(8,'format')]:
        summary[f'B{row}']=group_totals[k][0]
        summary[f'D{row}']=group_totals[k][1]
        summary[f'F{row}']=0
    for idx,g in enumerate(groups):
        row=10+idx*6
        if idx>0:
            # Rebuild the second+ section title and field merges.
            # Otherwise the long project title runs across unmerged cells.
            from openpyxl.utils import get_column_letter
            for offset in range(5):
                for col in range(1,7):
                    source=summary.cell(10+offset,col)
                    target=summary.cell(row+offset,col)
                    if target.__class__.__name__=='MergedCell':
                        continue
                    if source.has_style:target._style=copy(source._style)
                    if source.value is not None:target.value=source.value
        if idx > 0:
            summary.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
            for rr in range(row+1,row+5):
                for c1,c2 in [(2,3),(4,5),(6,7)]:
                    summary.merge_cells(start_row=rr,start_column=c1,end_row=rr,end_column=c2)
        summary[f'A{row}']=g['name']
        summary[f'A{row}'].fill=PatternFill('solid',fgColor='333333')
        summary[f'A{row}'].alignment=Alignment(vertical='center',horizontal='center')
        summary.row_dimensions[row].height=30
        summary.row_dimensions[row+1].height=26
        for offset,k in [(2,'file'),(3,'struct'),(4,'format')]:
            r=row+offset
            summary[f'A{r}']={'file':'구조 오류(파일)','struct':'구조 오류(항목)','format':'형식 오류'}[k]
            summary[f'B{r}']=g['counts'][k][0]
            summary[f'D{r}']=g['counts'][k][1]
            summary[f'F{r}']=0

    summary.column_dimensions['A'].width=85
    for c in 'BCDEFG':summary.column_dimensions[c].width=16

    for idx,g in enumerate(groups,start=1):
        ov=wb[f'{idx}. 보완 종합 현황']
        ov['B5']=g['name']
        ov['B6']=f'{when} / {round_no}차 보완(검토 초안)'
        for row,k in [(11,'file'),(12,'struct'),(13,'format')]:
            ov[f'B{row}']=g['counts'][k][0]
            ov[f'D{row}']=g['counts'][k][1]
            ov[f'F{row}']=0
        ov['G11']=(f"규칙 불일치 {len(g['differences'])}셀 · 확인 필요"
                   if g['differences'] else '규칙 비교 일치')
        ov['G12']='검사결과서 기반 · 실제 보완 미확인'
        ov['G13']='실제 원본 재검증 후 완료 집계'
        for key,tab,base_header in [('file','파일 보완 목록',6),('struct','구조 보완 목록',5),('format','형식 보완 목록',5)]:
            ws=wb[f'{idx}. {tab}']
            # Remove ALL embedded demonstration entries and old count rows.
            to_clear_end=max(ws.max_row,8)
            for line in ws.iter_rows(min_row=3,max_row=to_clear_end,min_col=1,max_col=6):
                for c in line:
                    if c.__class__.__name__!='MergedCell':
                        c.value=None
            headers=(['파일경로','파일명','보완 항목','보완 전','보완 후','비고']
                     if key=='file' else ['파일명','오류 유형','프로퍼티','보완 전','보완 후','비고'])
            entries=[row_builder(key,rowidx,cells,g['rules']['inspection']) for rowidx,cells in g['errors'][key]]
            type_counts=Counter(str(cells.get(3) if key=='file' else cells.get(2)) for _,cells in g['errors'][key])
            for j,(name,count) in enumerate(sorted(type_counts.items()),start=3):
                ws.cell(j,1,name);ws.cell(j,2,count);ws.cell(j,3,0)
            header_row=max(base_header,4+len(type_counts))
            for col,heading in enumerate(headers,1):
                ws.cell(header_row,col,heading)
            for ri,row in enumerate(entries,header_row+1):
                for col,value in enumerate(row,1):
                    ws.cell(ri,col,value)
            for col,width in zip('ABCDEF',[29,23,24,56,70,35]):
                ws.column_dimensions[col].width=width
            for row_idx in range(header_row+1,header_row+1+len(entries)):
                ws.row_dimensions[row_idx].height=32
                for col_idx in (4,5,6):
                    ws.cell(row_idx,col_idx).alignment=Alignment(vertical='center',wrap_text=True)
            ws.freeze_panes=f'A{header_row+1}'

        for which,tab in [('company','구축기업 구문규칙'),('inspection','검사 구문규칙')]:
            ws=wb[f'{idx}. {tab}']
            # 템플릿에 남아 있는 다른 사업의 병합 범위는 이번 사업의 셀 위치와
            # 충돌할 수 있으므로 규칙 시트에서만 해제합니다. (원본 템플릿은 보존)
            for merged_range in list(ws.merged_cells.ranges):
                ws.unmerge_cells(str(merged_range))
            # Throw away the unrelated sample sheet FORMATS and values too.
            rule_matrix=g['rules'][which]
            height=max(ws.max_row,len(rule_matrix))
            width=max(ws.max_column,max(map(len,rule_matrix),default=0))
            for line in ws.iter_rows(min_row=1,max_row=height,min_col=1,max_col=width):
                for cell in line:
                    cell.value=None
                    cell._style=None
                    cell.comment=None
            for rowidx,row in enumerate(rule_matrix,start=1):
                for col,val in enumerate(row,start=1):
                    if val is not None:ws.cell(rowidx,col).value=val
            from openpyxl.utils import get_column_letter as letter
            type_col=next((i for i,v in enumerate(rule_matrix[1],1) if v=='타입'),5)
            ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=type_col-1)
            ws.merge_cells(start_row=1,start_column=type_col,end_row=1,end_column=len(rule_matrix[1]))
            for row in ws.iter_rows(min_row=1,max_row=len(rule_matrix),min_col=1,max_col=len(rule_matrix[1])):
                for c in row:
                    if c.__class__.__name__=='MergedCell':continue
                    c.font=Font(name='맑은 고딕',size=10,bold=c.row in (1,2),color='000000')
                    c.alignment=Alignment(vertical='center',wrap_text=c.column>=len(rule_matrix[1])-1)
                    if c.row<=2:
                        c.fill=PatternFill('solid',fgColor='333333' if c.row==1 else 'C0C0C0')
            ws.row_dimensions[1].height=31
            ws.row_dimensions[2].height=27
            for col in range(1,len(rule_matrix[1])+1):
                ws.column_dimensions[letter(col)].width=(46 if col>=len(rule_matrix[1])-1 else 17)
            ws.freeze_panes='A3'

    # Remove all template notes/comments and normalize ALL text to black.
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.__class__.__name__=='MergedCell':continue
                if c.comment is not None:c.comment=None
                if c.value is not None or c.has_style:
                    current=copy(c.font)
                    current.color='000000'
                    c.font=current
    # Dark title bars must remain dark, with white text for legibility.
    # This runs AFTER global black font normalization, so no title is reverted.
    def apply_dark(ws, rowno, startcol, endcol):
        for ci in range(startcol, endcol + 1):
            cell = ws.cell(rowno, ci)
            cell.fill = PatternFill('solid', fgColor='333333')
            if cell.__class__.__name__ != 'MergedCell':
                f = copy(cell.font)
                f.color = 'FFFFFF'
                f.bold = True
                cell.font = f

    def apply_gray(ws, rowno, startcol, endcol):
        for ci in range(startcol, endcol + 1):
            cell = ws.cell(rowno, ci)
            cell.fill = PatternFill('solid', fgColor='C0C0C0')
            if cell.__class__.__name__ != 'MergedCell':
                f = copy(cell.font)
                f.color = '000000'
                f.bold = True
                cell.font = f

    for rr in [4] + [10 + 6*i for i in range(len(groups))]:
        apply_dark(summary, rr, 1, 7)
        apply_gray(summary, rr + 1, 1, 7)
    apply_gray(summary, 5, 1, 7)

    for idx in range(1, len(groups)+1):
        overview = wb[f'{idx}. 보완 종합 현황']
        for rr in [2, 4, 9]:
            apply_dark(overview, rr, 1, 7)
        apply_gray(overview, 10, 1, 7)
        for tab in ('파일 보완 목록', '구조 보완 목록', '형식 보완 목록'):
            sheet = wb[f'{idx}. {tab}']
            apply_dark(sheet, 1, 1, 3)
        for tab in ('구축기업 구문규칙', '검사 구문규칙'):
            sheet = wb[f'{idx}. {tab}']
            width = max(1, len(groups[idx-1]['rules']['company'][1]))
            apply_dark(sheet, 1, 1, width)
            apply_gray(sheet, 2, 1, width)

    destination=Path(destination)
    destination.parent.mkdir(parents=True,exist_ok=True)
    wb.save(str(destination))
