from pathlib import Path
from collections import Counter
from artifact_tool import Workbook, Blob, SpreadsheetFile

# Input matrices & issue row formatting prepared by report_generator.py

def _col(i):
    s=''
    while i:
        i,z=divmod(i-1,26);s=chr(65+z)+s
    return s

BLACK='#000000'
WHITE='#FFFFFF'
FILL='#333333'   # 제목 바: 검정(원래 표준양식의 진회색)
FILL2='#C0C0C0'  # 표 머리글: 원래 표준양식 회색


def copy_range(dst,src,dest_range,src_range):
    dst.get_range(dest_range).copy_from(src.get_range(src_range),'all')


def text_style(sh,rng,*,bold=False,fill=None,height=None):
    r=sh.get_range(rng)
    cfg={'font':{'color':WHITE if fill == FILL else BLACK,'name':'Malgun Gothic','bold':bold}, 'vertical_alignment':'center'}
    if fill:cfg['fill']=fill
    if height:cfg['row_height']=height
    r.format=cfg


def merge_safe(sh,span):
    sh.merge_cells(span)


def make_summary(wb,template,project,groups,when,round_no):
    sh=wb.worksheets.add('데이터 그룹 보완 종합 현황')
    copy_range(sh,template,'A1:G14','A1:G14')
    group_total=len(groups)
    for idx,g in enumerate(groups):
        title=10+6*idx
        if idx:
            copy_range(sh,sh,f'A{title}:G{title+4}','A10:G14')
        # Copy From may bring formulas into merged placeholder cells. Clear BEFORE merging.
        for rr in [title+1,title+2,title+3,title+4]:
            for cc in ['C','E','G']:
                sh.get_range(f'{cc}{rr}').clear({'apply_to':'contents'})
        merge_safe(sh,f'A{title}:G{title}')
        for rr in [title+1,title+2,title+3,title+4]:
            for left,right in [('B','C'),('D','E'),('F','G')]:
                merge_safe(sh,f'{left}{rr}:{right}{rr}')
        sh.get_range(f'A{title}').values=[[str(g['name'])]]
        text_style(sh,f'A{title}:G{title}',bold=True,fill=FILL,height=28)
        text_style(sh,f'A{title+1}:G{title+1}',bold=True,fill=FILL2,height=25)
        for off,kind in [(2,'file'),(3,'struct'),(4,'format')]:
            rr=title+off
            total,errors=g['counts'][kind]
            sh.get_range(f'A{rr}').values=[[{'file':'구조 오류(파일)','struct':'구조 오류(항목)','format':'형식 오류'}[kind]]]
            sh.get_range(f'B{rr}').values=[[total]]
            sh.get_range(f'D{rr}').values=[[errors]]
            sh.get_range(f'F{rr}').values=[[0]]
            text_style(sh,f'A{rr}:G{rr}',height=22)
    merge_safe(sh,'A4:G4')
    for rr in [5,6,7,8]:
        for cc in ['C','E','G']:
            sh.get_range(f'{cc}{rr}').clear({'apply_to':'contents'})
    for rr in [5,6,7,8]:
        for left,right in [('B','C'),('D','E'),('F','G')]:merge_safe(sh,f'{left}{rr}:{right}{rr}')
    sh.get_range('A3').values=[[None]]  # remove stray template placeholder
    sh.get_range('A4').values=[[project+' 종합']]
    sh.get_range('C3').values=[[f'보완기준일: {when} / {round_no}차 보완 (검토 초안)']]
    text_style(sh,'A4:G4',bold=True,fill=FILL,height=28)
    text_style(sh,'A5:G5',bold=True,fill=FILL2,height=25)
    for rr,kind in [(6,'file'),(7,'struct'),(8,'format')]:
        rows=[10+6*idx+{'file':2,'struct':3,'format':4}[kind] for idx in range(group_total)]
        for col in ['B','D','F']:
            sh.get_range(f'{col}{rr}').formulas=[['=SUM('+','.join(f'{col}{r}' for r in rows)+')']]
        text_style(sh,f'A{rr}:G{rr}',height=22)
    sh.get_range('A:A').format.column_width=82
    for cc in 'BCDEFG':sh.get_range(f'{cc}:{cc}').format.column_width=15
    sh.get_range('A2:G3').format.font.color=BLACK
    sh.freeze_panes.freeze_rows(5)
    return sh


def make_overview(wb,temp,idx,g,when,round_no):
    sh=wb.worksheets.add(f'{idx}. 보완 종합 현황')
    copy_range(sh,temp,'A1:G13','A1:G13')
    # Template overview merged areas are recreated only where needed, independent of example text.
    for rr in [5,6,11,12,13]:
        if rr in [11,12,13]:
            for a,b in [('B','C'),('D','E'),('F','G')]:merge_safe(sh,f'{a}{rr}:{b}{rr}')
    sh.get_range('B5').values=[[g['name']]]
    sh.get_range('B6').values=[[f'{when} / {round_no}차 보완 (검토 초안)']]
    for rr,k in [(11,'file'),(12,'struct'),(13,'format')]:
        total,errors=g['counts'][k]
        sh.get_range(f'B{rr}').values=[[total]]
        sh.get_range(f'D{rr}').values=[[errors]]
        sh.get_range(f'F{rr}').values=[[0]]
    sh.get_range('G11').values=[[f'규칙 불일치 {len(g["differences"])}셀 · 검토 필요' if g['differences'] else '규칙 일치']]
    sh.get_range('G12').values=[['원본 미검증 · 보완 완료 0건']]
    sh.get_range('G13').values=[['원본 재검증 후 집계']]
    sh.get_range('A1:G13').format.font.color=BLACK
    for rr in [2,4,9]:text_style(sh,f'A{rr}:G{rr}',bold=True,fill=FILL,height=26)
    for rr in [10]:text_style(sh,f'A{rr}:G{rr}',bold=True,fill=FILL2,height=25)
    sh.get_range('A:A').format.column_width=22
    sh.get_range('B:B').format.column_width=38
    for c in 'CDEFG':sh.get_range(f'{c}:{c}').format.column_width=22
    return sh


def issue_row(kind,cells,fn):
    return fn(kind,0,cells,[])


def make_issue(wb,temp,idx,g,kind,tab,rowbuilder):
    sh=wb.worksheets.add(f'{idx}. {tab}')
    minimum={'file':6,'struct':5,'format':5}[kind]
    copy_range(sh,temp,'A1:F'+str(minimum),'A1:F'+str(minimum))
    merge_safe(sh,'A1:C1')
    # Previous template's sample stats and records are NOT copied.
    sh.get_range(f'A3:F{minimum}').values=[[None]*6 for _ in range(minimum-2)]
    data=g['errors'][kind]
    counts=Counter(str(cells.get(3) if kind=='file' else cells.get(2)) for _,cells in data)
    for k,(label,num) in enumerate(sorted(counts.items()),3):
        sh.get_range(f'A{k}:C{k}').values=[[label,num,0]]
    header_row=max(minimum,4+len(counts))
    heads=(['파일경로','파일명','보완 항목','보완 전','보완 후','비고'] if kind=='file'
           else ['파일명','오류 유형','프로퍼티','보완 전','보완 후','비고'])
    sh.get_range(f'A{header_row}:F{header_row}').values=[heads]
    text_style(sh,'A1:C1',bold=True,fill=FILL,height=26)
    text_style(sh,'A2:C2',bold=True,fill=FILL2,height=24)
    for row in range(3,3+len(counts)):
        text_style(sh,f'A{row}:C{row}',height=21)
    text_style(sh,f'A{header_row}:F{header_row}',bold=True,fill=FILL,height=27)
    last=header_row
    if data:
        rows=[rowbuilder(kind,n,cells,g['rules']['inspection']) for n,cells in data]
        for begin in range(0,len(rows),550):
            batch=rows[begin:begin+550]
            start=header_row+begin+1
            sh.get_range(f'A{start}:F{start+len(batch)-1}').values=batch
        last=header_row+len(rows)
        body=sh.get_range(f'A{header_row+1}:F{last}')
        body.format.font.color=BLACK
        body.format.row_height=32
        sh.get_range(f'D{header_row+1}:F{last}').format.wrap_text=True
    for col,w in zip('ABCDEF',[28,25,26,57,73,43]):
        sh.get_range(f'{col}:{col}').format.column_width=w
    sh.freeze_panes.freeze_rows(header_row)
    return sh,last


def make_rule(wb,idx,which,g):
    sh=wb.worksheets.add(f'{idx}. {"구축기업 구문규칙" if which=="company" else "검사 구문규칙"}')
    mat=g['rules'][which]
    nrow=len(mat)
    ncol=max(map(len,mat))
    for start in range(0,nrow,100):
        batch=mat[start:start+100]
        sh.get_range(f'A{start+1}:{_col(ncol)}{start+len(batch)}').values=[list(row)+[None]*(ncol-len(row)) for row in batch]
    # Original rule content starts at row 1: row1 categories, row2 columns, row3 data.
    type_idx=next((i+1 for i,v in enumerate(mat[1]) if v=='타입'),6 if ncol>=15 else 5)
    sh.merge_cells(f'A1:{_col(type_idx-1)}1')
    sh.merge_cells(f'{_col(type_idx)}1:{_col(ncol)}1')
    text_style(sh,f'A1:{_col(ncol)}1',bold=True,fill=FILL,height=31)
    text_style(sh,f'A2:{_col(ncol)}2',bold=True,fill=FILL2,height=27)
    if nrow>=3:
        sh.get_range(f'A3:{_col(ncol)}{nrow}').format.font={'name':'Malgun Gothic','color':BLACK}
        sh.get_range(f'A3:{_col(ncol)}{nrow}').format.row_height=23
    for i in range(1,ncol+1):
        col=_col(i)
        width=17 if i<type_idx else (45 if i>=ncol-1 else 18)
        sh.get_range(f'{col}:{col}').format.column_width=width
    if nrow>=7:
        sh.get_range('7:7').format.row_height=54
    if nrow>=11 and ncol>=15:
        sh.get_range('11:11').format.row_height=48
    if nrow>=3 and ncol>=14:
        sh.get_range(f'{_col(ncol-1)}3:{_col(ncol)}{nrow}').format.wrap_text=True
    sh.freeze_panes.freeze_rows(2)
    return sh


def create_report(template,output,project,groups,when,round_no,rowbuilder):
    tmpl=SpreadsheetFile.import_xlsx(Blob.load(str(template)))
    wb=Workbook.create()
    make_summary(wb,tmpl.worksheets.get_item('데이터 그룹 보완 종합 현황'),project,groups,when,round_no)
    results=[]
    for idx,g in enumerate(groups,1):
        make_overview(wb,tmpl.worksheets.get_item('보완 종합 현황'),idx,g,when,round_no)
        for kind,tab in [('file','파일 보완 목록'),('struct','구조 보완 목록'),('format','형식 보완 목록')]:
            sh,last=make_issue(wb,tmpl.worksheets.get_item(tab),idx,g,kind,tab,rowbuilder)
            results.append((idx,tab,last))
        for which in ['company','inspection']:
            make_rule(wb,idx,which,g)
    SpreadsheetFile.export_xlsx(wb).save(str(output))
    return results
