# -*- coding: utf-8 -*-
"""Apply the actual supplied TTA correction-report sample's Excel style to a generated report.

Content is never regenerated or modified: only styles, row heights and column widths are
changed. Python standard library ZIP/XML utilities are used,
so no second in-memory copy of tens of thousands of issue records is required.

The output keeps all 019-specific content, including original JSON/TXT rules, formulas,
counts, issue details and workbook sheet names.
"""
from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZipFile
import os
import posixpath
import re

from xml.etree import ElementTree as ET

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _sheets(z):
    book = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    by_id = {r.get("Id"): r.get("Target") for r in rels}
    result = {}
    for sh in book.find(X + "sheets"):
        target = by_id[sh.get(R + "id")]
        path = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
        result[sh.get("name")] = path
    return result


def _cell_row_col(address):
    m = re.match(r"^([A-Z]+)(\d+)$", address)
    if not m:
        return None
    n = 0
    for ch in m.group(1):
        n = 26 * n + ord(ch) - 64
    return int(m.group(2)), n


def _source_styles(sheet_xml):
    root = ET.fromstring(sheet_xml)
    styles = {}
    heights = {}
    for row in root.findall(".//" + X + "sheetData/" + X + "row"):
        rn = int(row.get("r"))
        if row.get("ht"):
            heights[rn] = (row.get("ht"), row.get("customHeight", "1"))
        for cell in row.findall(X + "c"):
            rc = _cell_row_col(cell.get("r", ""))
            if rc:
                styles[rc] = cell.get("s", "0")
    widths = root.find(X + "cols")
    return styles, heights, widths


def _lookup(styles, row, col):
    return styles.get((row, col), "0")


def _style_for(sheet_name, row, col, standard, first_data_header):
    """Map each generated cell to the same semantic place in the user's sample."""
    if sheet_name == "데이터 그룹 보완 종합 현황":
        if row <= 8:
            return _lookup(standard, row, col)
        if row >= 10:
            phase = (row - 10) % 6
            return _lookup(standard, 10 + phase, col) if phase < 5 else "0"
        return "0"

    group_match = re.match(r"^\d+\. (.+)$", sheet_name)
    if not group_match:
        return "0"
    kind = group_match.group(1)

    if kind == "보완 종합 현황":
        return _lookup(standard, row, col)

    if kind in ("구축기업 구문규칙", "검사 구문규칙"):
        # Exactly match the sample: light-gray group title and column headings,
        # no dark bars in rule worksheets; original nested columns stay unchanged.
        if row == 1:
            return _lookup(standard, 1, 1)  # sample's light-gray title fill
        if row == 2:
            return _lookup(standard, 2, 1)  # sample's light-gray headings
        return _lookup(standard, 3, 1)      # sample's white body cells

    if kind in ("파일 보완 목록", "구조 보완 목록", "형식 보완 목록"):
        source_header = 6 if kind == "파일 보완 목록" else 5
        source_body = source_header + 1
        if row == 1:
            return _lookup(standard, 1, min(col, 3)) if col <= 3 else "0"
        if row == 2:
            return _lookup(standard, 2, min(col, 3)) if col <= 3 else "0"
        if row == first_data_header:
            return _lookup(standard, source_header, col)
        if row > first_data_header:
            return _lookup(standard, source_body, col)
        if 3 <= row < first_data_header:
            # Rows with type counts follow sample's count styling, last spacer is blank.
            if row == first_data_header - 1:
                return _lookup(standard, min(source_header - 1, 4), col)
            return _lookup(standard, 3, col) if col <= 3 else "0"
    return "0"


def _find_header(root, kind):
    for row in root.findall(".//" + X + "sheetData/" + X + "row"):
        idx = int(row.get("r"))
        if idx <= 2:
            continue
        if idx > 40:
            break
        names = []
        for cell in row.findall(X + "c"):
            if cell.get("r") != "A" + str(idx):
                continue
            inline = cell.find(X + "is")
            if inline is not None:
                names = ["".join(inline.itertext())]
        if names and names[0] in ("파일명", "파일경로"):
            return idx
    return 6 if kind == "파일 보완 목록" else 5


def _new_cols(source_cols, kind, output_rule_width=0, prefix=""):
    if source_cols is None:
        return None
    if kind in ("구축기업 구문규칙", "검사 구문규칙") and output_rule_width:
        specs = []
        for i in range(1, output_rule_width + 1):
            if output_rule_width >= 15:
                widths = [10, 10, 16, 10, 10, 13, 12, 38, 13, 13, 36, 14, 16, 49, 44]
                w = widths[min(i - 1, len(widths) - 1)]
            else:
                w = 12 if i == 1 else (15 if i <= 4 else 20)
                if i == output_rule_width - 1: w = 45
                if i == output_rule_width: w = 38
            specs.append(f'<{prefix}col min="{i}" max="{i}" width="{w}" customWidth="1"/>')
        return (f'<{prefix}cols>' + ''.join(specs) + f'</{prefix}cols>').encode('utf-8')
    specs=[]
    for col in source_cols:
        attrs=' '.join(f'{k}="{v}"' for k,v in col.attrib.items())
        specs.append(f'<{prefix}col {attrs}/>')
    return (f'<{prefix}cols>' + ''.join(specs) + f'</{prefix}cols>').encode('utf-8')


def _set_attr(raw, name, value):
    bname=name.encode('ascii')
    pat=rb'\s+'+bname+rb'="[^"]*"'
    repl=b' '+bname+b'="'+str(value).encode('ascii')+b'"'
    return (re.sub(pat, repl, raw, count=1) if re.search(pat,raw) else
            (raw[:-2].rstrip()+repl+b'/>' if raw.endswith(b'/>') else raw[:-1]+repl+b'>'))


def _restore_template_merges(raw, kind):
    """Reinstate important merged regions that a prior generator didn't copy."""
    required = {
        '데이터 그룹 보완 종합 현황': ['A2:G2', 'C3:G3', 'A4:G4'],
        '보완 종합 현황': ['A2:G2', 'A4:G4', 'A9:G9', 'B5:G5', 'B6:G6'],
    }.get(kind, [])
    if not required:
        return raw
    pref=re.search(rb'<([A-Za-z_][A-Za-z0-9_.-]*:)?worksheet\b',raw)
    prefix=pref.group(1) if pref and pref.group(1) else b''
    tag=prefix+b'mergeCells'
    merged= re.search(rb'<'+tag+rb'(?=[\s>])[^>]*>.*?</'+tag+rb'>', raw, flags=re.DOTALL)
    existing=set()
    if merged:
        existing={x.decode('ascii') for x in re.findall(rb'<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?mergeCell\s+ref="([^"]+)"',merged.group(0))}
    new=[x for x in required if x not in existing]
    if not new:
        return raw
    nodes=b''.join(b'<'+prefix+b'mergeCell ref="'+x.encode('ascii')+b'"/>' for x in new)
    if merged:
        part=merged.group(0)
        part=part.replace(b'</'+tag+b'>', nodes+b'</'+tag+b'>')
        total=len(existing)+len(new)
        part=re.sub(rb' count="[0-9]+"',b' count="'+str(total).encode('ascii')+b'"',part,count=1)
        return raw[:merged.start()]+part+raw[merged.end():]
    tail=b'</'+prefix+b'sheetData>'
    if tail in raw:
        content=b'<'+tag+b' count="'+str(len(new)).encode('ascii')+b'">'+nodes+b'</'+tag+b'>'
        return raw.replace(tail,tail+content,1)
    return raw


def _read_source_rule_title_merges(inspection_path):
    """Get the row-1 header merge coordinates from each original inspection rule sheet.

    The first header may be A1:C1 (JSON) or A1:C1 (TXT), while the
    '설정' header starts at F1 or E1 respectively. Never guess those columns.
    """
    if inspection_path is None:
        return {}
    source_path = Path(inspection_path)
    if not source_path.is_file():
        raise FileNotFoundError(f"검사결과서 파일을 찾을 수 없습니다: {source_path}")
    result = {}
    with ZipFile(source_path, 'r') as input_zip:
        for name, member in _sheets(input_zip).items():
            if not re.fullmatch(r'\d+\. (?:구축기업|검사) 구문규칙', name):
                continue
            rule_sheet = ET.fromstring(input_zip.read(member))
            merged = rule_sheet.find(X + 'mergeCells')
            refs = [] if merged is None else [
                cell.get('ref') for cell in merged
                if re.fullmatch(r'[A-Z]+1:[A-Z]+1', cell.get('ref') or '')
            ]
            if refs:
                result[name] = refs
    return result


def _restore_rule_title_merges(raw, required):
    """Repair only first-row merges. Keep all populated cell XML untouched."""
    if not required:
        return raw
    # Works with both default and prefix-qualified OOXML serialization.
    ws = re.search(rb'<([A-Za-z_][A-Za-z0-9_.-]*:)?worksheet\b', raw)
    prefix = ws.group(1) if ws and ws.group(1) else b''
    tag = prefix + b'mergeCells'
    pattern = rb'<' + tag + rb'(?=[\s>])[^>]*>.*?</' + tag + rb'>'
    existing = re.search(pattern, raw, re.DOTALL)
    other = []
    if existing:
        for value in re.findall(rb'<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?mergeCell\s+ref="([^"]+)"', existing.group(0)):
            ref = value.decode('ascii')
            if not re.fullmatch(r'[A-Z]+1:[A-Z]+1', ref):
                other.append(ref)
    merged = list(dict.fromkeys(other + list(required)))
    nodes = b''.join(
        b'<' + prefix + b'mergeCell ref="' + ref.encode('ascii') + b'"/>'
        for ref in merged
    )
    new_tag = b'<' + tag + b' count="' + str(len(merged)).encode('ascii') + b'">' + nodes + b'</' + tag + b'>'
    if existing:
        return raw[:existing.start()] + new_tag + raw[existing.end():]
    close = b'</' + prefix + b'sheetData>'
    if close not in raw:
        raise ValueError('시트에 sheetData 종료 태그가 없습니다.')
    return raw.replace(close, close + new_tag, 1)


def _patch_sheet_bytes(raw, name, kind, styles, heights, source_cols, header, rule_width, title_merges=None):
    # Preserve EVERY original value/formula/text byte: touch XML opening tags only.
    def fix_cell(m):
        tag=m.group(0)
        pos=re.search(rb'\sr="([A-Z]+[0-9]+)"',tag)
        if not pos:
            return tag
        rc=_cell_row_col(pos.group(1).decode('ascii'))
        if not rc:
            return tag
        val=_style_for(name,rc[0],rc[1],styles,header)
        return _set_attr(tag,'s',val)

    def fix_row(m):
        tag=m.group(0)
        pos=re.search(rb'\sr="([0-9]+)"',tag)
        if not pos:
            return tag
        rn=int(pos.group(1))
        hrow=10+(rn-10)%6 if kind=='데이터 그룹 보완 종합 현황' and rn>=10 else rn
        if kind in ('구축기업 구문규칙','검사 구문규칙'):
            h=heights.get(rn if rn<=2 else 3)
            if rn==7 and rule_width>=15:
                h=('48','1')
        elif kind in ('파일 보완 목록','구조 보완 목록','형식 보완 목록') and rn>header:
            h=('32','1')
        else:
            h=heights.get(hrow)
        if not h:
            return tag
        return _set_attr(_set_attr(tag,'ht',h[0]),'customHeight',h[1])

    raw=re.sub(rb'<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?c\s[^>]*>',fix_cell,raw)
    raw=re.sub(rb'<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?row\s[^>]*>',fix_row,raw)
    ns=re.search(rb'<([A-Za-z_][A-Za-z0-9_.-]*:)?worksheet\b',raw)
    prefix=ns.group(1).decode('ascii') if ns and ns.group(1) else ''
    cols=_new_cols(source_cols,kind,rule_width,prefix)
    if cols is not None:
        existing=re.search(rb'<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?cols(?:\s[^>]*)?>.*?</(?:[A-Za-z_][A-Za-z0-9_.-]*:)?cols>',raw,flags=re.DOTALL)
        if existing:
            raw=raw[:existing.start()]+cols+raw[existing.end():]
        else:
            before=re.search(rb'<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?sheetData(?:\s[^>]*)?>',raw)
            if before:
                raw=raw[:before.start()]+cols+raw[before.start():]
    raw = _restore_template_merges(raw,kind)
    if kind in ('구축기업 구문규칙','검사 구문규칙'):
        raw = _restore_rule_title_merges(raw, title_merges)
    return raw


def apply_template_style(report_path, template_path, inspection_path=None):
    """Style an existing report in-place without touching a single populated value.

    The function is idempotent, and safe for reports with any number of groups.
    It follows the uploaded 7-sheet report template for repeated group sheets.
    """
    report_path = Path(report_path)
    template_path = Path(template_path)
    if not report_path.is_file() or not template_path.is_file():
        raise FileNotFoundError(f"보고서/샘플 양식 확인: {report_path} / {template_path}")
    with ZipFile(template_path, "r") as sample_zip:
        sample_names = _sheets(sample_zip)
        source = {}
        for name, member in sample_names.items():
            source[name] = _source_styles(sample_zip.read(member))
        sample_styles = sample_zip.read("xl/styles.xml")

    original_rule_merges = _read_source_rule_title_merges(inspection_path)

    with ZipFile(report_path, "r") as generated:
        members = generated.infolist()
        book_sheets = _sheets(generated)
        changed = {}
        for name, member in book_sheets.items():
            kind_match = re.match(r"^\d+\. (.+)$", name)
            kind = kind_match.group(1) if kind_match else name
            if kind not in source:
                continue
            st, heights, widths = source[kind]
            sheet_bytes=generated.read(member)
            root=ET.fromstring(sheet_bytes)
            header=_find_header(root,kind)
            output_rule_width=0
            if kind in ("구축기업 구문규칙", "검사 구문규칙"):
                for cell in root.findall(".//" + X + "sheetData/" + X + "row/" + X + "c"):
                    pos=_cell_row_col(cell.get("r", ""))
                    if pos and pos[0] == 2:
                        output_rule_width=max(output_rule_width,pos[1])
            changed[member]=_patch_sheet_bytes(sheet_bytes,name,kind,st,heights,widths,header,output_rule_width,original_rule_merges.get(name))
        changed["xl/styles.xml"] = sample_styles

        with NamedTemporaryFile(prefix="dataguard_styled_", suffix=".xlsx", dir=report_path.parent, delete=False) as fp:
            tmp = Path(fp.name)
        try:
            with ZipFile(tmp, "w") as dst:
                for entry in members:
                    dst.writestr(entry, changed.get(entry.filename, generated.read(entry.filename)))
            # All relevant OOXML was written. Do not overwrite the output on failure.
            with ZipFile(tmp, "r") as check:
                bad = check.testzip()
                if bad:
                    raise ValueError(f"엑셀 파일 내부 검증 실패: {bad}")
        except BaseException:
            # 생성 도중 예외가 나면 임시 파일만 삭제하고 원본은 그대로 둔다.
            if tmp.exists():
                tmp.unlink()
            raise

    # 중요: Windows에서는 위 generated ZipFile 핸들을 닫기 전에
    # os.replace(tmp, report_path)를 실행하면 WinError 5가 발생할 수 있다.
    try:
        try:
            os.replace(tmp, report_path)
            return report_path
        except PermissionError:
            # 기존 보고서가 Excel에서 열려 잠겼다면 파일명을 변경하지 않고
            # 별도 폴더에 저장한다 (사용자 지정 파일명 규칙 유지).
            from datetime import datetime
            import uuid
            folder = report_path.parent / ("재생성_" + datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6])
            folder.mkdir(parents=True, exist_ok=False)
            alternate = folder / report_path.name
            os.replace(tmp, alternate)
            return alternate
    finally:
        if tmp.exists():
            tmp.unlink()


if __name__ == "__main__":
    import argparse
    p=argparse.ArgumentParser()
    p.add_argument("--report", required=True)
    p.add_argument("--template", required=True)
    args=p.parse_args()
    print(apply_template_style(args.report,args.template))
