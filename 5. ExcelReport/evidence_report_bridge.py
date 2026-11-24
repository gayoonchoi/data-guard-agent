# -*- coding: utf-8 -*-
"""DataGuard ④ 검증 CSV -> ③ 보완결과서 초안 (서식 무손실 연결).

Only writes D(보완 전), E(보완 후 수정지시), F(비고) of matching detail rows.
Never modifies rule sheets, original labeling data, totals, or claims completion.
Uses OOXML cell-value patching to preserve the source report's Excel styles exactly.
"""
from __future__ import annotations

import csv
import html
import importlib.util
import os
import re
import tempfile
import unicodedata
import zipfile
from collections import Counter
from pathlib import Path

SHEET_KINDS = {'파일': '파일', '구조': '구조', '형식': '형식'}
CELL_RE = re.compile(r'<c\b[^>]*\br="([A-Z]+\d+)"[^>]*(?:/>|>.*?</c>)', re.DOTALL)
ROW_RE = re.compile(r'<row\b(?=[^>]*\br="(\d+)")[^>]*>.*?</row>', re.DOTALL)
BAD_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _load_report_module():
    module_file = Path(__file__).with_name('report_generator.py')
    if not module_file.is_file():
        raise FileNotFoundError(f'보고서 생성기 누락: {module_file}')
    spec = importlib.util.spec_from_file_location('dg_report_bridge_reader', module_file)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _norm(value):
    return unicodedata.normalize('NFKC', str(value or '')).strip()


def _evidence_rows(csv_file, project_id):
    path = Path(csv_file)
    if not path.is_file() or path.stat().st_size > 500 * 1024**2:
        raise ValueError('④에서 생성한 정상적인 CSV 파일을 지정해 주세요. (최대 500MB)')
    rows = {}
    with path.open('r', encoding='utf-8-sig', newline='') as stream:
        reader = csv.DictReader(stream)
        required = {'사업ID', '규칙그룹', '검사결과서 시트', '검사결과서 행',
                    '파일명(원문)', '매칭상태', '보완 전(실제값/미리보기)',
                    '보완 후(규칙상 요구사항/초안)', '현재 검증 상태'}
        if not required.issubset(set(reader.fieldnames or [])):
            raise ValueError('④ 보완 전후 CSV 형식이 아닙니다. 먼저 ④에서 검토 초안을 생성해 주세요.')
        for rec in reader:
            if str(rec.get('사업ID', '')).zfill(3) != project_id:
                raise ValueError('④ CSV와 ③ 검사결과서의 사업 ID가 다릅니다. 잘못된 값 반영을 중단합니다.')
            sheet_name = _norm(rec['검사결과서 시트'])
            m = re.fullmatch(r'(\d+)\.\s*(파일|구조|형식)\s*오류\s*목록', sheet_name)
            if not m:
                raise ValueError(f'CSV의 검사결과서 시트 이름을 해석하지 못했습니다: {sheet_name}')
            if str(rec['규칙그룹']).strip() != str(int(m.group(1))):
                raise ValueError(f'CSV 규칙그룹/시트 불일치: {sheet_name}')
            try:
                source_row = int(rec['검사결과서 행'])
            except (TypeError, ValueError) as exc:
                raise ValueError('검사결과서 행번호가 숫자가 아닙니다.') from exc
            key = (int(m.group(1)), m.group(2), source_row)
            if key in rows:
                raise ValueError(f'CSV에 동일 오류 식별키가 중복되어 있습니다: {key}')
            rows[key] = rec
    return rows


def _build_target_map(source, report_path, csv_rows, reader_module):
    inspector = reader_module.WorkbookReader(source)
    output = reader_module.WorkbookReader(report_path)
    targets = {}  # output sheet -> {excel row -> {D: str, E: str, F: str}}
    stats = Counter()
    try:
        pid = reader_module.project_id_from_inspection(source)
        groups = reader_module.extract_all(inspector)
        # The report generator uses groups sorted by group id, and preserves their error list order.
        for index, group in enumerate(groups, 1):
            for kind, kind_name in [('file', '파일'), ('struct', '구조'), ('format', '형식')]:
                issues = group['errors'][kind]
                if not issues:
                    continue
                report_sheet = f'{index}. {kind_name} 보완 목록'
                if report_sheet not in output.sheets:
                    raise ValueError(f'생성된 보완결과서에 시트가 없습니다: {report_sheet}')
                row_data = {}
                header_row = None
                for r, cells in output.rows(report_sheet):
                    row_data[r] = cells
                    if kind == 'file':
                        found = _norm(cells.get(1)) == '파일경로' and _norm(cells.get(2)) == '파일명'
                    else:
                        found = _norm(cells.get(1)) == '파일명' and _norm(cells.get(2)) == '오류 유형'
                    if found:
                        header_row = r
                if header_row is None:
                    raise ValueError(f'{report_sheet}에서 상세 오류 표의 제목행을 찾을 수 없습니다.')
                for pos, (source_row, source_cells) in enumerate(issues):
                    key = (group['id'], kind_name, source_row)
                    evidence = csv_rows.get(key)
                    if evidence is None:
                        stats['CSV에 없는 항목'] += 1
                        continue
                    original_name = source_cells.get(2) if kind == 'file' else source_cells.get(1)
                    if _norm(original_name) != _norm(evidence['파일명(원문)']):
                        raise ValueError(f'{key}: 원본 파일명이 CSV와 다릅니다. 자동 반영을 중단했습니다.')
                    report_row = header_row + pos + 1
                    actual = row_data.get(report_row, {})
                    expected = reader_module._issue_row(kind, source_row, source_cells,
                                                         group['rules']['inspection'])
                    if _norm(actual.get(1)) != _norm(expected[0]):
                        raise ValueError(f'{report_sheet}!A{report_row} 상세 목록 순서가 바뀌었습니다. 잘못된 셀 쓰기를 차단했습니다.')
                    # Note: state '현재 규칙 통과' is not '보완 완료'. Original remains unchanged.
                    matched = evidence.get('매칭상태') == '일치'
                    state = _norm(evidence.get('현재 검증 상태'))
                    observed = evidence.get('보완 전(실제값/미리보기)', '').strip()
                    suggestion = evidence.get('보완 후(규칙상 요구사항/초안)', '').strip()
                    replacements = {}
                    if matched and observed and state not in ('파일 연결 실패', '대상 위치 확인 필요'):
                        replacements['D'] = '현재 원본 확인값: ' + observed
                        if suggestion:
                            replacements['E'] = suggestion
                        stats['실제값 확인'] += 1
                    else:
                        stats['검토 대기'] += 1
                    # Keep existing notes for traceability; append current review state.
                    existing_note = str(actual.get(6) or '').strip()
                    if state:
                        review_note = f'검토: {state} · 미완료'
                        replacements['F'] = (existing_note + ' / ' + review_note).strip(' /')
                    if replacements:
                        targets.setdefault(report_sheet, {})[report_row] = replacements
                    stats['대조 완료'] += 1
        unexpected = set(csv_rows) - {
            (g['id'], name, rowid)
            for g in groups for kind, name in [('file', '파일'), ('struct', '구조'), ('format', '형식')]
            for rowid, _ in g['errors'][kind]
        }
        if unexpected:
            raise ValueError(f'CSV에 원본 검사결과서와 다른 오류 행이 {len(unexpected)}건 있습니다. 자동 반영을 중단합니다.')
        stats['대상 오류'] = sum(len(v) for g in groups for v in g['errors'].values())
        return targets, dict(stats), output.sheets
    finally:
        inspector.close()
        output.close()


def _inline_cell(coord, old_xml, value):
    # Preserve style number (s) and any non-value attributes: only replace the text value.
    xml = old_xml or f'<c r="{coord}"/>'
    head = re.match(r'<c\b([^>]*?)(?:/?>)', xml, re.DOTALL)
    if not head:
        raise RuntimeError(f'보완결과서 셀 XML을 해석할 수 없습니다: {coord}')
    attrs = head.group(1)
    attrs = re.sub(r'\s+t="[^"]*"', '', attrs)
    text = BAD_XML.sub('', str(value))[:32767]
    return f'<c{attrs} t="inlineStr"><is><t xml:space="preserve">{html.escape(text, quote=False)}</t></is></c>'


def _patch_sheet_xml(payload, row_updates):
    source = payload.decode('utf-8')
    changed = 0
    def replace_row(found):
        nonlocal changed
        raw = found.group(0)
        row_num = int(found.group(1))
        cells_to_change = row_updates.get(row_num)
        if not cells_to_change:
            return raw
        for col, val in cells_to_change.items():
            coord = f'{col}{row_num}'
            cell_pattern = re.compile(r'<c\b(?=[^>]*\br="' + re.escape(coord) +
                                      r'")[^>]*?(?:/>|>.*?</c>)', re.DOTALL)
            result = cell_pattern.search(raw)
            replacement = _inline_cell(coord, result.group(0) if result else '', val)
            if result:
                raw = raw[:result.start()] + replacement + raw[result.end():]
            else:
                raw = raw.replace('</row>', replacement + '</row>', 1)
            changed += 1
        return raw
    updated = ROW_RE.sub(replace_row, source)
    if changed != sum(len(v) for v in row_updates.values()):
        raise ValueError('원본 보완결과서와 오류 행 수가 맞지 않습니다. 생성 파일을 수정하지 않았습니다.')
    return updated.encode('utf-8')


def apply_evidence(report_path, inspection_path, evidence_csv, output_dir=None):
    """Return new Excel with unchanged styles/rule sheets; do not modify the original report."""
    report = Path(report_path)
    inspection = Path(inspection_path)
    if not report.is_file() or not inspection.is_file() or not zipfile.is_zipfile(report):
        raise ValueError('보완결과서 또는 검사결과서 Excel 파일을 찾을 수 없습니다.')
    module = _load_report_module()
    pid = module.project_id_from_inspection(inspection)
    rows = _evidence_rows(evidence_csv, pid)
    changes, stats, output_sheets = _build_target_map(inspection, report, rows, module)
    # Only detail worksheets are touched; rules/summary/styles remain byte-for-byte intact.
    replacements = {output_sheets[name]: update for name, update in changes.items()}
    destination = Path(output_dir) if output_dir else report.parent / 'evidence_applied'
    destination.mkdir(parents=True, exist_ok=True)
    out = destination / report.name
    fd, tmpname = tempfile.mkstemp(prefix='._dg_review_', suffix='.xlsx', dir=destination)
    os.close(fd)
    tmp = Path(tmpname)
    try:
        with zipfile.ZipFile(report, 'r') as zsrc, zipfile.ZipFile(tmp, 'w') as zdst:
            for item in zsrc.infolist():
                raw = zsrc.read(item.filename)
                if item.filename in replacements:
                    raw = _patch_sheet_xml(raw, replacements[item.filename])
                zdst.writestr(item, raw)
        with zipfile.ZipFile(tmp) as verify:
            if verify.testzip():
                raise ValueError('저장한 Excel ZIP 구조가 손상되었습니다.')
            for name in replacements:
                verify.read(name)
        # If workbook open/locked on Windows, fail safely; never overwrite original.
        os.replace(tmp, out)
    finally:
        tmp.unlink(missing_ok=True)
    return {'path': str(out), 'project_id': pid, 'statistics': stats}
