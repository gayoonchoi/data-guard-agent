"""DataGuard Step 4 v1.1: inspection note-aware matching + evidence-based before/after draft.

READ ONLY: does not alter original data. No network/LLM calls.
'보완 후' is a REQUIRED CHANGE / DRAFT, never an assertion that a change was made.
For JSON arrays, a trailing '(N)' is interpreted as a 1-based record number,
which must be confirmed against the dataset's own layout.
"""
from __future__ import annotations

from collections import Counter
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
import csv
import hashlib
import importlib
import json
import re
import uuid
import zipfile

from data_loader_matcher import match_inspection, save_results

BEFORE_COLUMNS = [
    '사업ID', '규칙그룹', '검사결과서 시트', '검사결과서 행', '파일명(원문)',
    '원본 상대경로', '행표기', '검사 프로퍼티', '검사 항목', '검사 규칙',
    '매칭상태', '실제 데이터 위치', '원본 레코드 해석', '현재 검증 상태',
    '보완 전(실제값/미리보기)', '원본값 SHA256', '보완 후(규칙상 요구사항/초안)',
    '담당자 조치', '보완 완료 여부',
]
PREVIEW_LIMIT = 1500
MAX_JSON_IN_MEMORY = 100 * 1024 * 1024
MAX_ZIP_MEMBER = 120 * 1024 * 1024
REGEX_LIMIT = 2000


def _render(value):
    s = json.dumps(value, ensure_ascii=False, default=str)
    if len(s) <= PREVIEW_LIMIT:
        return s
    return s[:PREVIEW_LIMIT] + f' ... [총 {len(s)}자, 원본 파일에서 전체 확인]'


def _digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, default=str).encode('utf-8')).hexdigest()


@contextmanager
def _read_source_file(mode, path, source):
    if mode == 'folder':
        root = Path(source).resolve()
        resolved = (root / path).resolve()
        if not resolved.is_relative_to(root) or not resolved.is_file():
            raise ValueError('허용된 폴더에서 파일을 찾을 수 없습니다.')
        with resolved.open('rb') as f:
            yield f, resolved.stat().st_size
    else:
        with zipfile.ZipFile(source) as z:
            info = z.getinfo(path)
            if info.file_size > MAX_ZIP_MEMBER:
                raise ValueError('ZIP 내부 파일이 너무 큽니다. 로컬 폴더로 검사해 주세요.')
            with z.open(info, 'r') as f:
                yield f, info.file_size


def _pick_record(doc, number):
    """Select a row from an array of records, without guessing among several arrays."""
    if number is None:
        return doc, '전체 JSON 문서'
    if number < 1:
        raise ValueError('행번호 0: 1부터 시작하는지 데이터 명세 확인 필요')
    if isinstance(doc, list):
        if number > len(doc):
            raise ValueError(f'행번호 {number}가 JSON 레코드 수 {len(doc)} 초과')
        return doc[number - 1], f'JSON 배열 {number}번째(1-based 가정)'
    if isinstance(doc, dict):
        candidates = [(k,v) for k,v in doc.items() if isinstance(v, list) and
                      k in ('data', 'records', 'items', 'examples', 'rows')]
        if len(candidates) == 1:
            k, arr = candidates[0]
            if number > len(arr):
                raise ValueError(f'행번호 {number}가 {k} 배열 길이 {len(arr)} 초과')
            return arr[number - 1], f'{k} 배열 {number}번째(1-based 가정)'
        raise ValueError('JSON 객체의 레코드 배열이 불명확합니다. 담당자가 데이터 구조를 지정해야 합니다.')
    raise ValueError('JSON 자료구조에서 레코드 번호를 해석할 수 없습니다.')


def _jsonl_record(stream, n):
    if n is None:
        raise ValueError('JSONL은 파일명 뒤 (행번호)가 필요합니다.')
    if n < 1:
        raise ValueError('JSONL 행번호는 1 이상이어야 합니다.')
    for idx, line in enumerate(stream, 1):
        if idx == n:
            return json.loads(line.decode('utf-8-sig')), f'JSONL 실제 {n}번째 줄'
    raise ValueError(f'JSONL {n}번째 줄이 없습니다.')


def _stream_json_record(stream, n):
    """Optional ijson for large JSON; no entire large file in memory."""
    try:
        import ijson
    except ImportError as exc:
        raise ValueError('대용량 JSON은 스트리밍을 위해 pip install ijson 필요') from exc
    for prefix in ('item', 'data.item', 'records.item', 'items.item', 'examples.item', 'rows.item'):
        stream.seek(0)
        try:
            for idx, item in enumerate(ijson.items(stream, prefix), 1):
                if idx == n:
                    return item, f'{prefix} {n}번째(1-based 가정)'
                if idx > n:
                    break
        except Exception:
            continue
    raise ValueError('대용량 JSON에서 행번호를 찾을 수 없습니다. 배열 구조/행 기준 확인 필요')


def read_value(mode, path, source, row_str, property_path):
    """Returns (before_value, interpretation). Value is never automatically corrected."""
    ext = Path(path).suffix.lower()
    n = int(row_str) if str(row_str).strip() else None
    with _read_source_file(mode, path, source) as (stream, size):
        if ext == '.jsonl':
            record, desc = _jsonl_record(stream, n)
        elif ext == '.json':
            if size > MAX_JSON_IN_MEMORY and n is not None:
                record, desc = _stream_json_record(stream, n)
            elif size > MAX_JSON_IN_MEMORY:
                raise ValueError('대용량 JSON에 행번호가 없어 대상 레코드를 특정할 수 없습니다.')
            else:
                document = json.loads(stream.read().decode('utf-8-sig'))
                record, desc = _pick_record(document, n)
        else:
            raise ValueError(f'{ext} 파일의 행별 구문 검증은 아직 미지원(원본 확인 필요)')
    prop = str(property_path or '').strip()
    if not prop:
        raise ValueError('검사 프로퍼티가 없어 수정 대상 값을 특정할 수 없습니다.')
    # Simple dotted + integer subscript: dialogue[1].content or $.text
    prop = prop.removeprefix('$.').removeprefix('$')
    if not re.fullmatch(r'[A-Za-z_][\w-]*(?:\[\d+\]|\.[A-Za-z_][\w-]*)*', prop):
        raise ValueError(f'프로퍼티 표현식 미지원: {prop[:80]}')
    for part in re.findall(r'[^.]+', prop):
        key = part.split('[')[0]
        if not isinstance(record, dict) or key not in record:
            raise KeyError(f'{key} 필드가 해당 레코드에 없습니다.')
        record = record[key]
        for idx in re.findall(r'\[(\d+)\]', part):
            if not isinstance(record, list):
                raise ValueError(f'{key} 필드가 배열이 아닙니다.')
            record = record[int(idx)]
    return record, desc


def determine_issue(value, rule_text):
    """Conservative rule checking. Complex regex is NEVER auto-accepted."""
    src = str(rule_text or '').strip()
    if not src:
        return '규칙 미기재 · 확인 필요', '검사결과서의 구문규칙과 담당자 검토 후 수정안을 작성'
    pattern = src.split(':', 1)[1].strip() if src.startswith('유효패턴:') else ''
    if not pattern:
        return '규칙의 자동 해석 미지원', f'검사결과서 규칙 [{src[:160]}]에 맞도록 수정 요청'
    # Whitespace checks are explicitly deterministic.
    if pattern in (r'\S', r'\S+', r'^\S+$'):
        compliant = isinstance(value, str) and bool(re.search(pattern, value))
        return ('현재 규칙 통과' if compliant else '현재 규칙 위반',
                '공백이 아닌 실제 문자를 포함하도록 content 작성(담당자 확인 필요)')
    if (isinstance(value, str) and len(value) <= REGEX_LIMIT and len(pattern) <= 90
            and not any(x in pattern for x in ('(?', '\\s\\S', '(?:'))):
        try:
            compliant = bool(re.search(pattern, value))
            return ('현재 규칙 통과' if compliant else '현재 규칙 위반',
                    f'정규식 {pattern}을 만족하도록 값을 수정(임의 문구 생성 금지)')
        except re.error:
            pass
    # Structural hint for the 022 dialogue pattern; not a full regex proof.
    if '<참고문서>' in pattern and '질문:' in pattern and isinstance(value, str):
        if '<참고문서>' not in value or '질문:' not in value:
            return ('형식 요소 누락 의심 · 검토 필요',
                    '구문규칙에서 요구하는 <참고문서> 및 질문: 구성 확인 후 누락 요소 보완')
        reference = value.split('<참고문서>', 1)[1].split('질문:', 1)[0].strip()
        question = value.split('질문:', 1)[1].strip()
        if not reference:
            return ('참고문서 내용 누락 의심 · 검토 필요',
                    '<참고문서>와 질문: 사이의 실제 근거 내용을 원천 데이터 담당자에게 확인해 보완')
        if not question:
            return ('질문 내용 누락 의심 · 검토 필요',
                    '질문: 뒤에 공백이 아닌 실제 질문 내용을 원천 데이터 담당자에게 확인해 보완')
    # Pattern with many negative lookaheads, huge texts: avoid ReDoS / semantic invention.
    return ('복잡한 정규식 · 담당자 검토 필요',
            f'검사결과서 유효패턴을 만족하도록 보완해야 함: {pattern[:180]} [원천 근거를 확인한 뒤 수정]')


def analyze(inspection, source, mode='folder', output_dir='outputs/step4_evidence',
            treat_absent_as_removed=False):
    rows, summary = match_inspection(inspection, folder=source if mode == 'folder' else None,
                                     archive=source if mode == 'zip' else None)
    if treat_absent_as_removed and summary.get('file_count', 0) == 0:
        raise ValueError('선택한 최종 데이터 폴더/ZIP에 검증 대상 파일이 하나도 없습니다. 폴더 경로를 확인하세요.')
    results = []
    counters = Counter()
    for row in rows:
        out = {**row, '원본 레코드 해석': '', '현재 검증 상태': '검토 필요',
               '보완 전(실제값/미리보기)': '', '원본값 SHA256': '',
               '보완 후(규칙상 요구사항/초안)': '', '담당자 조치': '', '보완 완료 여부': '미완료'}
        if row['매칭상태'] == '파일 없음' and treat_absent_as_removed:
            # Explicitly enabled for an authoritative FINAL data folder/ZIP only.
            # A missing file is NOT evidence that its original field was repaired.
            out['현재 검증 상태'] = '삭제·제외 간주(최종 폴더 부재)'
            out['보완 전(실제값/미리보기)'] = (
                '검사결과서 기록: ' + str(row.get('검사 프로퍼티', '')) +
                ' = ' + str(row.get('오류 메시지', '') or '(오류값 별도 확인)') +
                ' / 검사 규칙: ' + str(row.get('검사 규칙', ''))
            )
            out['보완 후(규칙상 요구사항/초안)'] = (
                '해당 오류 파일이 최종 보완 데이터셋에서 확인되지 않아 '
                '삭제·제외된 것으로 간주하며, 파일 내부 값 수정 대상에서 제외함.'
            )
            out['담당자 조치'] = (
                '최종 보완 데이터셋에 해당 파일이 없음을 확인; '
                '삭제 이력은 검사결과서만으로 확인할 수 없음'
            )
            out['보완 완료 여부'] = '삭제·제외 간주(확인 필요)'
        elif row['매칭상태'] != '일치':
            out['현재 검증 상태'] = '파일 연결 실패'
            out['담당자 조치'] = f'매칭 상태 확인: {row["매칭상태"]}'
        else:
            try:
                before, interpretation = read_value(mode, row['실제 데이터 위치'], source,
                                                    row['행표기'], row['검사 프로퍼티'])
                state, suggested = determine_issue(before, row['검사 규칙'])
                out['원본 레코드 해석'] = interpretation
                out['보완 전(실제값/미리보기)'] = _render(before)
                out['원본값 SHA256'] = _digest(before)
                out['현재 검증 상태'] = state
                out['보완 후(규칙상 요구사항/초안)'] = suggested
                out['담당자 조치'] = ('구문 규칙과 데이터의 행 기준 확인; 실제 수정·재검증 후 보완 완료 판정'
                                    if row['행표기'] else '구문 규칙에 따라 담당자 검토')
            except Exception as exc:
                out['현재 검증 상태'] = '대상 위치 확인 필요'
                out['담당자 조치'] = f'{type(exc).__name__}: {str(exc)[:180]}'
        counters[out['현재 검증 상태']] += 1
        results.append(out)
    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + uuid.uuid4().hex[:6]
    name = f'{summary["project_id"]}_evidence_draft_{stamp}.csv'
    csv_path = destination / name
    with csv_path.open('w', encoding='utf-8-sig', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=BEFORE_COLUMNS, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(results)
    meta = {**summary, 'evidence_status_counts': dict(counters),
            'all_originals_unchanged': True, 'approved_corrections': 0,
            'note': '행번호는 데이터셋의 논리 레코드 기준과 검증해야 함. 수정 지시 초안이지 실제 수정 완료 아님.'}
    summary_path = destination / f'{summary["project_id"]}_evidence_summary_{stamp}.json'
    summary_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding='utf-8')
    return {'rows': results, 'summary': meta, 'csv': str(csv_path), 'json': str(summary_path)}


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='DataGuard: 보완 전 실제값과 보완 지시 초안 (원본 불변)')
    parser.add_argument('--inspection', required=True)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--folder')
    group.add_argument('--zip', dest='archive')
    parser.add_argument('--out', default='outputs/step4_evidence')
    args = parser.parse_args()
    result = analyze(args.inspection, args.folder or args.archive,
                     mode='folder' if args.folder else 'zip', output_dir=args.out)
    print(json.dumps(result['summary']['evidence_status_counts'], ensure_ascii=False))
    print(result['csv'])
