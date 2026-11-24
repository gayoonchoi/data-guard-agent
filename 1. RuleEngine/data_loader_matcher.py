"""DataGuard Agent · 4단계 파일 매칭 v1.0 (READ ONLY).

019번 검사결과서의 파일 오류/구조 오류/형식 오류 목록을 로컬 폴더 또는
ZIP에 들어 있는 실제 라벨링 파일과 연결한다. 원본 파일을 열거나 수정하지 않는다.
Python 표준 라이브러리만 사용한다. .xlsx는 OOXML ZIP/XML로 읽기 전용 분석.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path, PurePosixPath
import csv
import io
import json
import os
import re
import unicodedata
import uuid
import zipfile
from xml.etree import ElementTree as ET

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
MAIN = lambda tag: '{' + NS_MAIN + '}' + tag
COL_RE = re.compile(r'^([A-Z]+)\d+$', re.I)
SHEET_RE = re.compile(r'^(\d+)\.\s*(파일|구조|형식)\s*오류\s*목록$')
# 검사결과서 비고: '오류 위치는 파일명(행번호)로 작성하였습니다.'
# 022 예시: moel_counsel_mrc_v.01/dialog.json (128), 제4장(2).jsonl (31)
ROW_SUFFIX_RE = re.compile(r'^(?P<name>.+?\.(?:jsonl|json|txt|csv|xlsx))\s*\(\s*(?P<line>\d+)\s*\)$', re.I)
EXT_RE = re.compile(r'\.(jsonl|json|txt|csv|xlsx)$', re.I)
ALLOWED_EXTENSIONS = {'.jsonl', '.json', '.txt', '.csv', '.xlsx'}
MAX_ZIP_ENTRIES = 600_000
MAX_TOTAL_UNCOMPRESSED = 50 * 1024**3  # 50 GiB upper bound for metadata validation
MAX_SINGLE_UNCOMPRESSED = 4 * 1024**3
MAX_COMPRESSION_RATIO = 1000

CSV_COLUMNS = [
    '사업ID', '규칙그룹', '오류구분', '검사결과서 시트', '검사결과서 행',
    '파일명(원문)', '매칭파일명', '원본 상대경로', '행표기', '검사 항목', '검사 프로퍼티',
    '검사 규칙', '오류 메시지', '매칭상태', '실제 데이터 위치', '동일명 후보수', '후보 경로',
]


def _column_number(ref: str) -> int:
    m = COL_RE.match(ref or '')
    if not m:
        return -1
    n = 0
    for letter in m.group(1).upper():
        n = n * 26 + ord(letter) - 64
    return n


def _read_strings(z: zipfile.ZipFile) -> list[str]:
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    out = []
    with z.open('xl/sharedStrings.xml') as stream:
        for _, elem in ET.iterparse(stream, events=('end',)):
            if elem.tag == MAIN('si'):
                out.append(''.join(t.text or '' for t in elem.iter(MAIN('t'))))
                elem.clear()
    return out


def _sheet_paths(z: zipfile.ZipFile) -> list[tuple[str, str]]:
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    relroot = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rels = {r.attrib['Id']: r.attrib['Target'] for r in relroot}
    out = []
    for sh in wb.find(MAIN('sheets')):
        rid = sh.attrib.get('{' + NS_REL + '}id')
        target = rels.get(rid, '')
        if target.startswith('/xl/'):
            path = target.lstrip('/')
        else:
            path = 'xl/' + target.removeprefix('/').removeprefix('xl/')
        if path not in z.namelist():
            raise ValueError(f'검사결과서 시트 경로를 읽지 못했습니다: {sh.attrib["name"]}')
        out.append((sh.attrib['name'], path))
    return out


def _iter_rows(z: zipfile.ZipFile, path: str, strings: list[str]):
    with z.open(path) as stream:
        for _, row in ET.iterparse(stream, events=('end',)):
            if row.tag != MAIN('row'):
                continue
            cells = {}
            for cell in row.findall(MAIN('c')):
                col = _column_number(cell.attrib.get('r', ''))
                if col < 1:
                    continue
                kind = cell.attrib.get('t')
                v = cell.find(MAIN('v'))
                text = v.text if v is not None and v.text is not None else ''
                if kind == 's' and text:
                    try:
                        text = strings[int(text)]
                    except (ValueError, IndexError):
                        text = ''
                elif kind == 'inlineStr':
                    inline = cell.find(MAIN('is'))
                    text = ''.join(x.text or '' for x in inline.iter(MAIN('t'))) if inline is not None else ''
                elif kind == 'b':
                    text = 'TRUE' if text == '1' else 'FALSE'
                cells[col] = str(text)
            yield int(row.attrib.get('r', '0')), cells
            row.clear()


def _project_id(path: Path) -> str:
    """Use the same ID resolution as the RuleKB (e.g. historical (8) climate => 007)."""
    import importlib.util
    extractor = Path(__file__).resolve().parent.parent / 'DataGuard_RuleKB_v01' / 'extract_rules.py'
    if extractor.is_file():
        spec = importlib.util.spec_from_file_location('dataguard_matcher_extractor', extractor)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        try:
            return module.identify_project_id(Path(path), {'projects': {}})
        except ValueError as exc:
            # Retain legacy filename fallback only when the name is unknown.
            if '식별하지 못했습니다' not in str(exc):
                raise
    m = re.match(r'^\s*\((\d{1,4})\)', path.name)
    return m.group(1).zfill(3) if m else '미확인'


def parse_inspection(inspection_file: str | Path, project_id: str | None = None) -> tuple[list[dict], dict]:
    """Extract actual detail rows (not aggregates) from numbered error sheets."""
    path = Path(inspection_file).expanduser()
    if not path.is_file():
        raise FileNotFoundError(f'검사결과서 파일이 없습니다: {path}')
    if path.suffix.lower() != '.xlsx' or not zipfile.is_zipfile(path):
        raise ValueError('정상적인 .xlsx 검사결과서를 선택해 주세요.')
    pid = str(project_id).zfill(3) if project_id else _project_id(path)
    results = []
    discovered = []
    with zipfile.ZipFile(path) as z:
        strings = _read_strings(z)
        for name, xmlpath in _sheet_paths(z):
            m = SHEET_RE.match(name)
            if not m:
                continue
            group, category = m.group(1), m.group(2)
            discovered.append(name)
            header = None
            for excel_row, cells in _iter_rows(z, xmlpath, strings):
                # Some Excel files have count summaries before detail headers.
                if header is None:
                    cols = {str(s).strip(): idx for idx, s in cells.items() if str(s).strip()}
                    if '파일명' in cols and ('검사 항목' in cols or '오류 유형' in cols):
                        header = cols
                    continue
                def col(*names):
                    for key in names:
                        idx = header.get(key)
                        if idx is not None:
                            return (cells.get(idx) or '').strip()
                    return ''
                raw_name = col('파일명')
                if not raw_name:
                    continue
                # skip accidental repeat headers inside a large sheet.
                if raw_name == '파일명':
                    continue
                parsed_name, index = normalize_filename(raw_name)
                results.append({
                    '사업ID': pid,
                    '규칙그룹': group,
                    '오류구분': category,
                    '검사결과서 시트': name,
                    '검사결과서 행': excel_row,
                    '파일명(원문)': raw_name,
                    '매칭파일명': parsed_name,
                    '원본 상대경로': parse_file_reference(raw_name)[0],
                    '행표기': index,
                    '검사 항목': col('검사 항목', '오류 유형'),
                    '검사 프로퍼티': col('검사 프로퍼티', '프로퍼티'),
                    '검사 규칙': col('검사 규칙'),
                    '오류 메시지': col('오류 메시지', '오류 사유'),
                })
    if not discovered:
        raise ValueError('검사결과서에서 「n. 구조 오류 목록 / 형식 오류 목록」 시트를 찾지 못했습니다.')
    return results, {'project_id': pid, 'sheets': discovered, 'issues': len(results)}


def parse_file_reference(raw: str) -> tuple[str, str]:
    """Extract relative path and row number; preserve inner parentheses in filename.

    Examples: 'moel/.../dialog.json (128)' -> ('moel/.../dialog.json','128')
              '제4장(2).jsonl (31)' -> ('제4장(2).jsonl','31')
    """
    value = unicodedata.normalize('NFKC', str(raw or '')).strip().strip('"\'')
    value = value.replace('\\', '/')
    m = ROW_SUFFIX_RE.fullmatch(value)
    if m:
        return m.group('name'), m.group('line')
    return value, ''


def normalize_filename(raw: str) -> tuple[str, str]:
    path, row_index = parse_file_reference(raw)
    return PurePosixPath(path).name, row_index


def _normalized(s: str) -> str:
    return unicodedata.normalize('NFKC', s).casefold()


def _safe_zip_name(name: str) -> bool:
    if not name or '\x00' in name or '\\' in name or name.startswith('/'):
        return False
    parts = PurePosixPath(name).parts
    return bool(parts) and '..' not in parts and not re.match(r'^[A-Za-z]:', parts[0])


def scan_folder(folder: str | Path) -> dict[str, list[str]]:
    root = Path(folder).expanduser()
    if not root.is_dir():
        raise NotADirectoryError(f'라벨링데이터 폴더를 찾지 못했습니다: {root}')
    db = defaultdict(list)
    # local server filesystem, not an upload: no copy, no file content load.
    for base, subdirs, files in os.walk(root, followlinks=False):
        subdirs[:] = [d for d in subdirs if not (Path(base) / d).is_symlink()]
        for file in files:
            p = Path(base) / file
            if p.is_symlink() or p.suffix.lower() not in ALLOWED_EXTENSIONS:
                continue
            relative = p.relative_to(root).as_posix()
            db[_normalized(file)].append(relative)
    return db


def scan_zip(zip_path: str | Path) -> dict[str, list[str]]:
    path = Path(zip_path).expanduser()
    if not path.is_file() or not zipfile.is_zipfile(path):
        raise ValueError('정상적인 ZIP 파일을 선택해 주세요.')
    db = defaultdict(list)
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        if len(infos) > MAX_ZIP_ENTRIES:
            raise ValueError('ZIP 내부 파일이 너무 많아 안전하게 열 수 없습니다.')
        total = 0
        for info in infos:
            if info.is_dir():
                continue
            if not _safe_zip_name(info.filename):
                raise ValueError(f'ZIP 내부 경로가 안전하지 않습니다: {info.filename[:80]}')
            if info.flag_bits & 1:
                raise ValueError('암호화된 ZIP 파일은 지원하지 않습니다.')
            if info.file_size > MAX_SINGLE_UNCOMPRESSED:
                raise ValueError('ZIP에 너무 큰 단일 파일이 있습니다.')
            if info.file_size > MAX_COMPRESSION_RATIO * max(info.compress_size, 1):
                raise ValueError('압축률이 비정상적으로 높은 ZIP 파일이 있습니다.')
            total += info.file_size
            if total > MAX_TOTAL_UNCOMPRESSED:
                raise ValueError('ZIP 압축 해제 예상 크기가 허용 한도를 초과합니다.')
            p = PurePosixPath(info.filename)
            if p.suffix.lower() in ALLOWED_EXTENSIONS:
                db[_normalized(p.name)].append(info.filename)
    return db


def match_inspection(inspection_path: str | Path, *, folder: str | Path | None = None,
                     archive: str | Path | None = None, project_id: str | None = None,
                     group_extensions: dict[str, list[str]] | None = None) -> tuple[list[dict], dict]:
    if bool(folder) == bool(archive):
        raise ValueError('로컬 폴더 또는 ZIP 중 하나만 선택해 주세요.')
    rows, meta = parse_inspection(inspection_path, project_id)
    index = scan_folder(folder) if folder else scan_zip(archive)
    extmap = group_extensions or ({'1': ['.json'], '2': ['.txt']}
                                   if meta['project_id'] == '019' else {})
    matched_unique = set()
    counts = Counter()
    for row in rows:
        name = row['매칭파일명']
        ext = Path(name).suffix.casefold()
        permitted = [e.casefold() for e in extmap.get(row['규칙그룹'], [])]
        if not name or ext not in ALLOWED_EXTENSIONS:
            status, candidates = '파일명 해석 실패', []
        elif permitted and ext not in permitted:
            status, candidates = '그룹 확장자 불일치', []
        else:
            candidates = index.get(_normalized(name), [])
            # Explicit relative folders in the inspection are authoritative.
            # Do not accidentally match a same-named file from a different group.
            request = row.get('원본 상대경로', '').strip('/').replace('\\', '/')
            if '/' in request:
                key = _normalized(request)
                matching_paths = [c for c in candidates if _normalized(c) == key or
                                  _normalized(c).endswith('/' + key)]
                status = '경로 불일치' if candidates and not matching_paths else None
                candidates = matching_paths
            else:
                status = None
            if status == '경로 불일치':
                pass
            elif len(candidates) == 1:
                status = '일치'
                matched_unique.add(candidates[0])
            elif len(candidates) > 1:
                status = '동일 파일명 중복'
            else:
                status = '파일 없음'
        row['매칭상태'] = status
        row['실제 데이터 위치'] = candidates[0] if len(candidates) == 1 else ''
        row['동일명 후보수'] = len(candidates)
        row['후보 경로'] = ' | '.join(candidates[:5])  # at most five; no arbitrary selection
        counts[status] += 1
    meta.update({
        'source_type': 'folder' if folder else 'zip',
        'source': str(folder or archive),
        'file_count': sum(map(len, index.values())),
        'unique_names': len(index),
        'unique_matched_files': len(matched_unique),
        'status_counts': dict(counts),
        'read_only': True,
    })
    return rows, meta


def save_results(rows: list[dict], meta: dict, dest: str | Path) -> tuple[Path, Path]:
    output_dir = Path(dest)
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + uuid.uuid4().hex[:6]
    stem = f"{meta['project_id']}_file_match_{stamp}"
    csv_path, json_path = output_dir / (stem + '.csv'), output_dir / (stem + '_summary.json')
    with csv_path.open('w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=CSV_COLUMNS, extrasaction='ignore')
        w.writeheader()
        w.writerows(rows)
    with json_path.open('w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    return csv_path, json_path


def run(inspection_path, source_path, mode='folder', output_dir='outputs/step4_matching',
        project_id=None):
    rows, info = match_inspection(inspection_path,
        folder=source_path if mode == 'folder' else None,
        archive=source_path if mode == 'zip' else None,
        project_id=project_id,
    )
    csv_file, json_file = save_results(rows, info, output_dir)
    return {'rows': rows, 'summary': info, 'csv': str(csv_file), 'json': str(json_file)}

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='DataGuard · 검사결과서 오류 파일 매칭 (원본 변경 없음)')
    parser.add_argument('--inspection', required=True, help='검사결과서 .xlsx 경로')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--folder', help='로컬 라벨링데이터 폴더')
    group.add_argument('--zip', dest='zip_path', help='라벨링데이터 ZIP 경로')
    parser.add_argument('--out', default='outputs/file_matching', help='매칭 결과 CSV/JSON 저장 폴더')
    args = parser.parse_args()
    result = run(args.inspection, args.folder or args.zip_path,
                 mode='folder' if args.folder else 'zip', output_dir=args.out)
    print(json.dumps(result['summary'], ensure_ascii=False, indent=2))
    print('CSV:', result['csv'])
    print('요약:', result['json'])
