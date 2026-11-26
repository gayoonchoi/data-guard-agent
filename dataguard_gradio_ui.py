"""DataGuard Agent — 누적 구축용 Rule KB UI v0.5.3.

화면은 ① 검사결과서 등록, ② 불일치 규칙 승인 두 탭만 제공한다.
파일/시트 미리보기·시연용 오류 및 결과서 UI를 제거했다.
규칙 등록은 기존 extract_rules.run(), 승인은 별도 오버레이에만 저장한다.
"""
from __future__ import annotations

import importlib.util
import csv
import hashlib
import json
import os
from pathlib import Path
import re
import threading
import sys

import gradio as gr
import pandas as pd

KB_ROOT = Path(os.environ.get('DATAGUARD_RULE_KB_DIR',
    str(Path(__file__).resolve().parent / 'DataGuard_RuleKB_v01'))).resolve()
USER_ROOT = KB_ROOT.parent / 'DataGuard_RuleKB_UserBuilt'
WORKSPACES = [('새로 구축하는 사전 (누적)', 'new'), ('기존 사전 (참고·갱신)', 'existing')]


def data_root(workspace: str) -> Path:
    return KB_ROOT if workspace == 'existing' else USER_ROOT
_LOCK = threading.RLock()
VIEW_COLS = ['사업ID', '그룹', '항목 경로', '규칙 셀', '설정', '구축기업 규칙', '검사 규칙', '처리 상태']


def _load_backend(module_name: str):
    path = KB_ROOT / f'{module_name}.py'
    if not path.exists():
        raise FileNotFoundError(f'{module_name}.py를 찾지 못했습니다: {path}')
    spec = importlib.util.spec_from_file_location(f'dg_ui_{module_name}', path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _txt(value):
    if value is None:
        return '(빈 셀)'
    return str(value)


def _view_conflicts(workspace='new'):
    approval = _load_backend('rule_approval')
    return approval.list_conflicts(data_root(workspace))


def _registration_warnings(catalog: dict, extractor) -> list[str]:
    """Flag suspicious entries created by earlier versions, without deleting data."""
    projects = catalog.get("projects", {})
    problems = []
    fingerprints = {}
    for pid, data in sorted(projects.items()):
        basename = data.get("source_workbook", "")
        if basename:
            try:
                true_id = extractor.identify_project_id(Path(basename), {"projects": {}})
                if true_id != pid:
                    problems.append(f"{pid} (실제 파일명은 {true_id}번)")
            except ValueError:
                # Unknown filename alone isn't proof of a wrong registration.
                pass
        digest = data.get("source_sha256")
        if digest:
            fingerprints.setdefault(digest, []).append(pid)
    for ids in fingerprints.values():
        if len(ids) > 1:
            problems.append("동일한 원본으로 등록된 사업: " + ", ".join(sorted(ids)))
    return problems


def _summary(rows, workspace):
    catalog = data_root(workspace) / 'catalog.json'
    audit = data_root(workspace) / 'audit.json'
    if not catalog.is_file() or not audit.is_file():
        return '**등록된 사업 0개** · 검사결과서 Excel을 등록해 주세요.'
    c = json.loads(catalog.read_text(encoding='utf-8'))
    a = json.loads(audit.read_text(encoding='utf-8'))
    active = sum(r['status'].startswith('승인됨') for r in rows)
    pending = len(rows) - active
    info = (f"**사업 {len(c.get('projects', {}))}개 · 그룹 {a.get('groups', 0)}개 · "
            f"규칙 항목 {a.get('rule_rows', 0)}개**  \n"
            f"불일치 {len(rows)}건 (승인 {active} · 검토 대기 {pending})")
    issues = _registration_warnings(c, _load_backend('extract_rules'))
    if issues:
        info += ("  \n**⚠️ 이전에 잘못 연결된 사업이 있어요:** " + "; ".join(issues) +
                 "  \n해당 사업의 **실제 검사결과서**를 업로드하면 교체할 수 있어요.")
    return info


def _table(rows):
    return pd.DataFrame([[
        r['project_id'], r['group_id'], r['path'] or '(연결되지 않음)',
        r['cell'], r['setting'], _txt(r['company']),
        _txt(r['inspection']), r['status']
    ] for r in rows], columns=VIEW_COLS)


def _choices(rows):
    return [f"{r['key']} | {r['path'] or r['setting']}" for r in rows]


def _key(label):
    return (label or '').split(' | ', 1)[0]


def refresh(workspace='new', selected=None):
    with _LOCK:
        rows = _view_conflicts(workspace)
        choices = _choices(rows)
        picked = selected if selected in choices else (choices[0] if choices else None)
        return (_summary(rows, workspace), _table(rows),
                gr.update(choices=choices, value=picked))


def on_uploaded_files(files, workspace='new'):
    """Detect each uploaded workbook ID; show manual input ONLY when unknown.

    This is a UI convenience only. Registration independently validates source identity.
    """
    paths = [Path(x) for x in (files if isinstance(files, list) else ([files] if files else []))]
    if not paths:
        return gr.update(value='', visible=False), ''
    extractor = _load_backend('extract_rules')
    catalog = data_root(workspace) / 'catalog.json'
    existing = (json.loads(catalog.read_text(encoding='utf-8'))
                if catalog.is_file() else {'projects': {}})
    recognized, unknown = [], []
    for p in paths:
        try:
            pid = extractor.identify_project_id(p, existing)
            recognized.append(f'{pid} ({p.name})')
        except ValueError:
            unknown.append(p.name)
    if unknown:
        if len(paths) == 1:
            return (gr.update(value='', visible=True),
                    '사업 ID 자동 인식 실패 · 아래에 사업 번호(예: 99)를 직접 입력해 주세요.')
        return (gr.update(value='', visible=False),
                '일부 파일의 사업 ID를 인식할 수 없습니다. 파일명 앞에 (099)처럼 번호를 붙여서 다시 업로드해 주세요: '
                + ', '.join(unknown))
    return gr.update(value='', visible=False), '인식한 사업: ' + ', '.join(x.split(' (')[0] for x in recognized)


def register_workbooks(files, workspace='new', project_id_override=''):
    paths = [Path(x) for x in (files if isinstance(files, list) else ([files] if files else []))]
    if not paths:
        return (*refresh(workspace), '검사결과서 Excel을 선택해 주세요.', gr.update(), gr.update())
    invalid = [p.name for p in paths if not p.is_file() or p.suffix.lower() != '.xlsx'
               or '검사결과서' not in p.name]
    if invalid:
        return (*refresh(workspace), '검사결과서(.xlsx)만 등록할 수 있습니다: ' + ', '.join(invalid),
                gr.update(), gr.update())
    try:
        with _LOCK:
            root = data_root(workspace)
            catalog_file = root / 'catalog.json'
            old_projects = (json.loads(catalog_file.read_text(encoding='utf8')).get('projects', {})
                            if catalog_file.is_file() else {})
            extractor = _load_backend('extract_rules')
            manual = (project_id_override or '').strip() or None
            # An automatically identifiable file ALWAYS takes priority over a stale
            # textbox value from the previous upload. Manual ID is fallback only.
            detected = []
            for p in paths:
                try:
                    pid = extractor.identify_project_id(p, {'projects': old_projects})
                    detected.append((p, pid, False))
                except ValueError as exc:
                    if '식별하지 못했습니다' not in str(exc):
                        raise  # A contradictory filename is NOT a manual-ID case.
                    detected.append((p, None, True))
            if manual and len(paths) > 1:
                raise ValueError('사업 ID 직접 입력은 Excel 파일 1개일 때만 사용할 수 있습니다.')
            if len(paths) > 1 and any(unknown for _, _, unknown in detected):
                raise ValueError('사업 ID가 없는 파일은 먼저 (099)처럼 번호를 붙여 다시 올려주세요.')
            resolved = [(p, pid if pid is not None else extractor.identify_project_id(
                            p, {'projects': old_projects}, manual))
                        for p, pid, _ in detected]
            # Pass manual ID to the extractor only when it actually was needed;
            # for known files, never reinterpret or overwrite the true ID.
            override_for_extractor = (manual if any(x[2] for x in detected) else None)
            ids = [pid for _, pid in resolved]
            if len(ids) != len(set(ids)):
                raise ValueError('동일 사업의 검사결과서를 한 번에 두 개 등록할 수 없습니다.')
            unchanged, pending = [], []
            for p, pid in resolved:
                digest = hashlib.sha256(p.read_bytes()).hexdigest()
                if old_projects.get(pid, {}).get('source_sha256') == digest:
                    unchanged.append(pid)
                else:
                    # Cross-business deduplication, but tolerate legacy bad entries
                    # when the genuine project is about to replace its wrong data.
                    for other_pid, old in old_projects.items():
                        if other_pid != pid and old.get('source_sha256') == digest:
                            raise ValueError(
                                f'{pid}번 파일이 이미 {other_pid}번으로 등록되어 있습니다. '
                                '실제 검사결과서를 확인해 주세요.'
                            )
                    pending.append((p, pid))
            if pending:
                _, audit = extractor.run([p for p, _ in pending], root, merge=True,
                                         project_id_override=override_for_extractor)
            else:
                audit = json.loads((root / 'audit.json').read_text(encoding='utf-8'))
            added = [pid for _, pid in pending if pid not in old_projects]
            updated = [pid for _, pid in pending if pid in old_projects]
            state = refresh(workspace)
            group_infos = []
            saved = json.loads((root / 'catalog.json').read_text(encoding='utf8'))['projects']
            for _, pid in resolved:
                proj = saved[pid]
                n_groups = len(proj['groups'])
                n_rules = sum(g['rule_count'] for g in proj['groups'].values())
                group_infos.append(f"{pid} {proj['name']} ({n_groups}그룹 · {n_rules}항목)")
        parts = []
        if added:
            parts.append('신규 등록 ' + ', '.join(added))
        if updated:
            parts.append('기존 사업 갱신 ' + ', '.join(updated))
        if unchanged:
            parts.append('동일 파일 건너뜀 ' + ', '.join(unchanged))
        return (*state, '✅ ' + ' · '.join(parts) + '  \n' + '  \n'.join(group_infos),
                gr.update(value=None), gr.update(value='', visible=False))
    except Exception as exc:
        return (*refresh(workspace), f'등록 실패: {type(exc).__name__}: {exc}',
                gr.update(), gr.update())


def conflict_detail(choice, workspace):
    key = _key(choice)
    row = next((r for r in _view_conflicts(workspace) if r['key'] == key), None)
    if row is None:
        return '', '', '', '확인할 규칙을 선택해 주세요.'
    info = (f"**{row['project_id']}번 / 그룹 {row['group_id']} / {row['path'] or '항목 미연결'}**  \n"
            f"설정: `{row['setting']}` · 원본 셀: `{row['cell']}` · **{row['status']}**")
    if not row['approvable']:
        info += ('  \n이 항목은 현재 승인 후 자동 검사 활성화를 지원하지 않아 '\
                 '**수동 확인 대상**이야.')
    else:
        info += '  \n승인 전 두 원본 규칙을 확인하고, 적용할 출처와 근거를 기록해.'
    if row['setting'] == 'pattern' and row['inspection'] == r'^\d{5}.jpg$':
        info += ('  \n⚠️ 검사 규칙 `^\\d{5}.jpg$`에서 확장자 앞의 `.`는 임의의 한 문자와도 일치해. '
                 '정확한 점을 의미하려면 원본 규칙에서 `\\.jpg`로 정정해야 해.')
    return _txt(row['company']), _txt(row['inspection']), row['status'], info


def change_decision(choice, source, reviewer, reason, acknowledged, action, workspace):
    key = _key(choice)
    if not key:
        return (*refresh(workspace, choice), '먼저 규칙을 선택해 주세요.')
    if not acknowledged:
        return (*refresh(workspace, choice), '원본 두 값을 확인했는지 체크해 주세요.')
    try:
        with _LOCK:
            backend = _load_backend('rule_approval')
            if action == 'approve':
                rec = backend.approve_conflict(data_root(workspace), key, source, reviewer, reason)
                message = (f"승인 기록 완료 · {key} · "
                           f"{'검사 규칙' if source == 'inspection' else '구축기업 규칙'} 선택. "
                           '이 값은 읽기 전용 검증기에서만 적용되고 원본 JSON은 그대로 유지돼.')
            else:
                backend.revoke_conflict(data_root(workspace), key, reviewer, reason)
                message = f'승인 철회 완료 · {key} · 해당 검사는 다시 보류돼.'
            state = refresh(workspace, choice)
        return (*state, message)
    except Exception as exc:
        return (*refresh(workspace, choice), f'처리 실패: {type(exc).__name__}: {exc}')



def create_report(inspection_file, round_number):
    """③ Create the inspection-only draft. Never read or apply ④ evidence here."""
    if not inspection_file:
        return '검사결과서 Excel을 선택해 주세요.', None
    try:
        root = Path(__file__).resolve().parent
        module_path = root / '5. ExcelReport' / 'report_generator.py'
        spec = importlib.util.spec_from_file_location('dataguard_report_generator', module_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        result = module.generate_report(inspection_file, round_no=int(round_number or 1))
        status = (f"**초안 생성 완료 — {result['project_id']}번**  \n"
                  f"규칙 {result['groups']}개 그룹 · 오류 목록 {result['issues']:,}건 · "
                  "보완 완료 0건  \n"
                  "검사결과서 기준 초안이며 ④ 실제 데이터 검토 내역은 반영하지 않았습니다.")
        status = (f"**③ 생성기: {result.get('generator_version', '확인 불가')}**  \n" + status)
        if result.get('warnings'):
            status += '\n**검사결과서 집계 불일치(검토 필요):** ' + ' / '.join(result['warnings'][:8])
        return status, result['path']
    except Exception as exc:
        return f'초안 생성 실패: {type(exc).__name__}: {exc}', None


def apply_step5_review(inspection_file, draft_report, evidence_csv_file):
    """⑤ Apply ④ review CSV to an existing ③ draft, as a NEW Excel file."""
    if not inspection_file or not draft_report or not evidence_csv_file:
        return '검사결과서·③ 초안 Excel·④ 검토 CSV를 모두 선택해 주세요.', None
    try:
        # Separate output folders retain the user's original report filename without
        # silently overwriting previous runs or the inspection-only ③ output.
        from datetime import datetime
        from uuid import uuid4
        root = Path(__file__).resolve().parent
        bridge_path = root / '5. ExcelReport' / 'evidence_report_bridge.py'
        spec = importlib.util.spec_from_file_location('dg_step5_evidence_bridge', bridge_path)
        bridge = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(bridge)
        out_dir = (root / '5. ExcelReport' / 'outputs' / 'evidence_applied'
                   / (datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + uuid4().hex[:8]))
        result = bridge.apply_evidence(
            report_path=draft_report,
            inspection_path=inspection_file,
            evidence_csv=evidence_csv_file,
            output_dir=out_dir,
        )
        stats = result['statistics']
        msg = (f"**⑤ 반영 완료 — 사업 {result['project_id']}**  \n"
               f"실제값 확인 {stats.get('실제값 확인',0):,}건 · "
               f"검토 대기 {stats.get('검토 대기',0):,}건 · "
               f"삭제·제외 간주 {stats.get('삭제·제외 간주',0):,}건 · "
               f"대조 완료 {stats.get('대조 완료',0):,}건 · "
               f"Ollama 제안 반영 {stats.get('로컬LLM 제안 반영',0):,}건  \n"
               "③ 초안·구문규칙·원본 라벨링데이터는 변경하지 않았습니다. "
               "수정 지시 초안이며 실제 보완 완료 건수는 0건입니다.")
        return msg, result['path']
    except Exception as exc:
        return f'⑤ 반영 실패: {type(exc).__name__}: {exc}', None


def load_step5_inputs(report_inspection, report_output, matcher_inspection, review_csv, ollama_csv=None):
    """Convenience: copy session files into ⑤ controls, without processing them."""
    inspection_file = report_inspection or matcher_inspection
    n_proposals = _model_csv_count(ollama_csv) if ollama_csv else None
    llm_note = ('⚠️ Ollama 수정안 0건: ④ 기본 검토 CSV를 사용합니다. '
                '원본 파일 매칭 실패 원인을 먼저 확인하세요.'
                if n_proposals == 0 else
                (f'Ollama 수정 제안 {n_proposals}건 포함' if n_proposals is not None else ''))
    review_csv = (ollama_csv if n_proposals else review_csv)
    if not inspection_file or not report_output or not review_csv:
        missing = []
        if not inspection_file: missing.append('검사결과서')
        if not report_output: missing.append('③ 보완결과서 초안')
        if not review_csv: missing.append('④ 검토 CSV')
        return (gr.update(value=inspection_file), gr.update(value=report_output),
                gr.update(value=review_csv),
                '불러올 파일이 없습니다: ' + ', '.join(missing) + ' · 아래에서 파일을 직접 선택할 수 있어요.')
    return (gr.update(value=inspection_file), gr.update(value=report_output),
            gr.update(value=review_csv), '③ 초안과 ④ CSV를 불러왔어. 사업 ID가 같은지 확인 후 ⑤를 실행해 주세요. ' + llm_note)


def choose_local_folder(current_path):
    """Open the Windows native folder picker on this SAME PC as the Gradio server."""
    if sys.platform != 'win32':
        return gr.update(), '폴더 선택 팝업은 Windows 로컬 실행에서 지원합니다. 경로 직접 입력은 가능합니다.'
    try:
        import tkinter as tk
        from tkinter import filedialog
        parent = tk.Tk()
        parent.withdraw()
        parent.attributes('-topmost', True)
        parent.update()
        start = str(current_path or '').strip().strip('"')
        if not Path(start).is_dir():
            start = str(Path.home())
        try:
            selected = filedialog.askdirectory(parent=parent, title='DataGuard 라벨링데이터 폴더 선택',
                                               initialdir=start, mustexist=True)
        finally:
            parent.destroy()
        if selected:
            return gr.update(value=selected), f'선택된 로컬 폴더: `{selected}`'
        return gr.update(), '폴더 선택을 취소했습니다.'
    except Exception as exc:
        return gr.update(), (f'Windows 폴더 선택 팝업을 열지 못했습니다: {type(exc).__name__}. '
                             '프로그램이 사용자 PC의 Windows 데스크톱에서 실행되는지 확인하세요.')


def _model_csv_count(path):
    if not path or not Path(path).is_file():
        return None
    with open(path, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        if '로컬LLM 결과' not in (reader.fieldnames or []):
            return None
        return sum(row.get('로컬LLM 결과', '').startswith('제안 생성') for row in reader)


def match_error_files(inspection_file, source_mode, folder_path, zip_file):
    """④ 로컬/ZIP 파일 매칭. 검사결과서와 라벨링 원본을 읽기 전용으로 취급."""
    empty = pd.DataFrame(columns=['규칙그룹', '파일명(원문)', '검사 항목', '매칭상태', '실제 데이터 위치'])
    if not inspection_file:
        return '검사결과서 Excel을 올려 주세요.', empty, None, None
    try:
        root = Path(__file__).resolve().parent
        module_path = root / '1. RuleEngine' / 'data_loader_matcher.py'
        spec = importlib.util.spec_from_file_location('dataguard_data_loader_matcher', module_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        if source_mode == 'folder':
            if not (folder_path or '').strip():
                raise ValueError('로컬 라벨링데이터 폴더 경로를 입력해 주세요.')
            src = folder_path.strip().strip('"')
        else:
            if not zip_file:
                raise ValueError('ZIP 파일을 올려 주세요.')
            src = zip_file
        result = module.run(
            inspection_path=inspection_file,
            source_path=src,
            mode=source_mode,
            output_dir=root / '1. RuleEngine' / 'outputs' / 'file_matching',
        )
        info = result['summary']
        counter = info['status_counts']
        summary = (f"**사업 {info['project_id']} · 지적 오류 {info['issues']:,}건 · "
                   f"실제 파일 {info['file_count']:,}개**  \n"
                   f"일치 {counter.get('일치', 0):,}건 · 파일 없음 {counter.get('파일 없음', 0):,}건 · "
                   f"동일 파일명 중복 {counter.get('동일 파일명 중복', 0):,}건 · "
                   f"기타 {sum(v for k,v in counter.items() if k not in ('일치','파일 없음','동일 파일명 중복')):,}건  \n"
                   f"**서로 다른 실제 파일 매칭:** {info['unique_matched_files']:,}개 · "
                   "원본 데이터 변경 없음")
        preview_cols = ['규칙그룹','오류구분','검사결과서 행','파일명(원문)',
                        '검사 항목','매칭상태','실제 데이터 위치']
        show = pd.DataFrame(result['rows'][:150], columns=preview_cols)
        return summary, show, result['csv'], result['json']
    except Exception as exc:
        return f'파일 연결 실패: {type(exc).__name__}: {exc}', empty, None, None



def inspect_before_after(inspection_file, source_mode, folder_path, zip_file,
                         final_dataset_missing_is_removed=False):
    """④ actual field values + rule-based correction instructions (NEVER mark completed)."""
    cols = ['파일명(원문)', '행표기', '검사 프로퍼티', '현재 검증 상태',
            '보완 전(실제값/미리보기)', '보완 후(규칙상 요구사항/초안)', '담당자 조치']
    empty = pd.DataFrame(columns=cols)
    if not inspection_file:
        return '검사결과서를 업로드해 주세요.', empty, None, None
    try:
        root = Path(__file__).resolve().parent
        module_path = root / '1. RuleEngine' / 'rule_evidence_inspector.py'
        sys.path.insert(0, str(module_path.parent))
        spec = importlib.util.spec_from_file_location('dataguard_evidence_inspector', module_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        source = (folder_path or '').strip().strip('"') if source_mode == 'folder' else zip_file
        if not source:
            raise ValueError('로컬 폴더 경로 또는 ZIP 파일을 지정해 주세요.')
        result = module.analyze(
            inspection=inspection_file, source=source, mode=source_mode,
            output_dir=root / '1. RuleEngine' / 'outputs' / 'evidence_review',
            treat_absent_as_removed=bool(final_dataset_missing_is_removed),
        )
        a = result['summary']
        counts = a['evidence_status_counts']
        warning = ''
        match = re.search(r'\((\d{1,4})\)', str(inspection_file).split('/')[-1])
        if match and source_mode == 'folder':
            pid = match.group(1).zfill(3)
            for other in re.findall(r'(?<!\d)0*(\d{2,4})(?!\d)\s*[.。]', str(source)):
                if other.zfill(3) != pid:
                    warning = f'  \n⚠️ 검사결과서 {pid}와 폴더의 사업 번호가 다른 것 같아 경로를 확인해 주세요.'
                    break
        message = (f"**사업 {a['project_id']} · 오류 {a['issues']:,}건 검토**  \n"
                   f"실제 규칙 위반 {counts.get('현재 규칙 위반',0):,}건 · "
                   f"현재 규칙 통과 {counts.get('현재 규칙 통과',0):,}건 · "
                   f"삭제·제외 간주 {counts.get('삭제·제외 간주(최종 폴더 부재)',0):,}건 · "
                   f"추가 확인 필요 {sum(v for k,v in counts.items() if k not in ('현재 규칙 위반','현재 규칙 통과','삭제·제외 간주(최종 폴더 부재)')):,}건  \n"
                   f"**원본 불변 · 보완 완료 0건** · 보완 후 칸은 수정 지시 초안" + warning)
        return message, pd.DataFrame(result['rows'][:40], columns=cols), result['csv'], result['json']
    except Exception as exc:
        return f'보완 전후 검토 실패: {type(exc).__name__}: {exc}', empty, None, None

def write_local_correction_suggestions(review_csv_path, model_name, limit, include_before):
    """Optional local-only ④-B: preserve ④ CSV and create a new ⑤-compatible review CSV."""
    empty = pd.DataFrame(columns=["검사 프로퍼티", "보완 전(실제값/미리보기)",
                                  "보완 후(규칙상 요구사항/초안)", "로컬LLM 결과"])
    if not review_csv_path:
        return "먼저 ④ 검토 CSV를 선택해 주세요.", empty, None
    try:
        root = Path(__file__).resolve().parent
        module_path = root / '1. RuleEngine' / 'local_correction_agent.py'
        if not module_path.is_file():
            raise FileNotFoundError(f'로컬 수정안 모듈 누락: {module_path}')
        spec = importlib.util.spec_from_file_location('dg_local_correction_agent', module_path)
        agent = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(agent)
        model = str(model_name or 'qwen2.5:7b-instruct').strip()
        count = int(limit or 10)
        if not 1 <= count <= 200:
            raise ValueError('한 번에 1~200건까지 실험할 수 있습니다.')
        from datetime import datetime
        from uuid import uuid4
        destination = (root / '1. RuleEngine' / 'outputs' / 'ollama_review'
                       / (datetime.now().strftime('%Y%m%d_%H%M%S') + '_'
                          + uuid4().hex[:8] + '.csv'))
        result = agent.enhance_csv(
            source=review_csv_path, destination=destination,
            model=model, limit=count, include_before=bool(include_before),
        )
        preview = pd.read_csv(result['output'], encoding='utf-8-sig',
                              dtype=str, keep_default_na=False, nrows=30)
        show_cols = [x for x in ('파일명(원문)', '검사 프로퍼티',
                                 '보완 전(실제값/미리보기)',
                                 '보완 후(규칙상 요구사항/초안)',
                                 '로컬LLM 결과', '로컬LLM 제안 근거')
                     if x in preview.columns]
        preview = preview[show_cols]
        info = (f"**로컬 수정 제안 처리 완료** · 모델 `{model}`  \n"
                f"제안 생성 {result['generated']:,}건 · "
                f"실패 {result['failures']:,}건 · 건너뜀 {result['skipped']:,}건 · "
                f"모델 추론 호출 시도 {result['inference_attempts']:,}회  \n")
        if result['inference_attempts'] == 0:
            reasons = ', '.join(f'{k} {v}건' for k, v in result['skip_reasons'].items())
            info += f"⚠️ **이번 실행에서 모델 추론 0회** — {reasons or '대상 행 없음'}  \n"
        info += ("실제 데이터와 ④ CSV는 수정하지 않았습니다. "
                 "결과 CSV를 검토한 뒤 ⑤에서 반영해 주세요.")
        return info, preview, result['output']
    except Exception as exc:
        return f"로컬 수정안 생성 실패: {type(exc).__name__}: {exc}", empty, None


CSS = """
.gradio-container {max-width: 1050px !important; margin: auto !important;}
#title {margin-bottom: 3px;}
footer {display: none !important;}
"""


def build_app():
    with gr.Blocks(title='DataGuard Agent | 규칙 사전') as demo:
        gr.Markdown('# DataGuard Agent', elem_id='title')
        gr.Markdown('검사결과서 등록 · 규칙 승인')
        workspace = gr.Radio(choices=WORKSPACES, value='new',
                             label='작업할 지식 사전')
        with gr.Tab('① 검사결과서 등록'):
            gr.Markdown('검사결과서(.xlsx)를 올리고 **사업 추가**를 누르세요.')
            uploads = gr.File(label='검사결과서 업로드', file_count='multiple',
                              file_types=['.xlsx'], type='filepath')
            identified = gr.Markdown('')
            override = gr.Textbox(label='사업 ID (자동 인식되지 않을 때만 입력)',
                                  placeholder='예: 99', max_lines=1, visible=False)
            with gr.Row():
                add_button = gr.Button('사업 추가 / 규칙 갱신', variant='primary')
                refresh_button = gr.Button('현재 사전 새로고침')
            register_msg = gr.Markdown('')
            summary = gr.Markdown('현재 사전 상태를 불러오는 중')
            conflicts = gr.Dataframe(label='구축기업 ↔ 검사 규칙 불일치',
                value=pd.DataFrame(columns=VIEW_COLS), interactive=False, wrap=True)

        with gr.Tab('② 규칙 승인'):
            gr.Markdown('불일치 규칙만 검토해. **검사 규칙과 구축기업 규칙 중 하나를 선택**하면 '
                        '원본과 별도의 승인 이력으로 저장돼.  '
                        '현재 자동 활성화 지원 범위는 **정규식(pattern) 충돌**이야.')
            selector = gr.Dropdown(label='불일치 규칙 선택', choices=[], interactive=True)
            info = gr.Markdown('불일치 규칙을 선택해 주세요.')
            with gr.Row():
                company = gr.Textbox(label='구축기업 구문규칙', interactive=False, lines=2)
                inspection = gr.Textbox(label='검사 구문규칙', interactive=False, lines=2)
            status = gr.Textbox(label='현재 승인 상태', interactive=False)
            source = gr.Radio(choices=[('검사 구문규칙', 'inspection'),
                                       ('구축기업 구문규칙', 'company')],
                              value='inspection', label='승인할 출처')
            reviewer = gr.Textbox(label='검토자', placeholder='승인 이력에 기록할 이름')
            reason = gr.Textbox(label='결정 근거 / 철회 사유', lines=2,
                                placeholder='예: 원본 데이터 파일명은 숫자 5자리 + .jpg 형식임을 확인')
            ack = gr.Checkbox(label='원본 검사결과서의 두 규칙을 직접 확인했습니다.')
            with gr.Row():
                approve_button = gr.Button('선택한 규칙 승인', variant='primary')
                revoke_button = gr.Button('기존 승인 철회')
            outcome = gr.Markdown('')
            gr.Markdown('**주의:** 승인 화면은 로컬 PoC이며 사용자 인증·권한 관리 기능은 없어. '
                        '승인 이력은 `rule_approvals.json`, `rule_approval_history.jsonl`에 기록돼. '
                        '새 검사결과서로 규칙 원문이 달라지면 과거 승인값은 적용하지 않아.')

        with gr.Tab('③ 보완결과서 생성'):
            report_input = gr.File(label='검사결과서 Excel', file_count='single',
                                   file_types=['.xlsx'], type='filepath')
            report_round = gr.Number(label='보완 차수', value=1, precision=0, minimum=1, maximum=99)
            gr.Markdown('검사결과서만으로 보완결과서 초안을 생성합니다. ④ 검토 결과는 ⑤에서 별도로 반영합니다.')
            make_report = gr.Button('보완결과서 초안 생성', variant='primary')
            report_notice = gr.Markdown('')
            report_download = gr.File(label='보완결과서 다운로드', interactive=False)
            gr.Markdown('원본 라벨링데이터가 없는 경우 실제 보완 완료 건수는 0건으로 작성됩니다.')

        with gr.Tab('④ 실제 데이터 연결'):
            gr.Markdown('검사결과서에 지적된 파일을 로컬 폴더 또는 ZIP에서 찾아 연결합니다. **원본은 수정하지 않습니다.**')
            matcher_inspection = gr.File(label='검사결과서 Excel', file_count='single',
                                        file_types=['.xlsx'], type='filepath')
            matcher_mode = gr.Radio(choices=[('로컬 폴더', 'folder'), ('ZIP 파일', 'zip')],
                                    value='folder', label='라벨링데이터 위치')
            with gr.Row():
                matcher_folder = gr.Textbox(label='로컬 폴더 경로 (자동 입력)',
                    placeholder=r'C:\...\2.라벨링데이터', scale=5,
                    info='버튼으로 폴더 선택. Windows 로컬 PC에서 실행 중일 때 팝업이 열립니다.')
                folder_browse = gr.Button('📁 폴더 선택...', variant='secondary', scale=1)
            browse_notice = gr.Markdown('')
            matcher_zip = gr.File(label='라벨링데이터 ZIP', file_count='single',
                                  file_types=['.zip'], type='filepath', visible=False)
            matcher_button = gr.Button('오류 파일 찾기', variant='primary')
            matcher_notice = gr.Markdown('')
            matcher_preview = gr.Dataframe(label='매칭 결과 미리보기 (처음 150건)',
                                           interactive=False, wrap=True)
            with gr.Row():
                matcher_csv = gr.File(label='전체 매칭 결과 CSV', interactive=False)
                matcher_json = gr.File(label='매칭 요약 JSON', interactive=False)
            final_dataset_removed = gr.Checkbox(
                label='선택한 폴더/ZIP이 최종 보완 데이터셋입니다. 여기에서 없는 오류 파일은 삭제·제외로 간주',
                value=False,
                info='최종 보완 데이터셋을 지정한 경우에만 선택하세요. 잘못된 폴더를 선택하면 삭제로 잘못 판정할 수 있습니다.',
            )
            gr.Markdown('**보완 전후 검토**: 실제 파일을 확인한 경우에만 현재값을 기록합니다. '
                        '최종 데이터셋의 미존재 파일은 체크한 경우 삭제·제외로 간주하고, 실제 삭제 이력 확인과는 구분합니다.')
            inspect_button = gr.Button('보완 전후 검토 초안 생성', variant='secondary')
            evidence_notice = gr.Markdown('')
            evidence_preview = gr.Dataframe(label='보완 전후 검토 (처음 40건)', interactive=False, wrap=True)
            with gr.Row():
                evidence_csv = gr.File(label='보완 전후 전체 CSV', interactive=False)
                evidence_json = gr.File(label='검토 요약 JSON', interactive=False)


            gr.Markdown('### 선택 기능: 로컬 Correction Agent (Qwen2.5 7B)')
            gr.Markdown('④의 보완 전후 CSV를 Qwen2.5 모델로 검토하여 한국어 보완 후 수정안을 제안합니다. '
                        '**외부 API 미사용 · 원본 불변 · 수정 완료 아님**. '
                        '회사에서 로컬 모델 처리가 허용된 데이터에만 사용하세요.')
            with gr.Row():
                agent_use_step4 = gr.Button('④ 검토 CSV 불러오기', variant='secondary')
                agent_input_csv = gr.File(label='④ 검사 검토 CSV', file_count='single',
                                          file_types=['.csv'], type='filepath')
            with gr.Row():
                agent_model = gr.Textbox(label='로컬 Ollama 모델', value='qwen2.5:7b-instruct')
                agent_limit = gr.Number(label='최대 제안 생성 건수', value=10,
                                        precision=0, minimum=1, maximum=200)
                agent_include_before = gr.Checkbox(
                    label='실제 보완 전 값 일부를 로컬 모델에 전달 (회사 사용 정책 확인)', value=True)
            agent_button = gr.Button('Qwen2.5 한국어 수정 제안 생성', variant='secondary')
            agent_notice = gr.Markdown('')
            agent_preview = gr.Dataframe(label='Correction Agent 결과 미리보기 (처음 30건)',
                                         interactive=False, wrap=True)
            agent_csv = gr.File(label='수정 제안 CSV — ⑤에 사용', interactive=False)


        with gr.Tab('⑤ 보완 전후 Excel 반영'):
            gr.Markdown('③ Excel과 ④ 검토 CSV 또는 로컬 Correction Agent의 새 CSV를 결합합니다. **원본은 그대로 보존**합니다.')
            with gr.Row():
                step5_import = gr.Button('현재 ③·④ 결과 불러오기', variant='secondary')
            step5_inspection = gr.File(label='검사결과서 Excel', file_count='single',
                                       file_types=['.xlsx'], type='filepath')
            step5_draft = gr.File(label='③ 보완결과서 초안 Excel', file_count='single',
                                  file_types=['.xlsx'], type='filepath')
            step5_review = gr.File(label='④ 보완 전후 검토 CSV (Correction Agent CSV 가능)', file_count='single',
                                   file_types=['.csv'], type='filepath')
            step5_info = gr.Markdown('')
            step5_button = gr.Button('보완 전후 Excel 반영', variant='primary')
            step5_status = gr.Markdown('')
            step5_download = gr.File(label='⑤ 반영된 보완결과서 다운로드', interactive=False)
            gr.Markdown('※ 실제 데이터 수정·재검증이 아닌, 보완 전후 검토 기록을 기존 Excel에 작성하는 단계입니다.')

        agent_use_step4.click(
            lambda path: gr.update(value=path),
            inputs=[evidence_csv], outputs=[agent_input_csv],
        )
        agent_button.click(
            write_local_correction_suggestions,
            inputs=[agent_input_csv, agent_model, agent_limit, agent_include_before],
            outputs=[agent_notice, agent_preview, agent_csv],
            concurrency_limit=1,
        )
        step5_import.click(
            load_step5_inputs,
            inputs=[report_input, report_download, matcher_inspection, evidence_csv, agent_csv],
            outputs=[step5_inspection, step5_draft, step5_review, step5_info],
        )
        step5_button.click(
            apply_step5_review,
            inputs=[step5_inspection, step5_draft, step5_review],
            outputs=[step5_status, step5_download],
            concurrency_limit=1,
        )

        folder_browse.click(choose_local_folder,
            inputs=[matcher_folder], outputs=[matcher_folder, browse_notice],
            concurrency_limit=1,
        )
        matcher_mode.change(
            lambda mode: (gr.update(visible=mode == 'folder'),
                          gr.update(visible=mode == 'zip')),
            inputs=[matcher_mode], outputs=[matcher_folder, matcher_zip]
        )
        matcher_button.click(
            match_error_files,
            inputs=[matcher_inspection, matcher_mode, matcher_folder, matcher_zip],
            outputs=[matcher_notice, matcher_preview, matcher_csv, matcher_json],
            concurrency_limit=1,
        )
        inspect_button.click(
            inspect_before_after,
            inputs=[matcher_inspection, matcher_mode, matcher_folder, matcher_zip,
                    final_dataset_removed],
            outputs=[evidence_notice, evidence_preview, evidence_csv, evidence_json],
            concurrency_limit=1,
        )
        uploads.change(on_uploaded_files, [uploads, workspace], [override, identified])
        workspace.change(on_uploaded_files, [uploads, workspace], [override, identified])
        add_button.click(register_workbooks, [uploads, workspace, override],
                         [summary, conflicts, selector, register_msg, uploads, override], concurrency_limit=1)
        refresh_button.click(refresh, inputs=[workspace], outputs=[summary, conflicts, selector])
        workspace.change(refresh, inputs=[workspace], outputs=[summary, conflicts, selector])
        selector.change(conflict_detail, [selector, workspace],
                        [company, inspection, status, info])
        approve_button.click(lambda k, s, r, n, a, w: change_decision(k, s, r, n, a, 'approve', w),
                             [selector, source, reviewer, reason, ack, workspace],
                             [summary, conflicts, selector, outcome], concurrency_limit=1)
        revoke_button.click(lambda k, s, r, n, a, w: change_decision(k, s, r, n, a, 'revoke', w),
                            [selector, source, reviewer, reason, ack, workspace],
                            [summary, conflicts, selector, outcome], concurrency_limit=1)
        make_report.click(create_report,
                          [report_input, report_round],
                          [report_notice, report_download], concurrency_limit=1)
        demo.load(refresh, inputs=[workspace], outputs=[summary, conflicts, selector])
    return demo


if __name__ == '__main__':
    build_app().launch(server_name='127.0.0.1', share=False, inbrowser=True,
                       show_error=True, footer_links=[], css=CSS)
