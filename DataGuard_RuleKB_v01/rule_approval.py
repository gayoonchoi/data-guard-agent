"""DataGuard Rule Approval v0.1.

구축기업/검사 시트의 상충하는 pattern 셀에 한해서 명시적 담당자 결정을 기록한다.
원본 추출 규칙 JSON은 바꾸지 않고, 결정은 검증 실행시 조건부 오버레이로만 적용한다.
다른 설정 타입은 아직 안전하게 자동 활성화하지 않으며 계속 수동 확인 대상으로 둔다.
"""
from __future__ import annotations

from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import tempfile
from typing import Any

APPROVAL_FILE = 'rule_approvals.json'
HISTORY_FILE = 'rule_approval_history.jsonl'
APPROVABLE = {'pattern'}  # 확장 시 각 설정별 파서/검증/테스트가 필요함


def _read_json(path: Path, default: Any):
    return json.loads(path.read_text(encoding='utf-8')) if path.is_file() else default


def _store(root: Path):
    data = _read_json(root / APPROVAL_FILE, {'schema_version': '0.1', 'decisions': {}})
    if not isinstance(data, dict) or not isinstance(data.get('decisions'), dict):
        raise ValueError('승인 결정 파일 형식이 올바르지 않습니다.')
    return data


def _write_atomic(path: Path, content: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix='.rule_approval_', suffix='.tmp', dir=path.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(content)
            f.flush()
            os.fsync(f.fileno())
        os.replace(name, path)
    finally:
        if os.path.exists(name):
            os.unlink(name)


def _write_store(root: Path, data: dict):
    _write_atomic(root / APPROVAL_FILE, json.dumps(data, ensure_ascii=False, indent=2) + '\n')


def _append_history(root: Path, payload: dict):
    # 프로토타입 전용. 동시 작업은 UI 측의 단일 락으로 직렬화한다.
    with (root / HISTORY_FILE).open('a', encoding='utf-8') as f:
        f.write(json.dumps(payload, ensure_ascii=False) + '\n')
        f.flush()
        os.fsync(f.fileno())


def _decision_key(pid: str, gid: str | int, cell: str):
    return f'{pid}/{int(gid)}/{cell}'


def _rule_hash(rule: dict):
    raw = json.dumps(rule, sort_keys=True, ensure_ascii=False, separators=(',', ':'))
    return hashlib.sha256(raw.encode('utf-8')).hexdigest()


def list_conflicts(root: str | Path) -> list[dict]:
    """실제 JSON에 있는 시트 차이와 승인 상태를 매번 대조한다."""
    root = Path(root)
    catalog = _read_json(root / 'catalog.json', {'projects': {}})
    decisions = _store(root)['decisions']
    rows = []
    for pid, project in sorted(catalog.get('projects', {}).items()):
        for gid, meta in sorted(project.get('groups', {}).items(), key=lambda kv: int(kv[0])):
            group_file = (root / meta['file']).resolve()
            if not group_file.is_relative_to(root.resolve()):
                raise ValueError(f'규칙 그룹 경로가 지식 사전 밖을 가리킵니다: {meta["file"]}')
            group = _read_json(group_file, {})
            by_cell = {}
            for rule in group.get('rules', []):
                for setting, cell in rule.get('source', {}).get('settings_cells', {}).items():
                    by_cell[cell] = (rule, setting)
            for diff in group.get('source', {}).get('sheet_differences', []):
                cell = diff['cell']
                rule, setting = by_cell.get(cell, (None, '규칙과 연결되지 않은 셀'))
                key = _decision_key(pid, gid, cell)
                row = {
                    'key': key, 'project_id': pid, 'group_id': str(gid),
                    'group_name': group.get('dataset_group_name', ''),
                    'cell': cell, 'setting': setting,
                    'path': rule['path'] if rule else '',
                    'rule_id': rule['rule_id'] if rule else None,
                    'rule_hash': _rule_hash(rule) if rule else None,
                    'company': diff.get('company'),
                    'inspection': diff.get('inspection'),
                    'approvable': bool(rule and setting in APPROVABLE),
                }
                decision = decisions.get(key)
                if decision:
                    if _matching(decision, row):
                        row['status'] = '승인됨 (' + ('검사 규칙' if decision['source'] == 'inspection' else '구축기업 규칙') + ')'
                    else:
                        row['status'] = '재검토 필요 (원본 변경)'
                else:
                    row['status'] = '승인 대기' if row['approvable'] else '수동 처리 필요 (자동 승인 미지원)'
                rows.append(row)
    return rows


def _matching(decision, current):
    return (decision.get('project_id') == current['project_id']
            and str(decision.get('group_id')) == str(current['group_id'])
            and decision.get('cell') == current['cell']
            and decision.get('rule_id') == current['rule_id']
            and decision.get('rule_hash') == current['rule_hash']
            and decision.get('path') == current['path']
            and decision.get('setting') == current['setting']
            and decision.get('company') == current['company']
            and decision.get('inspection') == current['inspection']
            and decision.get('source') in ('company', 'inspection'))


def approve_conflict(root: str | Path, key: str, source: str, reviewer: str, reason: str):
    """선택된 Excel 출처의 원문 값을 승인. 새로운 값을 임의 생성하지 않는다."""
    root = Path(root)
    reviewer, reason = (reviewer or '').strip(), (reason or '').strip()
    if len(reviewer) < 2 or not reason:
        raise ValueError('검토자(2자 이상)와 승인 사유를 입력해 주세요.')
    if len(reviewer) > 80 or len(reason) > 2000:
        raise ValueError('검토자 또는 승인 사유가 너무 깁니다.')
    if source not in ('company', 'inspection'):
        raise ValueError('구축기업 또는 검사 규칙 중 하나를 선택해야 합니다.')
    row = next((r for r in list_conflicts(root) if r['key'] == key), None)
    if row is None:
        raise ValueError('현재 등록된 불일치 규칙이 아닙니다.')
    if not row['approvable']:
        raise ValueError('이 설정은 아직 자동 승인과 검증 활성화를 지원하지 않습니다. 수동 검토해 주세요.')
    approved_value = row[source]
    if row['setting'] == 'pattern':
        if not isinstance(approved_value, str) or not approved_value or len(approved_value) > 500:
            raise ValueError('빈 값 또는 너무 긴 정규식은 활성화할 수 없습니다.')
        try:
            re.compile(approved_value)
        except re.error as exc:
            raise ValueError(f'선택한 정규식이 컴파일되지 않습니다: {exc}') from exc
    store = _store(root)
    record = {
        'project_id': row['project_id'], 'group_id': row['group_id'], 'cell': row['cell'],
        'rule_id': row['rule_id'], 'rule_hash': row['rule_hash'],
        'path': row['path'], 'setting': row['setting'],
        'company': row['company'], 'inspection': row['inspection'],
        'source': source, 'approved_value': approved_value,
        'reviewer': reviewer, 'reason': reason,
        'approved_at': datetime.now(timezone.utc).isoformat(timespec='seconds'),
    }
    store['decisions'][key] = record
    _write_store(root, store)
    _append_history(root, {'action': 'approve', 'key': key, **record})
    return record


def revoke_conflict(root: str | Path, key: str, reviewer: str, reason: str):
    root = Path(root)
    reviewer, reason = (reviewer or '').strip(), (reason or '').strip()
    if len(reviewer) < 2 or not reason:
        raise ValueError('검토자(2자 이상)와 철회 사유를 입력해 주세요.')
    store = _store(root)
    previous = store['decisions'].pop(key, None)
    if previous is None:
        raise ValueError('철회할 승인 기록이 없습니다.')
    _write_store(root, store)
    _append_history(root, {
        'action': 'revoke', 'key': key, 'prior_decision': previous,
        'reviewer': reviewer, 'reason': reason,
        'at': datetime.now(timezone.utc).isoformat(timespec='seconds'),
    })
    return previous


def apply_approvals(root: str | Path, group: dict) -> dict:
    """현재 시트 차이와 결정 기록이 일치할 때만 복사본의 해당 검사를 활성화."""
    root = Path(root)
    decisions = _store(root)['decisions']
    if not decisions:
        return group
    pid, gid = group['project_id'], str(group['group_id'])
    by_cell = {}
    for rule in group.get('rules', []):
        for setting, cell in rule.get('source', {}).get('settings_cells', {}).items():
            by_cell[cell] = (rule, setting)
    approved = []
    for diff in group.get('source', {}).get('sheet_differences', []):
        key = _decision_key(pid, gid, diff['cell'])
        decision = decisions.get(key)
        if not decision:
            continue
        rule, setting = by_cell.get(diff['cell'], (None, None))
        if rule is None or setting not in APPROVABLE:
            continue
        current = {
            'project_id': pid, 'group_id': gid, 'cell': diff['cell'],
            'rule_id': rule['rule_id'], 'rule_hash': _rule_hash(rule),
            'path': rule['path'],
            'setting': setting, 'company': diff.get('company'),
            'inspection': diff.get('inspection'),
        }
        if not _matching(decision, current):
            continue  # 새 Excel 재등록 이후 변경된 원본에는 과거 승인을 적용하지 않음
        if 'pattern' == setting:
            val = current[decision['source']]
            if not isinstance(val, str):
                continue
            try:
                re.compile(val)
            except re.error:
                continue
            approved.append((rule['rule_id'], setting, val, key, decision['source']))
    if not approved:
        return group
    copied = deepcopy(group)
    index = {r['rule_id']: r for r in copied.get('rules', [])}
    for rule_id, setting, val, key, source in approved:
        r = index[rule_id]
        r[setting] = val
        if setting not in r.get('available_field_checks', []):
            r['available_field_checks'].append(setting)
        r.setdefault('approved_conflict_checks', []).append(setting)
        r.setdefault('approval_provenance', []).append({'key': key, 'source': source})
    return copied
