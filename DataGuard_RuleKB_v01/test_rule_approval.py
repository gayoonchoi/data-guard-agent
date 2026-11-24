"""실제 회사 파일을 사용하지 않는 승인 오버레이 회귀 테스트.

python test_rule_approval.py
"""
from __future__ import annotations
import json
from pathlib import Path
import tempfile
import unittest

from rule_approval import approve_conflict, list_conflicts, revoke_conflict
from validate_dataset import load_rule_set, value_issues


class RuleApprovalTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        group_dir = self.root / 'rules' / '991'
        group_dir.mkdir(parents=True)
        self.group_file = group_dir / 'group_01.json'
        self.group = {
            'project_id': '991', 'group_id': 1,
            'dataset_group_name': '검증용 가상 JSON',
            'source': {'sheet_differences': [{
                'cell': 'K15', 'company': '$d{4}.jpg',
                'inspection': r'^\d{5}\.jpg$',
            }]},
            'rules': [{
                'rule_id': '991-G1-001', 'field': 'imagePath', 'path': '$.imagePath',
                'data_type': 'string', 'required': True, 'nullable': False,
                'allowed_values': None, 'minimum': None, 'maximum': None,
                'pattern': '$d{4}.jpg', 'available_field_checks': ['type', 'required', 'nullability'],
                'manual_review_flags': ['기업/검사 규칙 불일치: pattern 셀 K15 (수동확인 전 검사 금지)'],
                'source': {'settings_cells': {'pattern': 'K15'}},
            }],
        }
        self._save_group()
        (self.root / 'catalog.json').write_text(json.dumps({
            'projects': {'991': {'groups': {'1': {'file': 'rules/991/group_01.json'}}}}
        }), encoding='utf-8')
        self.key = '991/1/K15'

    def tearDown(self):
        self.temp.cleanup()

    def _save_group(self):
        self.group_file.write_text(json.dumps(self.group, ensure_ascii=False), encoding='utf-8')

    def _rule(self):
        return load_rule_set(self.root, '991', 1)['rules'][0]

    def test_approval_enables_only_selected_check_and_preserves_original(self):
        original_bytes = self.group_file.read_bytes()
        self.assertEqual(list_conflicts(self.root)[0]['status'], '승인 대기')
        self.assertNotIn('pattern', self._rule()['available_field_checks'])
        approve_conflict(self.root, self.key, 'inspection', '테스터', '가상 데이터에 숫자 5자리 적용')
        rule = self._rule()
        self.assertEqual(rule['pattern'], r'^\d{5}\.jpg$')
        self.assertIn('pattern', rule['available_field_checks'])
        self.assertEqual(value_issues('0001.jpg', rule), ['pattern_mismatch'])
        self.assertEqual(value_issues('00001.jpg', rule), [])
        self.assertEqual(self.group_file.read_bytes(), original_bytes)
        self.assertTrue(list_conflicts(self.root)[0]['status'].startswith('승인됨'))
        self.assertIn('"action": "approve"',
                      (self.root / 'rule_approval_history.jsonl').read_text(encoding='utf-8'))

    def test_reimport_changed_original_invalidates_effective_approval(self):
        approve_conflict(self.root, self.key, 'inspection', '테스터', '기존 규칙 승인')
        self.group['source']['sheet_differences'][0]['inspection'] = r'^\d{6}\.jpg$'
        self._save_group()
        self.assertIn('재검토 필요', list_conflicts(self.root)[0]['status'])
        self.assertNotIn('pattern', self._rule()['available_field_checks'])

    def test_reimport_changed_rule_disables_stale_approval(self):
        approve_conflict(self.root, self.key, 'inspection', '테스터', '기존 규칙 승인')
        self.group['rules'][0]['required'] = False
        self._save_group()
        self.assertIn('재검토 필요', list_conflicts(self.root)[0]['status'])
        self.assertNotIn('pattern', self._rule()['available_field_checks'])

    def test_revoke_restores_hold(self):
        approve_conflict(self.root, self.key, 'inspection', '테스터', '근거 확인')
        revoke_conflict(self.root, self.key, '테스터', '승인 재검토')
        self.assertNotIn('pattern', self._rule()['available_field_checks'])
        self.assertEqual(list_conflicts(self.root)[0]['status'], '승인 대기')
        history = (self.root / 'rule_approval_history.jsonl').read_text(encoding='utf-8').splitlines()
        self.assertEqual([json.loads(x)['action'] for x in history], ['approve', 'revoke'])

    def test_approval_requires_reason_and_known_key(self):
        with self.assertRaises(ValueError):
            approve_conflict(self.root, self.key, 'inspection', '테스터', '')
        with self.assertRaises(ValueError):
            approve_conflict(self.root, '991/1/A1', 'inspection', '테스터', '내용 확인')


if __name__ == '__main__':
    unittest.main()
