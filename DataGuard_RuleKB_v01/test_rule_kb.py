"""테스트용 가상 입력만 사용합니다. 실제 회사 데이터는 로드하지 않습니다."""
import csv
import json
from collections import Counter
import unittest
from pathlib import Path
from validate_dataset import value_issues,record_problems,group_problems,load_rule_set
from rule_approval import list_conflicts

ROOT=Path(__file__).resolve().parent

class RuleKBTests(unittest.TestCase):
    def test_number_range(self):
        rs=load_rule_set(ROOT,'007',1)
        age=next(r for r in rs['rules'] if r['field']=='Age')
        self.assertEqual(value_issues(100,age),['above_maximum'])
        self.assertEqual(value_issues(30,age),[])

    def test_group_specific_year(self):
        a=load_rule_set(ROOT,'007',1)
        b=load_rule_set(ROOT,'007',3)
        year_a=next(r for r in a['rules'] if r['field']=='Year')
        year_b=next(r for r in b['rules'] if r['field']=='Year')
        self.assertEqual(value_issues(2024,year_a),[])
        self.assertIn('not_in_allowed_values',value_issues(2024,year_b))

    def test_nested_speaker_id_pattern(self):
        a=load_rule_set(ROOT,'013',1)
        rule=next(r for r in a['rules'] if r['field']=='UserID')
        self.assertEqual(value_issues('12345',rule),[])
        self.assertEqual(value_issues('1234',rule),['pattern_mismatch'])
        nested=next(r for r in a['rules'] if r['field']=='eojeolID' and 'Segmentation' in r['path'])
        self.assertEqual(nested['path'],'$.Segmentation.eojeols[*].eojeolID')

    def test_tunnel_category_and_bbox(self):
        a=load_rule_set(ROOT,'017',1)
        idrule=next(r for r in a['rules'] if r['field']=='category_id')
        self.assertEqual(value_issues(14,idrule),[])
        self.assertEqual(value_issues(999,idrule),['not_in_allowed_values'])
        bbox=next(r for r in a['rules'] if r['field']=='bbox')
        self.assertEqual(value_issues([1,2,3],bbox),['below_minimum'])
        self.assertEqual(value_issues([1,2,3,4],bbox),[])

    def test_message_group_rules(self):
        a=load_rule_set(ROOT,'022',1)
        good=[{'role':'user','content':'Q'},{'role':'assistant','content':'A'}]
        bad=[{'role':'assistant','content':' '},{'role':'user','content':'Q'}]
        self.assertEqual(list(group_problems(good,a['group_level_rules'],'test.json',1)),[])
        issue_codes=[x['error_type'] for x in group_problems(bad,a['group_level_rules'],'test.json',1)]
        self.assertEqual(issue_codes,['role_sequence_mismatch','empty_content'])

    def test_019_json_and_txt_rule_groups(self):
        js = load_rule_set(ROOT,'019',1)
        txt = load_rule_set(ROOT,'019',2)
        self.assertGreater(len(js['rules']), 0)
        self.assertGreater(len(txt['rules']), 0)
        json_text = next(r for r in js['rules'] if r['field']=='text')
        self.assertEqual(json_text['path'],'$.shapes[*].text')
        self.assertNotIn('$.text',{r['path'] for r in js['rules']})
        coords = next(r for r in js['rules'] if r['field']=='$value$')
        self.assertEqual(coords['path'],'$.shapes[*].points[*][*]')
        path_rule=next(r for r in js['rules'] if r['field']=='imagePath')
        conflicts = path_rule.get('source', {}).get('conflicting_settings', [])
        if any(x['setting'] == 'pattern' for x in conflicts):
            # 승인된 충돌 규칙만 실행 시점에 활성화됨. 원본 추출 JSON은 비활성 상태 유지.
            active = any(
                r['key'] == '019/1/K15' and r['status'].startswith('승인됨')
                for r in list_conflicts(ROOT)
            )
            if active:
                self.assertIn('pattern', path_rule['available_field_checks'])
            else:
                self.assertNotIn('pattern', path_rule['available_field_checks'])
            self.assertIn('기업/검사 규칙 불일치',' '.join(path_rule['manual_review_flags']))
            self.assertEqual(value_issues('00001.jpg',path_rule),[])
        w=next(r for r in txt['rules'] if r['field']=='w')
        self.assertTrue(any('엄격 부등식' in f for f in w['manual_review_flags']))

    def test_disabled_checks_never_fire(self):
        """설정 충돌로 비활성화한 규칙이 실제 데이터에서 판정을 내리지 않는지 검사."""
        rs = load_rule_set(ROOT, '007', 1)
        age = next(r for r in rs['rules'] if r['field'] == 'Age')
        over = dict(age)
        over['available_field_checks'] = [x for x in age['available_field_checks'] if x != 'maximum']
        self.assertEqual(value_issues(100, over), [])
        nul = dict(age)
        nul['available_field_checks'] = [x for x in age['available_field_checks'] if x != 'nullability']
        self.assertEqual(value_issues(None, nul), [])
        if 'enum' in age['available_field_checks']:
            enum_off = dict(age)
            enum_off['available_field_checks'] = [x for x in age['available_field_checks'] if x != 'enum']
            enum_off['allowed_values'] = [-1234567]
            self.assertEqual(value_issues(30, enum_off), [])

    def test_source_comparison_and_count(self):
        """사업 추가 후에도 개수/불일치 값을 고정하지 않고 전체 원본 JSON과 대조."""
        audit = json.loads((ROOT / 'audit.json').read_text(encoding='utf-8'))
        catalog = json.loads((ROOT / 'catalog.json').read_text(encoding='utf-8'))
        dictionary = json.loads((ROOT / 'rule_dictionary.json').read_text(encoding='utf-8'))
        projects = catalog['projects']

        self.assertEqual(audit['projects'], len(projects))
        self.assertEqual(set(dictionary['projects']), set(projects))

        total_groups = total_rules = manual_count = 0
        mismatch_groups = []
        manual_types = Counter()
        expected_index = set()
        for pid, project in projects.items():
            self.assertEqual(set(project['groups']), set(dictionary['projects'][pid]['groups']))
            for gid, info in project['groups'].items():
                total_groups += 1
                path = ROOT / info['file']
                self.assertTrue(path.is_file(), str(path))
                group = json.loads(path.read_text(encoding='utf-8'))
                rules = group['rules']
                total_rules += len(rules)
                self.assertEqual(group['project_id'], pid)
                self.assertEqual(str(group['group_id']), gid)
                self.assertEqual(info['rule_count'], len(rules))
                self.assertEqual(group['rule_count'], len(rules))
                source = group['source']
                differences = source['sheet_differences']
                self.assertEqual(source['sheet_values_equal'], not bool(differences))
                self.assertEqual(info['sheet_equal'], not bool(differences))
                self.assertEqual(info['difference_count'], len(differences))
                if differences:
                    mismatch_groups.append(f'{pid}/{gid}')
                rule_paths = [r['path'] for r in rules]
                self.assertEqual(len(rule_paths), len(set(rule_paths)), f'{pid}/{gid}: 중복 경로')
                lookup = dictionary['projects'][pid]['groups'][gid]['rules_by_path']
                self.assertEqual(set(lookup), set(rule_paths))
                self.assertEqual({r['rule_id'] for r in rules}, {r['rule_id'] for r in lookup.values()})

                changes = {d['cell']: d for d in differences}
                check_names = {
                    'pattern': 'pattern', 'allowed_values': 'enum', 'minimum': 'minimum',
                    'maximum': 'maximum', 'required': 'required', 'nullable': 'nullability',
                    'data_type': 'type',
                }
                for rule in rules:
                    key = (pid, str(gid), rule['rule_id'])
                    self.assertNotIn(key, expected_index)
                    expected_index.add(key)
                    flags = rule.get('manual_review_flags', [])
                    if flags:
                        manual_count += 1
                        manual_types.update(flags)
                    for setting, cell in rule['source']['settings_cells'].items():
                        if cell not in changes:
                            continue
                        self.assertTrue(
                            any('기업/검사 규칙 불일치' in flag for flag in flags),
                            f'{pid}/{gid}/{rule["path"]}: 불일치 미표시',
                        )
                        check = check_names.get(setting)
                        if check:
                            self.assertNotIn(check, rule['available_field_checks'],
                                             f'{pid}/{gid}/{rule["path"]}: 충돌 검사 비활성화 실패')

        self.assertEqual(audit['groups'], total_groups)
        self.assertEqual(audit['rule_rows'], total_rules)
        self.assertCountEqual(audit['company_inspection_mismatch_groups'], mismatch_groups)
        self.assertEqual(audit['manual_flag_count'], manual_count)
        self.assertEqual(audit['manual_flags_by_type'], dict(manual_types))
        with (ROOT / 'rule_index.csv').open(encoding='utf-8-sig', newline='') as fh:
            rows = list(csv.DictReader(fh))
        self.assertEqual(len(rows), total_rules)
        actual_index = {(r['사업ID'], r['규칙그룹'], r['규칙ID']) for r in rows}
        self.assertEqual(actual_index, expected_index)

if __name__=='__main__':unittest.main()
