"""DataGuard Agent Rule Engine v0.1 — 읽기 전용 1차 구문 검증.

- 규칙 사전을 활용해 JSON/CSV/JSONL 파일(또는 폴더)을 검사합니다.
- 원본 데이터는 수정하지 않고 검출 이력만 출력합니다.
- 폴더명에서 규칙을 자동 추정하지 않습니다. --project/--group를 반드시 지정합니다.
- 서술형/교차 필드/역할별 정규식/코드 참조는 아직 자동 검증하지 않습니다.

예: python validate_dataset.py --kb ./ --project 007 --group 1 --input-path "C:/data/D1_개인" --output-dir "C:/validation_results"
"""
from __future__ import annotations
import argparse
import csv
import codecs
import json
import re
from collections import Counter
from decimal import Decimal, InvalidOperation
from pathlib import Path
from rule_approval import apply_approvals

SUFFIXES = {'.csv','.json','.jsonl'}


def numeric(value):
    if isinstance(value,bool): return None
    try:
        if isinstance(value,str) and not value.strip():return None
        x=Decimal(str(value))
        return x if x.is_finite() else None
    except (ValueError, TypeError, InvalidOperation):return None


def value_issues(value, rule, apply_pattern=True):
    """활성화된 검사만 실행한다. 충돌/수동확인 대상 조건은 검사하지 않는다."""
    enabled = set(rule.get('available_field_checks', []))
    if value is None:
        return ['null_forbidden'] if 'nullability' in enabled and rule['nullable'] is False else []
    # 자료형이 충돌하면 다른 설정도 해석이 달라질 수 있으므로 확정 판정을 보류
    if 'type' not in enabled:
        return []
    t=rule['data_type']
    if t in ('int','integer','number'):
        n=numeric(value)
        if n is None:return ['invalid_number']
        if t in ('int','integer') and n!=n.to_integral_value():return ['expected_integer']
        measure=n
    elif t=='string':
        if not isinstance(value,str):return ['expected_string']
        measure=Decimal(len(value))
    elif t=='array':
        if not isinstance(value,list):return ['expected_array']
        measure=Decimal(len(value))
    elif t=='object':
        if not isinstance(value,dict):return ['expected_object']
        measure=None
    elif t in ('bool','boolean'):
        if type(value) is not bool:return ['expected_boolean']
        measure=None
    else:
        return [] # 사전의 자료형이 불명확하면 추측해서 검사하지 않음
    result=[]
    if measure is not None:
        if 'minimum' in enabled and rule['minimum'] is not None and measure<Decimal(str(rule['minimum'])):result.append('below_minimum')
        if 'maximum' in enabled and rule['maximum'] is not None and measure>Decimal(str(rule['maximum'])):result.append('above_maximum')
    allowed=rule['allowed_values']
    if 'enum' in enabled and allowed is not None:
        if t in ('int','integer','number'):
            if all(numeric(a)!=measure for a in allowed):result.append('not_in_allowed_values')
        elif value not in allowed:result.append('not_in_allowed_values')
    if apply_pattern and 'pattern' in enabled and rule['pattern'] is not None and isinstance(value,str):
        if 'pattern' in enabled and not any(
            '따옴표 원문 확인 필요' in f or '정규식 컴파일 실패' in f
            or ('기업/검사 규칙 불일치' in f and 'pattern' not in rule.get('approved_conflict_checks', []))
            for f in rule['manual_review_flags']
        ):
            try:
                if re.fullmatch(str(rule['pattern']),value) is None:result.append('pattern_mismatch')
            except re.error: pass
    return result


def all_targets(root, steps):
    """각 규칙을 중첩 JSON 객체에 적용할 경로를 찾아 (path, value, 존재여부) 반환."""
    def walk(obj, i, path):
        if i == len(steps):
            yield path,obj,True
            return
        st=steps[i]
        if st['kind']=='item':
            if isinstance(obj,list):
                for j,entry in enumerate(obj):
                    yield from walk(entry,i+1,f'{path}[{j}]')
            return
        name=st['name']
        if not isinstance(obj,dict):
            return  # 상위 객체 타입 오류에서 검출
        if name not in obj:
            if i==len(steps)-1:
                yield f'{path}.{name}',None,False
            return
        yield from walk(obj[name],i+1,f'{path}.{name}')
    yield from walk(root,0,'$')


def record_problems(record,rules,filename,record_id=None):
    for rule in rules:
        for path,value,exists in all_targets(record,rule['path_steps']):
            reasons=[]
            if not exists:
                if 'required' in rule.get('available_field_checks', []) and rule['required'] is True:reasons=['missing_required_field']
            else:
                # Excel CSV의 빈 필드는 null로 해석; JSON의 빈 문자열과는 별개
                if isinstance(value,str) and value=='' and filename.lower().endswith('.csv'):
                    value=None
                applies_to = rule.get('pattern_applies_to_role')
                apply_pattern = not (applies_to and isinstance(record,dict) and record.get('role') != applies_to)
                reasons=value_issues(value,rule,apply_pattern=apply_pattern)
            for reason in reasons:
                yield {'file':filename,'record':record_id,'path':path,'rule_id':rule['rule_id'],
                       'error_type':reason, 'observed_value':str(value)[:240] if exists else None,
                       'expected':{'type':rule['data_type'],'required':rule['required'],
                                   'nullable':rule['nullable'],'allowed_values':rule['allowed_values'],
                                   'minimum':rule['minimum'],'maximum':rule['maximum'],
                                   'pattern':rule['pattern']}}


def group_problems(messages, group_rules, filename, record_id):
    for rule in group_rules:
        if rule.get('kind') != 'message_sequence':
            continue
        expected=rule['expected_count']
        if len(messages) != expected:
            yield {'file': filename, 'record': record_id, 'path': '$',
                   'rule_id':rule['rule_id'],'error_type':'message_count_mismatch',
                   'observed_value':len(messages),'expected':expected}
        actual=[msg.get('role') if isinstance(msg,dict) else None for msg in messages]
        if actual != rule['role_order']:
            yield {'file':filename,'record':record_id,'path':'$[*].role',
                   'rule_id':rule['rule_id'],'error_type':'role_sequence_mismatch',
                   'observed_value':actual,'expected':rule['role_order']}
        if rule.get('content_must_be_nonblank'):
            for idx,msg in enumerate(messages):
                if isinstance(msg,dict) and 'content' in msg and isinstance(msg['content'],str) and not msg['content'].strip():
                    yield {'file':filename,'record':record_id,'path':f'$[{idx}].content',
                           'rule_id':rule['rule_id'],'error_type':'empty_content',
                           'observed_value':msg['content'][:240],'expected':'공백 제거 후 비어 있지 않은 문자열'}


def load_rule_set(kb_dir,project,group):
    if not re.fullmatch(r'\d{3}', str(project)):
        raise ValueError('사업 ID는 숫자 세 자리여야 합니다.')
    p=Path(kb_dir)/'rules'/str(project)/f'group_{int(group):02d}.json'
    if not p.is_file():raise FileNotFoundError(f'해당 규칙 그룹이 없습니다: {p}')
    data=json.loads(p.read_text(encoding='utf-8'))
    if data['project_id']!=project or data['group_id']!=int(group):raise ValueError('잘못된 규칙 그룹 파일')
    return apply_approvals(kb_dir, data)


def csv_rows(path):
    # CP949 인코딩을 인지하기 전에 이미 레코드를 출력하는 중복 위험을 방지합니다.
    encoding='utf-8-sig'
    decoder=codecs.getincrementaldecoder('utf-8-sig')(errors='strict')
    with path.open('rb') as raw:
        for chunk in iter(lambda:raw.read(1024*1024),b''):
            try:
                decoder.decode(chunk)
            except UnicodeDecodeError:
                encoding='cp949'
                break
        else:
            try:decoder.decode(b'',final=True)
            except UnicodeDecodeError:encoding='cp949'
    with path.open('r',encoding=encoding,newline='') as f:
        reader=csv.DictReader(f)
        for idx,rec in enumerate(reader,start=2):
            yield idx,rec


def records_of(path,project):
    ext=path.suffix.lower()
    if ext=='.csv':
        yield from csv_rows(path)
    elif ext=='.jsonl':
        with path.open('r',encoding='utf-8-sig') as fp:
            for idx,line in enumerate(fp,1):
                if line.strip():yield idx,json.loads(line)
    elif ext=='.json':
        with path.open('r',encoding='utf-8-sig') as fp:
            obj=json.load(fp)
        if project=='022':
            if isinstance(obj,dict):
                raise ValueError('022 JSON 객체의 메시지 배열 경로는 명시되지 않았습니다. 루트 메시지 배열만 검증합니다.')
            if not isinstance(obj,list):raise ValueError('022 JSON의 최상위가 메시지 배열이 아닙니다')
            yield 1,obj
        elif isinstance(obj,list):
            for idx,entry in enumerate(obj,1):yield idx,entry
        else:
            yield 1,obj


def execute(kb,project,group,input_path,output_dir):
    rs=load_rule_set(kb,project,group)
    source=Path(input_path).resolve()
    if not source.exists():raise FileNotFoundError(str(source))
    files=sorted(x for x in (source.rglob('*') if source.is_dir() else [source]) if x.is_file() and x.suffix.lower() in SUFFIXES)
    if not files:raise FileNotFoundError('CSV/JSON/JSONL 파일이 없습니다')
    out=Path(output_dir).resolve();out.mkdir(parents=True,exist_ok=True)
    result_path=out/f'issues_{project}_G{int(group)}.jsonl'
    counter=Counter();failed_files=[];processed=0;rows=0
    with result_path.open('w',encoding='utf-8') as target:
        for file in files:
            relative=file.relative_to(source) if source.is_dir() else Path(file.name)
            try:
                for n,rec in records_of(file,project):
                    rows+=1
                    if project=='022' and isinstance(rec,list):
                        issues=list(group_problems(rec,rs['group_level_rules'],str(relative).replace('\\','/'),n))
                        for mid,msg in enumerate(rec,1):
                            if not isinstance(msg,dict):
                                raise ValueError(f'{mid}번째 메시지가 JSON 객체가 아닙니다')
                            issues.extend(record_problems(msg,rs['rules'],str(relative).replace('\\','/'),f'{n}:{mid}'))
                    else:
                        if not isinstance(rec,dict):
                            raise ValueError(f'{n}번째 레코드가 JSON 객체가 아닙니다')
                        issues=record_problems(rec,rs['rules'],str(relative).replace('\\','/'),n)
                    for issue in issues:
                        target.write(json.dumps(issue,ensure_ascii=False)+'\n')
                        counter[issue['error_type']]+=1
                processed+=1
            except Exception as e:
                failed_files.append({'file':str(relative),'error':str(e)})
    report={
        'project_id':project,'group_id':int(group),'dataset_group_name':rs['dataset_group_name'],
        'original_files_modified':False,'files_scanned':processed,'records_scanned':rows,
        'issue_count':sum(counter.values()),'counts_by_error_type':dict(counter),
        'failures':failed_files,
        'not_implemented':[
            '규칙 적용 데이터 폴더 자동 식별','022 메시지 순서 외 서술형 조건 및 역할별 예외 검증',
            '교차 컬럼 및 코드 시트 참조 검증','검사결과서와 오류 비교',
            '자동 보완/승인/보완결과서 생성'
        ],
    }
    (out/'summary.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    return report


if __name__=='__main__':
    parser=argparse.ArgumentParser(description='외부 JSON 지식 사전을 이용한 읽기 전용 구문 검사 (PoC)')
    parser.add_argument('--kb',default='.',help='catalog.json 및 rules 폴더가 있는 위치')
    parser.add_argument('--project',required=True,help='규칙 사전에 등록된 숫자 3자리 사업 ID (예: 019)')
    parser.add_argument('--group',type=int,required=True)
    parser.add_argument('--input-path',required=True)
    parser.add_argument('--output-dir',required=True)
    args=parser.parse_args()
    print(json.dumps(execute(args.kb,args.project,args.group,args.input_path,args.output_dir),ensure_ascii=False,indent=2))
