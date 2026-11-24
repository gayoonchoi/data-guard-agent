"""DataGuard Agent — TTA 검사결과서에서 구문 규칙 지식 사전 추출.

Python 표준 라이브러리만 사용. 대용량 오류 목록 시트를 읽지 않고
'구축기업 구문규칙' / '검사 구문규칙' 및 '검사 종합 현황'만 추출합니다.
원본 파일은 읽기 전용으로 접근하고 변경하지 않습니다.

실행 예시:
    python extract_rules.py --input-dir "C:\\자료\\검사결과서" --output-dir "C:\\DataGuard\\rule_kb"
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import shutil
import tempfile
from datetime import datetime
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"m": MAIN, "r": REL}
TAG_ROW = f"{{{MAIN}}}row"
TAG_CELL = f"{{{MAIN}}}c"
TAG_SI = f"{{{MAIN}}}si"
TAG_T = f"{{{MAIN}}}t"
COLS = "ABCD"
HEADERS = {
    "타입": "data_type", "필수여부": "required", "유효값": "allowed_values",
    "최소값": "minimum", "최대값": "maximum", "패턴": "pattern",
    "기타": "other", "null 허용여부": "nullable", "설명": "description", "비고": "notes",
}
PROJECTS = {"007": "기후 변수 임계 임산물 생산 예측 학습데이터", "013": "카자흐스탄 모국어 사용자의 한국어 음성데이터", "017": "터널 결함 이미지 데이터셋", "019": "국토안전관리원 스마트 안전장비 실증현장 이동형 CCTV 기반 AI 이벤트 데이터", "022": "초거대 AI 근로감독관 지원시스템 구축 LLM 학습데이터"}


def load_shared_strings(z: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    values: list[str] = []
    with z.open("xl/sharedStrings.xml") as fp:
        for _, elem in ET.iterparse(fp, events=("end",)):
            if elem.tag == TAG_SI:
                values.append("".join(node.text or "" for node in elem.iter(TAG_T)))
                elem.clear()
    return values


def get_sheets(z: zipfile.ZipFile) -> list[tuple[str, str]]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {node.get("Id"): node.get("Target") for node in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    result = []
    for entry in wb.findall("m:sheets/m:sheet", NS):
        target = rels[entry.get(f"{{{REL}}}id")].replace("\\", "/")
        filename = "xl/" + target.lstrip("/").removeprefix("xl/")
        result.append((entry.get("name"), filename))
    return result


def read_nonempty_rows(z: zipfile.ZipFile, filename: str, strings: list[str]) -> list[tuple[int, dict[str, object]]]:
    records = []
    with z.open(filename) as fp:
        for _, row in ET.iterparse(fp, events=("end",)):
            if row.tag != TAG_ROW:
                continue
            row_id = int(row.get("r"))
            cells: dict[str, object] = {}
            for cell in row.findall(TAG_CELL):
                v = cell.find(f"{{{MAIN}}}v")
                instr = cell.find(f"{{{MAIN}}}is")
                text = v.text if v is not None else ("".join(x.text or "" for x in instr.iter(TAG_T)) if instr is not None else None)
                if text is None:
                    continue
                typ = cell.get("t")
                if typ == "s":
                    text = strings[int(text)]
                elif typ in ("n", None):
                    # XML 숫자 타입을 보존하되 규칙 파싱 단계에서 다시 숫자로 변환
                    try:
                        text = int(text) if re.fullmatch(r"-?\d+", str(text)) else float(text)
                    except (ValueError, OverflowError):
                        pass
                elif typ == "b":
                    text = text == "1"
                if str(text).strip():
                    cells[cell.get("r")] = text
            if cells:
                records.append((row_id, cells))
            row.clear()
    return records


def col(cell_ref: str) -> str:
    return re.match(r"[A-Z]+", cell_ref).group(0)


def row_values(row: dict[str, object]) -> dict[str, object]:
    return {col(k): v for k, v in row.items()}


def cast_num(x):
    if x is None or x == "":
        return None
    if isinstance(x, (int, float)):
        return x
    sx = str(x).strip()
    try:
        return int(sx) if re.fullmatch(r"-?\d+", sx) else float(sx)
    except ValueError:
        return None


def parse_enumeration(text, data_type):
    if text is None:
        return None
    raw = str(text)
    try:
        parsed = next(csv.reader(io.StringIO(raw), skipinitialspace=True))
    except (csv.Error, StopIteration):
        return None
    if not parsed:
        return None
    values = []
    for item in parsed:
        if data_type in ("number", "integer", "int"):
            number = cast_num(item)
            if number is None:
                return None  # 숫자가 아닌 유효값은 추정하지 않고 원문만 보존
            values.append(number)
        else:
            values.append(item.strip())
    return values


def build_path(level: int, name: str, hierarchy: dict[int, dict]) -> tuple[str, list[dict]]:
    for dep in list(hierarchy):
        if dep >= level:
            del hierarchy[dep]
    if level > 0 and level-1 not in hierarchy:
        raise ValueError(f"계층 부모가 없습니다: level={level}, name={name!r}")
    prefix = [dict(s) for s in hierarchy[level-1]["steps"]] if level else []
    kind = "item" if name in ("{}", "[]", "$value$") else "key"
    steps = prefix + ([{"kind": "item"}] if kind == "item" else [{"kind": "key", "name": name}])
    path = "$"
    for s in steps:
        if s["kind"] == "item":
            path += "[*]"
        else:
            key = s["name"]
            path += "." + key if re.fullmatch(r"[\w가-힣]+", key) else "[" + json.dumps(key, ensure_ascii=False) + "]"
    hierarchy[level] = {"steps": steps, "path": path}
    return path, steps


def convert_rule_rows(rows: list[tuple[int, dict]], project_id: str, group_id: int, sheet_name: str) -> list[dict]:
    if len(rows) < 3:
        return []
    r2 = next((rv for nr, rv in rows if nr == 2), {})
    header_columns = {HEADERS[str(value).strip()]: col(ref) for ref, value in r2.items() if str(value).strip() in HEADERS}
    if not {"data_type", "required"}.issubset(header_columns):
        raise ValueError(f"규칙 표 헤더를 찾을 수 없습니다: {sheet_name}: {header_columns}")
    field_columns = [chr(i) for i in range(ord("A"), ord(header_columns["data_type"]) )]
    if not field_columns:
        raise ValueError(f"필드 열 구성이 잘못되었습니다: {sheet_name}")
    stack = {}
    rules = []
    for rn, cells in rows:
        if rn <= 2:
            continue
        v = row_values(cells)
        filled = [(index, v[c]) for index, c in enumerate(field_columns) if c in v and str(v[c]).strip()]
        if not filled:
            continue
        if len(filled) != 1:
            raise ValueError(f"항목 계층이 불명확합니다: {sheet_name} 행 {rn}: {filled}")
        level, fname = filled[0]
        fname = str(fname)
        path, path_steps = build_path(level, fname, stack)
        raw = {key: v.get(cc) for key, cc in header_columns.items()}
        dtype = str(raw.get("data_type") or "").strip().lower() or None
        required = {"Y": True, "N": False}.get(str(raw.get("required") or "").upper())
        nullable = {"Y": True, "N": False}.get(str(raw.get("nullable") or "").upper())
        minimum = cast_num(raw.get("minimum"))
        maximum = cast_num(raw.get("maximum"))
        allowed = parse_enumeration(raw.get("allowed_values"), dtype)
        pattern = raw.get("pattern")
        issues = []
        if pattern is not None:
            try:
                re.compile(str(pattern))
            except re.error as e:
                issues.append(f"정규식 컴파일 실패: {e}")
            if str(pattern).strip().startswith("'") or str(pattern).strip().endswith("'"):
                issues.append("패턴 문자열의 따옴표 원문 확인 필요")
        if raw.get("allowed_values") is not None and allowed is None:
            issues.append("유효값 자동 파싱 실패: 원문 확인 필요")
        if dtype not in ("string", "integer", "int", "number", "object", "array", "boolean", "bool"):
            issues.append(f"자동 인식하지 못한 자료형: {dtype}")
        if raw.get("notes") and project_id == "022":
            issues.append("메시지 길이·순서·역할별 본문 조건: 별도 그룹 수준 검증 로직 필요")
        if project_id == "019" and ("point[*]" in str(raw.get("notes") or "") or "0초과" in str(raw.get("description") or "")):
            issues.append("비고의 좌표별 범위 또는 엄격 부등식은 별도 검증 로직 필요")
        if project_id == "017" and fname in ("category_id", "category_name"):
            issues.append("category_id ↔ category_name의 1:1 매핑 조건은 별도 교차 필드 검증 필요")
        if project_id == "007" and raw.get("notes") and "코드값 시트 참조" in str(raw["notes"]):
            issues.append("참조 코드값 시트가 현재 검사결과서 내 존재하는지 확인 필요")
        checks = ["type"]
        if required is not None: checks.append("required")
        if nullable is not None: checks.append("nullability")
        if minimum is not None: checks.append("minimum")
        if maximum is not None: checks.append("maximum")
        if allowed is not None: checks.append("enum")
        if pattern is not None and not any("정규식 컴파일 실패" in s for s in issues): checks.append("pattern")
        applies_to_role = None
        if project_id == "022" and fname == "content" and raw.get("notes"):
            role = re.search(r"\b(user|assistant|system)\.content\s*조건", str(raw["notes"]))
            applies_to_role = role.group(1) if role else None
        rule = {
            "pattern_applies_to_role": applies_to_role,
            "rule_id": f"{project_id}-G{group_id}-{len(rules)+1:03d}",
            "field": fname, "path": path, "path_steps": path_steps, "depth": level,
            "data_type": dtype, "required": required, "nullable": nullable,
            "allowed_values": allowed, "allowed_values_raw": raw.get("allowed_values"),
            "minimum": minimum, "maximum": maximum, "pattern": pattern,
            "other_raw": raw.get("other"),
            "description": raw.get("description"), "notes": raw.get("notes"),
            "available_field_checks": checks, "manual_review_flags": issues,
            "source": {"sheet": sheet_name, "row": rn, "field_cell": f"{field_columns[level]}{rn}",
                       "settings_cells": {k: f"{c}{rn}" for k,c in header_columns.items()}},
            "original_settings": raw,
        }
        rules.append(rule)
    return rules


def group_info(rows):
    for n, cells in rows:
        if n == 5 and "B5" in cells:
            return str(cells["B5"])
    return None


def compare_pairs(a, b):
    # 내용과 좌표가 같은지 비교. 서식은 고려하지 않음.
    da = {k: v for _, row in a for k,v in row.items()}
    db = {k: v for _, row in b for k,v in row.items()}
    keys = sorted(set(da) | set(db))
    differences = [{"cell": k, "company": da.get(k), "inspection": db.get(k)} for k in keys if da.get(k)!=db.get(k)]
    return differences


def _normalized_name(name: str) -> str:
    """파일명에 붙은 (8), 공백, 밑줄 같은 표기차를 무시해 사업명을 비교한다."""
    return re.sub(r"[^0-9a-z가-힣]+", "", str(name).lower())


def _name_candidate_ids(filename: Path, catalog: dict | None = None) -> list[str]:
    """Identify a named dataset without trusting a manually entered project ID.

    Match known project names to the filename, preferring the longest unique match.
    Records from a prior erroneous import should not override known project names.
    """
    source = _normalized_name(filename.stem)
    known = dict(PROJECTS)
    for pid, entry in (catalog or {}).get("projects", {}).items():
        if pid not in known:
            known[pid] = entry.get("name", "")
    candidates = []
    for pid, pname in known.items():
        name = _normalized_name(pname)
        if len(name) >= 9 and name in source:
            candidates.append((len(name), pid))
    if not candidates:
        return []
    longest = max(n for n, _ in candidates)
    return sorted(pid for n, pid in candidates if n == longest)


def identify_project_id(filename: Path, catalog: dict | None = None,
                        override: str | None = None) -> str:
    """Resolve one workbook's *actual* business ID, refusing contradictory overrides.

    - (007), (019), (099), (100) are trustworthy three-digit prefixes.
    - A different known business name in the filename is a contradiction.
    - A manually entered ID is only a fallback for a filename without an ID/name.
    """
    prefix = re.match(r"^\((\d{3})\)", filename.name)
    filename_pid = prefix.group(1) if prefix else None
    matched = _name_candidate_ids(filename, catalog)
    if len(matched) > 1:
        raise ValueError(f"파일명에서 사업명이 모호합니다: {filename.name} ({', '.join(matched)})")
    name_pid = matched[0] if matched else None
    if filename_pid and name_pid and filename_pid != name_pid:
        raise ValueError(
            f"파일명 앞의 사업 ID {filename_pid}와 사업명으로 확인한 ID {name_pid}가 다릅니다. "
            "검사결과서 파일명/원본을 확인해 주세요."
        )
    detected = filename_pid or name_pid
    manual = None
    if override is not None and str(override).strip():
        val = str(override).strip()
        if not re.fullmatch(r"\d{1,3}", val):
            raise ValueError("사업 ID는 숫자 1~3자리로 입력하세요. 예: 7 → 007, 99 → 099")
        manual = val.zfill(3)
        if detected and manual != detected:
            raise ValueError(
                f"등록 차단: 업로드한 '{filename.name}'은 {detected}번 검사결과서입니다. "
                f"입력한 사업 ID {manual}로 규칙을 복제할 수 없습니다. "
                f"{manual}번의 실제 검사결과서를 업로드해 주세요."
            )
    if detected:
        return detected
    if manual:
        return manual
    raise ValueError(
        f"사업 ID를 식별하지 못했습니다: {filename.name}. "
        "파일명에 (099)처럼 번호를 표시하거나, 사업 ID를 직접 입력해 주세요."
    )


def extract_workbook(filename: Path, output_dir: Path, *,
                     project_id: str | None = None, catalog: dict | None = None):
    pid = identify_project_id(filename, catalog=catalog, override=project_id)
    if pid not in PROJECTS:
        known_name = (catalog or {}).get('projects', {}).get(pid, {}).get('name')
        PROJECTS[pid] = known_name or re.sub(
            r"^\(\d+\).*?검사결과서[_\s]*", "", filename.stem
        ).strip()
    with zipfile.ZipFile(filename) as z:
        strings = load_shared_strings(z)
        targets = [(name,p) for name,p in get_sheets(z) if '구문규칙' in name or '검사 종합 현황' in name]
        result = {name: read_nonempty_rows(z,p,strings) for name,p in targets}
    groups = {}
    for sname, rows in result.items():
        if "구축기업 구문규칙" not in sname:
            continue
        prefix = re.match(r"^(\d+)\.\s*", sname)
        gid = int(prefix.group(1)) if prefix else 1
        inspect = f"{gid}. 검사 구문규칙" if prefix else "검사 구문규칙"
        if inspect not in result:
            raise ValueError(f"비교 대상 시트가 없습니다: {filename.name} / {inspect}")
        summary = f"{gid}. 검사 종합 현황" if prefix else "검사 종합 현황"
        if summary not in result and f"1. 검사 종합 현황" in result:
            summary = "1. 검사 종합 현황"
        differences = compare_pairs(rows, result[inspect])
        rules = convert_rule_rows(rows,pid,gid,sname)
        changed_cells = {d["cell"]: d for d in differences}
        for rule in rules:
            conflicting = [
                {"setting": key, **changed_cells[cell]}
                for key, cell in rule["source"]["settings_cells"].items()
                if cell in changed_cells
            ]
            if conflicting:
                rule["source"]["conflicting_settings"] = conflicting
                for item in conflicting:
                    rule["manual_review_flags"].append(
                        f"기업/검사 규칙 불일치: {item['setting']} 셀 {item['cell']} (수동확인 전 검사 금지)"
                    )
                    affected = {"pattern": "pattern", "allowed_values": "enum", "minimum": "minimum",
                                "maximum": "maximum", "required": "required", "nullable": "nullability", "data_type": "type"}.get(item["setting"])
                    if affected and affected in rule["available_field_checks"]:
                        rule["available_field_checks"].remove(affected)
        dataset = group_info(result.get(summary,[]))
        if not dataset:
            raise ValueError(f"검사 종합 현황에서 데이터명 확인 불가: {filename.name}/{sname}")
        group_level_rules = []
        if pid == "022":
            role_rule = next((r for r in rules if r["field"] == "role"), None)
            role_note = str(role_rule.get("notes") or "") if role_rule else ""
            count_match = re.search(r"메시지가 정확히\s*(\d+)개", role_note)
            role_order = role_rule.get("allowed_values") if role_rule else None
            if count_match and role_order and len(role_order) == int(count_match.group(1)):
                group_level_rules.append({
                    "rule_id": f"{pid}-G{gid}-MESSAGE",
                    "kind": "message_sequence",
                    "expected_count": int(count_match.group(1)),
                    "role_order": role_order,
                    "content_must_be_nonblank": "content" in role_note and "비어 있지 않아야" in role_note,
                    "source_rule_id": role_rule["rule_id"],
                    "applies_when": "단일 검증 레코드가 메시지 배열인 경우",
                })
        group = {
            "schema_version": "0.1", "project_id": pid, "project_name": PROJECTS[pid],
            "group_id": gid, "dataset_group_name": dataset, "rule_count": len(rules),
            "folder_matchers": [],  # 실제 데이터 폴더 경로는 명시적으로 연결해야 함
            "source": {
                "workbook_name": filename.name,
                "company_rule_sheet": sname,
                "inspection_rule_sheet": inspect,
                "overview_sheet": summary,
                "sheet_values_equal": len(differences)==0,
                "sheet_differences": differences,
            },
            "group_level_rules": group_level_rules,  # 원문에 명시된 메시지 수/순서/빈 본문 규칙만 구조화
            "rules": rules,
        }
        gdir = output_dir / "rules" / pid
        gdir.mkdir(parents=True,exist_ok=True)
        target = gdir / f"group_{gid:02d}.json"
        target.write_text(json.dumps(group,ensure_ascii=False,indent=2),encoding="utf-8")
        groups[gid] = {"name": dataset, "file": str(target.relative_to(output_dir)).replace('\\','/'),
                       "rule_count": len(rules), "sheet_equal": len(differences)==0,
                       "difference_count": len(differences),
                       "manual_flag_count": sum(bool(x['manual_review_flags']) for x in rules)}
    filehash=hashlib.sha256(filename.read_bytes()).hexdigest()
    if pid not in PROJECTS:
        first_group_name = next(iter(groups.values()))["name"] if groups else filename.stem
        PROJECTS[pid] = re.sub(r"\s*\((?:json|txt|csv)\)\s*$", "", first_group_name, flags=re.I)
    return pid, {"source_workbook": filename.name,"source_sha256":filehash,"name": PROJECTS[pid],
                 "groups": {str(i):g for i,g in sorted(groups.items())}}



INDEX_COLUMNS = [
    "사업ID", "사업명", "규칙그룹", "데이터그룹명", "규칙ID", "항목경로",
    "자료형", "필수여부", "NULL허용여부", "유효값(원문)", "최소값",
    "최대값", "정규식", "기타", "설명", "비고", "수동확인사항",
    "원본시트", "원본행"
]


def assemble_views(catalog: dict, output_dir: Path):
    """기존/추가 프로젝트를 합쳐 JSON 사전, 색인, 감사 기록을 재생성한다."""
    dictionary = {"schema_version": "0.1", "description": "TTA inspection syntax rules - extracted without automatically repairing data", "projects": {}}
    audit = {"projects": len(catalog["projects"]), "groups": 0, "rule_rows": 0,
             "company_inspection_mismatch_groups": [], "manual_flag_count": 0,
             "manual_flags_by_type": {}}
    flags = Counter()
    index_rows = []
    for pid, info in sorted(catalog["projects"].items()):
        proj = {"name": info["name"], "groups": {}}
        for gid, gs in sorted(info["groups"].items(), key=lambda p: int(p[0])):
            rule_path = output_dir / gs["file"]
            if not rule_path.is_file():
                raise FileNotFoundError(f"기존 규칙 파일을 찾을 수 없습니다: {rule_path}")
            group = json.loads(rule_path.read_text(encoding="utf-8"))
            rules = group["rules"]
            paths = [r["path"] for r in rules]
            if len(paths) != len(set(paths)):
                raise ValueError(f"규칙 경로가 중복됩니다: {pid}/G{gid}")
            proj["groups"][gid] = {
                "dataset_group_name": group["dataset_group_name"],
                "folder_matchers": group.get("folder_matchers", []),
                "rules_by_path": {r["path"]: r for r in rules},
                "group_level_rules": group.get("group_level_rules", []),
                "source": group["source"]
            }
            audit["groups"] += 1
            audit["rule_rows"] += len(rules)
            if not gs["sheet_equal"]:
                audit["company_inspection_mismatch_groups"].append(f"{pid}/{gid}")
            for r in rules:
                if r["manual_review_flags"]:
                    audit["manual_flag_count"] += 1
                    flags.update(r["manual_review_flags"])
                index_rows.append([
                    pid, info["name"], gid, group["dataset_group_name"],
                    r["rule_id"], r["path"], r["data_type"], r["required"], r["nullable"],
                    r["allowed_values_raw"], r["minimum"], r["maximum"], r["pattern"],
                    r["other_raw"], r["description"], r["notes"],
                    "; ".join(r["manual_review_flags"]), r["source"]["sheet"],
                    r["source"]["row"]
                ])
        dictionary["projects"][pid] = proj
    audit["manual_flags_by_type"] = dict(flags)
    return dictionary, index_rows, audit


def write_index(path: Path, rows: list[list]):
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(INDEX_COLUMNS)
        for row in rows:
            writer.writerow(["" if v is None else v for v in row])


def run(input_paths: list[Path], output_dir: Path, merge: bool = True,
        project_id_override: str | None = None):
    if not input_paths:
        raise ValueError("검사결과서 파일을 최소 하나 입력해야 합니다")
    if project_id_override and len(input_paths) != 1:
        raise ValueError('사업 ID 직접 입력은 파일 한 개를 등록할 때만 가능합니다.')
    output_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = output_dir / "catalog.json"
    if merge and manifest_path.exists():
        catalog = json.loads(manifest_path.read_text(encoding="utf-8"))
    else:
        catalog = {"schema_version":"0.1", "source_policy":"original source values retained; no automatic repair", "projects":{}}

    # 1단계: 모든 새 파일을 임시 폴더에 먼저 추출하고 검증한다.
    with tempfile.TemporaryDirectory(prefix="dataguard_rules_") as temp:
        staging = Path(temp)
        new_projects = {}
        for xlsx in input_paths:
            if not xlsx.is_file():
                raise FileNotFoundError(f"검사결과서 파일을 찾지 못했습니다: {xlsx}")
            if not xlsx.name.lower().endswith(".xlsx") or "검사결과서" not in xlsx.name:
                raise ValueError(f"검사결과서 이름의 .xlsx만 처리합니다: {xlsx.name}")
            result = extract_workbook(xlsx, staging,
                                      project_id=project_id_override, catalog=catalog)
            if not result:
                continue
            pid, info = result
            # Do not permit identical source bytes to masquerade as different businesses.
            for other_pid, old in catalog.get("projects", {}).items():
                if other_pid != pid and old.get("source_sha256") == info["source_sha256"]:
                    raise ValueError(
                        f"동일한 검사결과서가 이미 {other_pid}번 사업에 등록되어 있습니다. "
                        f"{pid}번에 중복 등록할 수 없습니다. 기존 {other_pid}번 항목을 확인해 주세요."
                    )
            for other_pid, new_info in new_projects.items():
                if other_pid != pid and new_info["source_sha256"] == info["source_sha256"]:
                    raise ValueError("동일한 Excel이 서로 다른 사업 ID로 입력되었습니다.")
            if pid in new_projects:
                raise ValueError(f"입력한 검사결과서에서 사업 ID가 중복됩니다: {pid}")
            if not info["groups"]:
                raise ValueError(f"구문 규칙 그룹을 찾지 못했습니다: {xlsx.name}")
            new_projects[pid] = info
            print(f"[{pid}] {len(info['groups'])} groups: " + ', '.join(f"{gid}={g['rule_count']}" for gid,g in info['groups'].items()))
        if not new_projects:
            raise FileNotFoundError("추출할 검사결과서를 찾지 못했습니다")

        # 2단계: 기존 사전 보존, 새 사업 추가 (같은 사업 ID는 새 버전으로 교체)
        candidate = json.loads(json.dumps(catalog, ensure_ascii=False))
        candidate["projects"].update(new_projects)
        # 독립된 스테이징 환경에 기존 모든 규칙과 새 규칙을 모아 검증한다.
        for pid, pinfo in candidate["projects"].items():
            for gid, g in pinfo["groups"].items():
                target = staging / g["file"]
                if not target.is_file():
                    existing = output_dir / g["file"]
                    if not existing.is_file():
                        raise FileNotFoundError(f"기존 규칙 파일 누락: {existing}")
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(existing, target)
        dictionary, index_rows, audit = assemble_views(candidate, staging)
        (staging / "catalog.json").write_text(json.dumps(candidate, ensure_ascii=False, indent=2), encoding="utf-8")
        (staging / "audit.json").write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")
        (staging / "rule_dictionary.json").write_text(json.dumps(dictionary, ensure_ascii=False, indent=2), encoding="utf-8")
        write_index(staging / "rule_index.csv", index_rows)
        mapping_path = output_dir / "folder_mapping.template.json"
        if mapping_path.is_file():
            folder_mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
        else:
            folder_mapping = {"note": "실제 폴더 경로 확인 전에는 자동 연결하지 않음", "projects": {}}
        for pid, info in candidate["projects"].items():
            groups_map = folder_mapping["projects"].setdefault(pid, {})
            for gid, group in info["groups"].items():
                groups_map.setdefault(gid, {"group_name": group["name"], "include_globs": [], "exclude_globs": []})
        (staging / "folder_mapping.template.json").write_text(
            json.dumps(folder_mapping, ensure_ascii=False, indent=2), encoding="utf-8"
        )

        # 3단계: 기존 산출물 백업 후 검증된 파일 복사
        backup_dir = output_dir / "_backups" / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        previous = ["catalog.json", "audit.json", "rule_dictionary.json", "rule_index.csv", "folder_mapping.template.json"]
        for pid in new_projects:
            for g in catalog.get("projects", {}).get(pid, {}).get("groups", {}).values():
                previous.append(g["file"])
            for gid, g in new_projects[pid]["groups"].items():
                previous.append(g["file"])
        old_files = [f for f in previous if (output_dir / f).is_file()]
        if old_files:
            for rel in old_files:
                destination = backup_dir / rel
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(output_dir / rel, destination)
            print("기존 산출물 백업:", backup_dir)
        for pid, info in new_projects.items():
            new_files = {group["file"] for group in info["groups"].values()}
            for old_group in catalog.get("projects", {}).get(pid, {}).get("groups", {}).values():
                obsolete = old_group["file"]
                if obsolete not in new_files and (output_dir / obsolete).is_file():
                    (output_dir / obsolete).unlink()  # already backed up above
            for gid, group in info["groups"].items():
                rel = group["file"]
                target = output_dir / rel
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(staging / rel, target)
        for rel in ["catalog.json", "audit.json", "rule_dictionary.json", "rule_index.csv", "folder_mapping.template.json"]:
            shutil.copy2(staging / rel, output_dir / rel)
    print("[OK] 총 사업:", audit["projects"], "규칙 그룹:", audit["groups"], "규칙 항목:", audit["rule_rows"])
    print("[OK] 저장 위치:", output_dir)
    print("[AUDIT]", json.dumps(audit, ensure_ascii=False))
    return candidate, audit


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="DataGuard 검사결과서 -> 통합 규칙 사전 JSON (기존 사업 보존)")
    mx = parser.add_mutually_exclusive_group(required=True)
    mx.add_argument("--input-file", type=Path, help="추가할 검사결과서 XLSX 파일 하나")
    mx.add_argument("--input-dir", type=Path, help="여러 검사결과서 XLSX를 모아 둔 폴더")
    parser.add_argument("--output-dir", type=Path, required=True, help="기존 DataGuard_RuleKB_v01 폴더")
    parser.add_argument("--project-id", help="파일명에서 사업 ID를 찾지 못할 때 지정하는 세 자리 ID(한 파일만)")
    parser.add_argument("--rebuild", action="store_true", help="기존 프로젝트를 무시하고 입력 파일만으로 새 사전 생성")
    args = parser.parse_args()
    if args.input_file is not None:
        paths = [args.input_file]
    else:
        if not args.input_dir.is_dir():
            parser.error(f"입력 폴더를 찾을 수 없습니다: {args.input_dir}")
        paths = sorted(p for p in args.input_dir.glob("*.xlsx") if not p.name.startswith("~$") and "검사결과서" in p.name)
    run(paths, args.output_dir, merge=not args.rebuild,
        project_id_override=args.project_id)
