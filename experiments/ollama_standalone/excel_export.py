"""Dependency-free XLSX export of independent Ollama test results.

Does not edit any source file, dataset, or DataGuard workbook.  Creates a
new two-worksheet Office Open XML workbook from a generated/selected CSV.
"""
from __future__ import annotations

import csv
import math
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

GROUND_TRUTH_COL = "모범답안_비교용(Ground Truth)"
KEYWORD_COL = "자동 키워드 검사(참고용)"
RUN_ERROR_COL = "Ollama 실행 오류"
LEGACY_NAMES = {
    "모범답안_비교용": GROUND_TRUTH_COL,
    "자동키워드확인": KEYWORD_COL,
    "오류": RUN_ERROR_COL,
}

NOTES_ROWS = [
    ("구분", "컬럼명 / 분류값", "뜻 · 판정 기준", "작성 주체 및 주의 사항"),
    ("기본", "실행시각", "해당 Ollama 테스트를 실행한 시각. ISO 형식으로 저장함", "테스트 프로그램; 원본 CSV의 시간값 그대로 기록"),
    ("기본", "case_id", "독립 테스트의 사례 식별자(S01, S02 등)", "테스트 사례 CSV; 실제 파일 고유 ID가 아님"),
    ("기본", "파일명", "테스트 사례에 기입한 오류 파일명", "테스트 입력값; 별도의 원본 파일 접근을 의미하지 않음"),
    ("기본", "검사_프로퍼티", "검사 오류가 기록된 속성 또는 경로. $.text와 $.shapes[*].text는 다른 위치", "테스트 입력값. 파일 구조 및 경로 계층을 확인해야 함"),
    ("평가 기준", GROUND_TRUTH_COL, "사람이 정의한 모범 수정안, 비교 평가용 기준 정답(Ground Truth 후보)", "Ollama에 전달하지 않으며 모델을 학습시키지도 않음. 원본·규칙 검증 전에는 확정 진실로 단정하지 않음"),
    ("모델 호출", "모델호출상태", "로컬 Ollama API 호출 및 응답 처리의 성공/실패", "프로그램 자동 기록; 호출성공이어도 수정안이 정확하다는 뜻은 아님"),
    ("모델 호출", "실제응답모델", "로컬 Ollama API 응답에서 확인한 모델 이름", "Ollama API 응답"),
    ("모델 출력", "Ollama 제안", "Ollama가 작성한 보완 후 상태 수정 지시문", "Ollama 생성 텍스트; 실제 수정 완료를 의미하지 않음"),
    ("모델 출력", "Ollama 근거", "Ollama가 제안 이유로 생성한 설명", "Ollama 생성 텍스트; 근거가 사실·구문규칙과 일치하는지 독립 검증 필요"),
    ("모델 출력", "추가확인자료", "모델이 추가 확인이 필요하다고 제안한 자료·정보", "Ollama 생성 텍스트; 실제 해당 자료가 존재한다는 증거는 아님"),
    ("모델 출력", "사람검토필요", "모델의 human_review_required 응답 여부(True/False)", "Ollama 생성 판정; 최종 승인 여부와 무관"),
    ("성능", "소요시간_초", "로컬 Ollama API 요청에서 응답 수신까지 걸린 초 단위 시간", "프로그램 성능 측정; 예: 18.04 = 18.04초"),
    ("정확성", "수정안 정확성 평가", "적절 / 부적절 / 추가 검토 필요 중 하나", "담당자 또는 잠정 평가 기록; 자동 키워드 검사 결과가 아님"),
    ("정확성", "평가 근거", "해당 제안이 적절/부적절/검토 필요인 이유", "평가자 서술; 실제 데이터·해당 사업 규칙과 대조해 확인해야 함"),
    ("평가 주체", "평가 주체", "누가 어떤 수준으로 수정안의 정확성을 평가했는지 구분", "프로그램 기록. 아래 3개 분류값 참조"),
    ("평가 주체 값", "미평가", "모델 제안에 대한 사람 평가가 진행되지 않음", "새 시험 결과 기본값. 평가 상태는 '추가 검토 필요'"),
    ("평가 주체 값", "참고용 잠정 내용 평가", "기존 테스트 결과를 참고해 작성한 임시 내용 평가", "확정 승인 또는 실데이터 검증을 뜻하지 않음"),
    ("평가 주체 값", "담당자 수동 평가", "테스트 UI에서 담당자가 직접 평가값과 근거를 저장함", "담당자 판단이 기록된 상태. 별도의 원본·규칙 검토 권장"),
    ("정확성 값", "적절", "수정 방향과 근거가 제시된 구문규칙 및 확인된 오류와 부합", "검토된 사례에만 부여. 실제 보완 완료를 뜻하지 않음"),
    ("정확성 값", "부적절", "제안이 규칙·경로·오류 상태와 충돌하거나 근거 없는 수정 요구", "Case KB의 승인 사례로 저장하면 안 됨"),
    ("정확성 값", "추가 검토 필요", "규칙·실데이터·삭제 기록 등이 불충분하여 정답 여부 확정 불가", "새로운 Ollama 결과의 기본 상태"),
    ("자동 점검", KEYWORD_COL, "정해진 키워드 포함 여부 및 한글 존재 여부만 점검", "결정론적 Python 문자열 검사; 의미 정확도 점수가 아님"),
    ("실행 오류", RUN_ERROR_COL, "Ollama API 연결 실패, 시간 초과, JSON 응답 해석 실패 등 프로그램 실행 오류", "프로그램 자동 기록. 모델의 잘못된 수정 제안은 여기에 기록되지 않음"),
    ("개념 구분", "Rule Engine", "규칙/경로/자료형/필수값/허용값을 Python 코드로 판정하는 결정론적 검증", "동일 입력·규칙에서 같은 판정이 나오도록 설계. 모든 Python 코드가 Rule Engine인 것은 아님"),
    ("개념 구분", "Correction Agent", "로컬 Ollama로 규칙만으로 해결하기 어려운 보완 후 문장을 제안", "생성형 모델; 응답은 결정론적 검증과 담당자 승인 필요"),
    ("향후 구조", "Case-Based RAG", "승인된 과거 오류·보완 사례를 검색해 모델의 판단 근거로 제공", "현재 독립 테스트는 검색형 Case KB를 구현하지 않음; 향후 개발 대상"),
    ("데이터 보호", "학습·검색 구분", "비교용 모범답안을 저장한다고 자동 파인튜닝되거나 모델이 기억하지는 않음", "승인된 사례만 별도 KB에 넣어 검색하는 구조를 권장"),
    ("주의", "파일 삭제/제외", "폴더에서 찾을 수 없다는 사실과 삭제 완료 승인 사실은 구별", "파일 매칭 실패만으로 삭제 완료라고 단정하지 않음"),
    ("주의", "기존 CSV 호환", "예전 열 이름 '모범답안_비교용', '자동키워드확인', '오류'도 불러오기 가능", "새 Excel에서는 표준화된 명확한 컬럼명으로 출력"),
]

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
C = "http://schemas.openxmlformats.org/package/2006/content-types"
ET.register_namespace("", S)
ET.register_namespace("r", R)


def _column_letter(n: int) -> str:
    result = ""
    while n:
        n, x = divmod(n - 1, 26)
        result = chr(65 + x) + result
    return result


def _xml_bytes(root) -> bytes:
    content = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    # Keep the OPC package relationship/content-type roots in their default
    # namespaces (some spreadsheet consumers are sensitive to ns0: prefixes).
    if root.tag.startswith("{" + P + "}") or root.tag.startswith("{" + C + "}"):
        content = content.replace(b"xmlns:ns0=", b"xmlns=").replace(b"ns0:", b"")
    return content


def _sheet_xml(rows: list[list[str]], widths: list[float], *, result_sheet: bool) -> bytes:
    root = ET.Element(f"{{{S}}}worksheet")
    last_col = _column_letter(max((len(r) for r in rows), default=1))
    ET.SubElement(root, f"{{{S}}}dimension", {"ref": f"A1:{last_col}{len(rows)}"})
    views = ET.SubElement(root, f"{{{S}}}sheetViews")
    view = ET.SubElement(views, f"{{{S}}}sheetView", {"workbookViewId": "0"})
    ET.SubElement(view, f"{{{S}}}pane", {"ySplit": "1", "topLeftCell": "A2", "activePane": "bottomLeft", "state": "frozen"})
    ET.SubElement(root, f"{{{S}}}sheetFormatPr", {"defaultRowHeight": "18"})
    cols = ET.SubElement(root, f"{{{S}}}cols")
    for idx, width in enumerate(widths, start=1):
        ET.SubElement(cols, f"{{{S}}}col", {"min": str(idx), "max": str(idx), "width": str(width), "customWidth": "1"})
    sheet_data = ET.SubElement(root, f"{{{S}}}sheetData")
    for rnum, row in enumerate(rows, start=1):
        row_element = ET.SubElement(sheet_data, f"{{{S}}}row", {
            "r": str(rnum), "ht": "31" if rnum == 1 else ("42" if not result_sheet else "46"), "customHeight": "1"
        })
        for cnum, val in enumerate(row, start=1):
            if val is None:
                val = ""
            value = str(val)
            address = f"{_column_letter(cnum)}{rnum}"
            style = "1" if rnum == 1 else "2"
            if result_sheet and rnum > 1 and rows[0][cnum-1] == "소요시간_초":
                try:
                    numeric = float(value)
                except ValueError:
                    numeric = None
                if numeric is not None and math.isfinite(numeric):
                    c = ET.SubElement(row_element, f"{{{S}}}c", {"r": address, "s": style})
                    ET.SubElement(c, f"{{{S}}}v").text = str(numeric)
                    continue
            c = ET.SubElement(row_element, f"{{{S}}}c", {"r": address, "t": "inlineStr", "s": style})
            inline = ET.SubElement(c, f"{{{S}}}is")
            ET.SubElement(inline, f"{{{S}}}t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = value
    if result_sheet:
        ET.SubElement(root, f"{{{S}}}autoFilter", {"ref": f"A1:{last_col}{len(rows)}"})
    return _xml_bytes(root)


def _styles_xml() -> bytes:
    xml = (f'<styleSheet xmlns="{S}"><fonts count="2">'
           '<font><sz val="10"/><name val="Malgun Gothic"/></font>'
           '<font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Malgun Gothic"/></font>'
           '</fonts><fills count="2"><fill><patternFill patternType="none"/></fill>'
           '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E63"/><bgColor indexed="64"/></patternFill></fill></fills>'
           '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
           '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
           '<cellXfs count="3">'
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
           '<xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>'
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
           '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
           '</styleSheet>')
    return xml.encode("utf-8")


def export_csv_to_xlsx(csv_path: str | Path) -> str:
    """Export CSV as a *new* XLSX with '테스트 결과' and '비고' sheets.

    Old report column names are normalized during export. The source CSV is
    not modified. This uses only the Python standard library to run on Windows.
    """
    src = Path(csv_path)
    with src.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        original_headers = list(reader.fieldnames or [])
        if not original_headers:
            raise ValueError("CSV에 헤더가 없습니다.")
        headers = [LEGACY_NAMES.get(h, h) for h in original_headers]
        if len(set(headers)) != len(headers):
            raise ValueError("CSV에 새 컬럼명과 이전 컬럼명이 중복되어 있습니다.")
        records = list(reader)
    results = [headers] + [[row.get(h, "") for h in original_headers] for row in records]
    notes = [list(x) for x in NOTES_ROWS]
    widths = [19, 12, 22, 28, 50, 14, 22, 58, 58, 28, 16, 16, 20, 45, 25, 55, 32]
    if len(headers) > len(widths):
        widths.extend([25] * (len(headers)-len(widths)))
    dest = src.with_suffix(".xlsx")
    rels = ET.Element(f"{{{P}}}Relationships")
    ET.SubElement(rels, f"{{{P}}}Relationship", {
        "Id":"rId1", "Type":R+"/officeDocument", "Target":"xl/workbook.xml"})
    wb = ET.Element(f"{{{S}}}workbook")
    sheets = ET.SubElement(wb, f"{{{S}}}sheets")
    for idx, name in enumerate(["테스트 결과", "비고"], start=1):
        ET.SubElement(sheets, f"{{{S}}}sheet", {"name":name,"sheetId":str(idx), f"{{{R}}}id":f"rId{idx}"})
    wb_rels = ET.Element(f"{{{P}}}Relationships")
    for idx in (1, 2):
        ET.SubElement(wb_rels, f"{{{P}}}Relationship", {
            "Id":f"rId{idx}", "Type":R+"/worksheet", "Target":f"worksheets/sheet{idx}.xml"})
    ET.SubElement(wb_rels, f"{{{P}}}Relationship", {
        "Id":"rId3", "Type":R+"/styles", "Target":"styles.xml"})
    types = ET.Element(f"{{{C}}}Types")
    ET.SubElement(types, f"{{{C}}}Default", {"Extension":"rels", "ContentType":"application/vnd.openxmlformats-package.relationships+xml"})
    ET.SubElement(types, f"{{{C}}}Default", {"Extension":"xml", "ContentType":"application/xml"})
    for name, content_type in [
        ("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
        ("/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
        ("/xl/worksheets/sheet2.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
        ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
    ]:
        ET.SubElement(types, f"{{{C}}}Override", {"PartName": name, "ContentType": content_type})
    with ZipFile(dest, "w", compression=ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", _xml_bytes(types))
        package.writestr("_rels/.rels", _xml_bytes(rels))
        package.writestr("xl/workbook.xml", _xml_bytes(wb))
        package.writestr("xl/_rels/workbook.xml.rels", _xml_bytes(wb_rels))
        package.writestr("xl/styles.xml", _styles_xml())
        package.writestr("xl/worksheets/sheet1.xml", _sheet_xml(results, widths[:len(headers)], result_sheet=True))
        package.writestr("xl/worksheets/sheet2.xml", _sheet_xml(notes, [22, 40, 77, 91], result_sheet=False))
    return str(dest)
