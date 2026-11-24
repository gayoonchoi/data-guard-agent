# -*- coding: utf-8 -*-
"""DataGuard Agent v0.9.8 - 구문규칙 원본 시트 보존.

목표
  - 검사결과서의 'N. 구축기업 구문규칙' / 'N. 검사 구문규칙' 시트를
    Microsoft Excel 자체 Worksheet.Copy로 보완결과서에 가져온다.
  - 1~2행: *보이는* 제목/표 헤더 디자인을 유지한다 (테마 색상 차이 보정).
  - 3행 이후: 원본의 모든 값/수식은 다시 쓰거나 가공하지 않는다.
  - 종합현황/오류목록 등 나머지 시트는 변경하지 않는다.
  - 스타일의 문자 단위 ColumnWidth 차이나 사소한 포맷 속성 차이로
    정상 생성이 중단되지 않도록 한다.

사용: 기존 report_generator.py가 이 파일의
      restore_original_rule_sheets(inspection_path, report_path)를 호출한다.
Windows + Microsoft Excel + pip install pywin32 필요.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import argparse
import os
import re
import shutil
import uuid
import zipfile
import xml.etree.ElementTree as ET
import posixpath

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_RULE_NAME = re.compile(r"^\d+\.\s*(?:구축기업|검사)\s*구문규칙$")
_XL_NONE = -4142


def _sheet_xml_members(xlsx: Path):
    """읽기 전용: 검사결과서의 규칙 시트 이름과 실제 데이터 범위를 얻는다."""
    with zipfile.ZipFile(xlsx) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rel = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.attrib["Id"]: r.attrib["Target"] for r in rel}
        found = []
        for sh in wb.find(_MAIN + "sheets"):
            name = sh.attrib["name"]
            if not _RULE_NAME.fullmatch(name):
                continue
            target = targets[sh.attrib[_REL + "id"]]
            member = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            root = ET.fromstring(z.read(member))
            last_row = 2
            last_col = 1
            for c in root.iter(_MAIN + "c"):
                # Values/formulas only, excluding cells with mere formatting.
                if c.find(_MAIN + "v") is None and c.find(_MAIN + "is") is None and c.find(_MAIN + "f") is None:
                    continue
                address = c.get("r", "")
                m = re.fullmatch(r"([A-Z]+)(\d+)", address)
                if m is None:
                    continue
                row = int(m.group(2))
                col = 0
                for ch in m.group(1):
                    col = 26 * col + ord(ch) - 64
                last_row = max(last_row, row)
                last_col = max(last_col, col)
            found.append((name, last_row, last_col))
    if not found:
        raise ValueError("검사결과서에서 'N. 구축기업 구문규칙' 등의 시트를 찾지 못했습니다.")
    return found


def _actual_color(color):
    """Excel COM의 Color 값을 RGB 정수로 해석 가능한 경우에만 반환."""
    if isinstance(color, bool) or not isinstance(color, (int, float)):
        return None
    num = int(color)
    return num if 0 <= num <= 0xFFFFFF else None


def _restore_header_appearance(original, copied, name, max_col):
    """1~2행 제목 디자인만 이식; 3행 이하 구문규칙은 수정하지 않는다.

    병합 셀은 그 병합 영역을 한 번만 다룬다. Excel COM에서 병합 영역의
    일부 셀을 직접 수정하는 경우 발생하는 오류도 예방한다.
    """
    handled = set()
    for r in (1, 2):
        for c in range(1, max(15, max_col) + 1):
            src_cell = original.Cells(r, c)
            dst_cell = copied.Cells(r, c)
            src_merged = bool(src_cell.MergeCells)
            source_range = src_cell.MergeArea if src_merged else src_cell
            target_range = dst_cell.MergeArea if bool(dst_cell.MergeCells) else dst_cell
            address = str(source_range.Address)
            if address in handled:
                continue
            handled.add(address)
            # 병합범위 내부에서는 대표 셀의 표시 스타일을 사용한다.
            source_style_cell = source_range.Cells(1, 1)
            src_interior = source_style_cell.Interior
            fill_pattern = src_interior.Pattern
            if fill_pattern in (_XL_NONE, None, 0):
                target_range.Interior.Pattern = _XL_NONE
            else:
                rgb = _actual_color(src_interior.Color)
                if rgb is not None:
                    target_range.Interior.Color = rgb
            sf, df = source_style_cell.Font, target_range.Font
            font_rgb = _actual_color(sf.Color)
            if font_rgb is not None:
                df.Color = font_rgb
            df.Bold = sf.Bold
            if sf.Name:
                df.Name = sf.Name
            if sf.Size:
                df.Size = sf.Size
            # 테두리·병합·행높이·열너비는 Excel 자체 Worksheet.Copy의 원본을 유지한다.
    print(f"[헤더 디자인 복원] {name}: 1~2행")

def _as_matrix(value):
    # Excel COM: 단일 셀은 scalar, 여러 셀은 tuple(tuple(...),...).
    if not isinstance(value, tuple):
        return ((value,),)
    if value and not isinstance(value[0], tuple):
        return (tuple(value),)
    return tuple(tuple(row) for row in value)


def _check_rule_values(original, copied, name, last_row, last_col):
    """3행부터 값 일치를 확인; 서식/열 너비 차이는 오류로 보지 않는다."""
    if last_row < 3:
        return
    src = original.Range(original.Cells(3, 1), original.Cells(last_row, last_col))
    dst = copied.Range(copied.Cells(3, 1), copied.Cells(last_row, last_col))
    values_src, values_dst = _as_matrix(src.Value2), _as_matrix(dst.Value2)
    if values_src != values_dst:
        for row_idx, (sr, dr) in enumerate(zip(values_src, values_dst), start=3):
            for col_idx, (a, b) in enumerate(zip(sr, dr), start=1):
                if a != b:
                    raise RuntimeError(
                        f"{name} 규칙 내용 불일치: 행 {row_idx}, 열 {col_idx} / 원본={a!r}, 생성본={b!r}"
                    )
        raise RuntimeError(f"{name} 규칙 내용 크기가 달라졌습니다.")
    # 식 자체가 있는 규칙 셀까지 보존한다.
    # Formula는 일반 문자열 셀에서 Value와 동일하게 동작한다.
    if _as_matrix(src.Formula) != _as_matrix(dst.Formula):
        raise RuntimeError(f"{name} 규칙 수식이 달라졌습니다.")
    print(f"[규칙 내용 검증] {name}: 3~{last_row}행 일치")


def restore_original_rule_sheets(inspection_file: str | Path,
                                 report_file: str | Path) -> Path:
    """보완결과서에서 구문규칙 시트만 원본 검사결과서로 교체한다.

    기존 report_generator.py의 호출 시그니처와 동일하다.
    작업은 임시 파일에서 처리하고 정상 완료 시 교체한다.
    원본 보고서가 Excel에서 열려 있으면 새 디렉터리에 완성본을 만든다.
    """
    src_path = Path(inspection_file).expanduser().resolve()
    dst_path = Path(report_file).expanduser().resolve()
    if not src_path.is_file() or not dst_path.is_file():
        raise FileNotFoundError("검사결과서 또는 생성된 보완결과서가 없습니다.")
    if src_path == dst_path:
        raise ValueError("검사결과서와 보완결과서는 다른 파일이어야 합니다.")
    if src_path.suffix.lower() != ".xlsx" or dst_path.suffix.lower() != ".xlsx":
        raise ValueError(".xlsx 파일만 처리합니다.")

    expected = _sheet_xml_members(src_path)
    present_names = {name for name, _, _ in _sheet_xml_members(dst_path)}
    missing = [name for name, _, _ in expected if name not in present_names]
    if missing:
        raise ValueError(f"보완결과서에 필요한 구문규칙 시트가 없습니다: {missing}")

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("Windows Excel과 pywin32가 필요합니다: python -m pip install pywin32") from exc

    temp_path = dst_path.parent / f".__dg_rules_{uuid.uuid4().hex}.xlsx"
    shutil.copy2(dst_path, temp_path)
    excel = source_book = target_book = None
    excel_completed = False
    pythoncom.CoInitialize()  # Gradio 워커 스레드에서도 실행 가능
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False
        try:
            excel.AutomationSecurity = 3
        except Exception:
            pass
        source_book = excel.Workbooks.Open(str(src_path), UpdateLinks=0, ReadOnly=True,
                                          IgnoreReadOnlyRecommended=True)
        target_book = excel.Workbooks.Open(str(temp_path), UpdateLinks=0, ReadOnly=False,
                                          IgnoreReadOnlyRecommended=True)
        if target_book.ReadOnly:
            raise PermissionError("임시 보완결과서를 읽기 전용으로 열었습니다.")
        original_count = target_book.Worksheets.Count
        for name, last_row, last_col in expected:
            original = source_book.Worksheets(name)
            obsolete = target_book.Worksheets(name)
            # Excel 자체 시트 복사: 병합, 원본 규칙값, 열 너비, 행 높이를 유지.
            original.Copy(Before=obsolete)
            copied = target_book.Worksheets(obsolete.Index - 1)
            obsolete.Delete()
            copied.Name = name
            try:
                copied.Cells.ClearComments()
            except Exception:
                pass
            # 제목 디자인에 필요한 부분만 보정하고, 아래 규칙값은 건드리지 않는다.
            _restore_header_appearance(original, copied, name, last_col)
            _check_rule_values(original, copied, name, last_row, last_col)

        if target_book.Worksheets.Count != original_count:
            raise RuntimeError("구문규칙 시트 교체 후 전체 시트 개수가 달라졌습니다.")
        target_book.Save()
        target_book.Close(SaveChanges=False)
        target_book = None
        source_book.Close(SaveChanges=False)
        source_book = None
        excel_completed = True
    finally:
        if target_book is not None:
            try:
                target_book.Close(SaveChanges=False)
            except Exception:
                pass
        if source_book is not None:
            try:
                source_book.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        if not excel_completed and temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass

    # 저장 이후에도 구문규칙 시트가 존재하는지 구조 확인 (수치 서식 강제 검증 X).
    try:
        final_names = {name for name, _, _ in _sheet_xml_members(temp_path)}
        if not {name for name, _, _ in expected}.issubset(final_names):
            raise RuntimeError("저장된 보고서에 구문규칙 시트가 누락됐습니다.")
        backup_dir = dst_path.parent / "_backups"
        backup_dir.mkdir(exist_ok=True)
        backup = backup_dir / f"{dst_path.stem}_교체전_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsx"
        shutil.copy2(dst_path, backup)
        try:
            os.replace(temp_path, dst_path)
            return dst_path
        except PermissionError:
            alternative = dst_path.parent / ("재생성_" + datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6])
            alternative.mkdir(parents=True, exist_ok=False)
            alt_path = alternative / dst_path.name
            shutil.move(str(temp_path), str(alt_path))
            return alt_path
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass


def main():
    parser = argparse.ArgumentParser(description="검사결과서의 구문규칙 시트를 보완결과서에 원본 복사")
    parser.add_argument("--inspection", required=True, help="원본 검사결과서.xlsx")
    parser.add_argument("--report", required=True, help="생성된 보완결과서.xlsx")
    args = parser.parse_args()
    print(restore_original_rule_sheets(args.inspection, args.report))


if __name__ == "__main__":
    main()
