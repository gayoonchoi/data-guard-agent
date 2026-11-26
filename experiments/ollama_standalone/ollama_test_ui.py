"""DataGuard independent local Ollama test UI.

NOT an integration patch.  Does not read/write the DataGuard project or dataset files.
Only sends explicitly entered test-case text to the local Ollama HTTP API.
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import re
import time
import urllib.error
import urllib.request
from pathlib import Path

import gradio as gr
from excel_export import (GROUND_TRUTH_COL, KEYWORD_COL, RUN_ERROR_COL,
                          LEGACY_NAMES, export_csv_to_xlsx)

ROOT = Path(__file__).resolve().parent
CASES_FILE = ROOT / "sample_cases.csv"
LOG_DIR = ROOT / "test_results"
API_ROOT = "http://127.0.0.1:11434"
FIELDS = ["case_id", "classification", "file_name", "error_type", "property", "reported_error",
          "observed_evidence", "rule", "similar_case", "expected_after", "keywords"]
SYSTEM = (
    "당신은 AI 학습데이터의 구문 검증 보완 제안 작성자입니다. "
    "반드시 한국어로만 답하고, 검사결과서/실제값/현재 적용 규칙/승인된 사례의 근거를 구별합니다. "
    "실제 파일을 읽었다거나 실제 수정·삭제를 완료했다고 주장하지 않습니다. "
    "파일 부재만으로 삭제 완료라고 단정하지 않습니다. "
    "규칙만으로 판단할 수 없으면 추측해 임의 값을 만들지 말고 구체적으로 추가 확인할 자료를 말합니다. "
    "경로 계층(예: $.text와 $.shapes[*].text)은 정확히 구분합니다. "
    "수정 제안은 한국어 한 문장으로, 근거도 한국어로 작성하세요. "
    "주어진 사실 외 사업별 고정 답안을 사용하지 마세요."
)
EVAL_CHOICES = ["적절", "부적절", "추가 검토 필요"]
EVAL_COL = "수정안 정확성 평가"
EVAL_REASON_COL = "평가 근거"
EVAL_REVIEWER_COL = "평가 주체"

FORMAT = {
    "type": "object", "properties": {
        "suggestion": {"type": "string"},
        "reason": {"type": "string"},
        "human_review_required": {"type": "boolean"},
        "missing_evidence": {"type": "string"},
    },
    "required": ["suggestion", "reason", "human_review_required", "missing_evidence"],
}


def load_cases(path: str | Path = CASES_FILE):
    with open(path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        missing = [x for x in FIELDS if x not in (reader.fieldnames or [])]
        if missing:
            raise ValueError("CSV 필수 열 누락: " + ", ".join(missing))
        cases = [{k: (row.get(k) or "") for k in FIELDS} for row in reader]
    if not cases:
        raise ValueError("테스트 사례가 없습니다.")
    ids = [c["case_id"] for c in cases]
    if len(set(ids)) != len(ids):
        raise ValueError("case_id에 중복이 있습니다.")
    return cases


def display_rows(cases):
    return [[r["case_id"], r["classification"], r["file_name"],
             r["property"], r["expected_after"]] for r in cases]


def get_case(cases, case_id):
    return next((r for r in cases if r["case_id"] == case_id), None)


def payload_text(case):
    return "\n".join([
        "[검사결과서에 기록된 오류]",
        "파일명: " + case["file_name"],
        "오류 유형: " + case["error_type"],
        "검사 프로퍼티: " + case["property"],
        "검사 당시 오류 기록: " + case["reported_error"],
        "[독립적으로 확인된 사실 / 테스트가정은 가정임]",
        case["observed_evidence"] or "제공되지 않음",
        "[해당 사업의 현재 적용 구문규칙]",
        case["rule"] or "확인되지 않음",
        "[검색된 과거 승인 사례: 없으면 없음]",
        case["similar_case"] or "없음",
        "[요청] 위 근거에 따라 '보완 후'에 기록할 수정 제안을 작성하세요. "
        "규칙과 실제값이 불충분하면 보완값을 지어내지 말고 확인 필요 사항을 제안하세요. "
        "최종 보완 완료 여부를 단정하지 마세요.",
    ])


def request_json(url, data=None, timeout=8):
    headers = {"Content-Type": "application/json"}
    req = urllib.request.Request(url, headers=headers,
                                 data=None if data is None else json.dumps(data, ensure_ascii=False).encode("utf-8"))
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def check_server(model):
    try:
        body = request_json(API_ROOT + "/api/tags", timeout=5)
        installed = [m.get("name", "") for m in body.get("models", [])]
        if model not in installed:
            return "서버 연결 성공 · 모델 없음: " + model + "\n설치된 모델: " + ", ".join(installed)
        return "연결 성공 · 설치된 모델 확인: " + model + " · 127.0.0.1:11434"
    except Exception as exc:
        return "연결 실패 · " + str(exc) + "\nOllama 실행 상태를 확인하세요."


def call_ollama(case, model):
    started = time.perf_counter()
    response = request_json(API_ROOT + "/api/chat", {
        "model": model,
        "stream": False,
        "think": False,
        "format": FORMAT,
        "messages": [
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": payload_text(case)},
        ],
        "options": {"temperature": 0, "num_predict": 400},
    }, timeout=180)
    elapsed = round(time.perf_counter() - started, 2)
    raw = response.get("message", {}).get("content", "")
    answer = json.loads(raw)
    if not isinstance(answer, dict) or not isinstance(answer.get("suggestion"), str):
        raise ValueError("Ollama 응답이 요구한 JSON 구조와 다릅니다: " + raw[:200])
    return {
        "model_response": response.get("model", model),
        "suggestion": answer.get("suggestion", ""),
        "reason": answer.get("reason", ""),
        "human_review_required": answer.get("human_review_required", ""),
        "missing_evidence": answer.get("missing_evidence", ""),
        "elapsed_sec": elapsed,
        "prompt_tokens": response.get("prompt_eval_count", ""),
        "answer_tokens": response.get("eval_count", ""),
        "raw": raw,
    }


def compare_keywords(case, suggestion):
    keywords = [x.strip() for x in case.get("keywords", "").split(";") if x.strip()]
    missing = [x for x in keywords if x not in suggestion]
    korean = bool(re.search(r"[가-힣]", suggestion))
    return ("자동 키워드 확인: " + ("충족" if not missing else "미충족 (" + ", ".join(missing) + ")") +
            " · 한국어: " + ("포함" if korean else "없음") +
            " · 의미 정확성은 담당자 검토 필요")


def record_row(case, response=None, error=""):
    response = response or {}
    return {
        "실행시각": dt.datetime.now().isoformat(timespec="seconds"),
        "case_id": case["case_id"],
        "파일명": case["file_name"],
        "검사_프로퍼티": case["property"],
        GROUND_TRUTH_COL: case["expected_after"],
        "모델호출상태": "호출성공" if response else "실패",
        "실제응답모델": response.get("model_response", ""),
        "Ollama 제안": response.get("suggestion", ""),
        "Ollama 근거": response.get("reason", ""),
        "추가확인자료": response.get("missing_evidence", ""),
        "사람검토필요": response.get("human_review_required", ""),
        "소요시간_초": response.get("elapsed_sec", ""),
        EVAL_COL: "추가 검토 필요",
        EVAL_REASON_COL: "담당자 검토 전" if response else "모델 호출 실패: 제안 없음",
        EVAL_REVIEWER_COL: "미평가",
        KEYWORD_COL: compare_keywords(case, response.get("suggestion", "")) if response else "",
        RUN_ERROR_COL: error,
    }


def save_results(records):
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    fname = "ollama_test_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".csv"
    path = LOG_DIR / fname
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(records[0]))
        writer.writeheader()
        writer.writerows(records)
    return str(path)


def do_one(case_id, cases, model):
    case = get_case(cases, case_id)
    if case is None:
        return "사례를 선택하세요.", "", "", "", "", None
    try:
        result = call_ollama(case, model)
        path = save_results([record_row(case, result)])
        info = f"로컬 API 실제 호출 성공 · 응답 모델 {result['model_response']} · {result['elapsed_sec']}초 · 생성 토큰 {result['answer_tokens']}"
        score = compare_keywords(case, result["suggestion"])
        return info, result["suggestion"], result["reason"] + "\n추가 확인: " + str(result["missing_evidence"]), case["expected_after"], score, path
    except Exception as exc:
        message = f"실제 호출 실패: {type(exc).__name__}: {exc}"
        path = save_results([record_row(case, error=message)])
        return message, "", "", case["expected_after"], "호출 미성공 · 모델 생성 결과 없음", path


def do_batch(cases, model, limit):
    records = []
    for case in cases[:max(1, min(int(limit), len(cases)))]:
        try:
            records.append(record_row(case, call_ollama(case, model)))
        except Exception as exc:
            records.append(record_row(case, error=f"{type(exc).__name__}: {exc}"))
    path = save_results(records)
    success = sum(row["모델호출상태"] == "호출성공" for row in records)
    preview = [[x["case_id"], x["모델호출상태"], x["Ollama 제안"], x[GROUND_TRUTH_COL], x[EVAL_COL], x[KEYWORD_COL]] for x in records]
    return f"총 {len(records)}건 · 호출 성공 {success}건 · 실패 {len(records)-success}건 · DataGuard 원본 변경 없음", preview, path


def read_review_records(path):
    """Read a generated CSV and keep its original data/column order."""
    if not path:
        raise ValueError("먼저 테스트를 실행하거나 기존 결과 CSV를 불러오세요.")
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        old_headers = list(reader.fieldnames or [])
        headers = [LEGACY_NAMES.get(x, x) for x in old_headers]
        if len(headers) != len(set(headers)):
            raise ValueError("결과 CSV에서 기존/신규 컬럼명이 중복되었습니다.")
        if "case_id" not in headers or "Ollama 제안" not in headers:
            raise ValueError("이 CSV에는 case_id 또는 Ollama 제안 열이 없습니다.")
        records = [{LEGACY_NAMES.get(k, k): v for k, v in row.items()} for row in reader]
    if not records:
        raise ValueError("CSV에 평가할 결과가 없습니다.")
    for col in [EVAL_COL, EVAL_REASON_COL, EVAL_REVIEWER_COL]:
        if col not in headers:
            headers.append(col)
    for row in records:
        if not row.get(EVAL_COL):
            row[EVAL_COL] = "추가 검토 필요"
        row.setdefault(EVAL_REASON_COL, "담당자 검토 전")
        row.setdefault(EVAL_REVIEWER_COL, "미평가")
    return headers, records


def review_select_source(path):
    """Load latest test or uploaded previous CSV for human evaluation."""
    try:
        _, records = read_review_records(path)
        choices = [r["case_id"] for r in records]
        return (f"평가 대상 {len(records)}건 불러옴 · 담당자 확인 후 저장하세요.",
                str(path), gr.update(choices=choices, value=choices[0]))
    except Exception as exc:
        return f"평가 CSV 불러오기 실패: {exc}", "", gr.update()


def review_case_detail(path, case_id):
    try:
        _, records = read_review_records(path)
        row = next(r for r in records if r["case_id"] == case_id)
        summary = ("Ollama 제안: " + row.get("Ollama 제안", "") + "\n"
                   "Ollama 근거: " + row.get("Ollama 근거", "") + "\n"
                   "모범답안: " + row.get(GROUND_TRUTH_COL, "") + "\n"
                   "※ 키워드 충족 여부는 정확성 판정 근거가 아닙니다.")
        return summary, row.get(EVAL_COL, "추가 검토 필요"), row.get(EVAL_REASON_COL, "")
    except (StopIteration, Exception) as exc:
        return "사례를 선택하거나 결과 CSV를 불러오세요: " + str(exc), "추가 검토 필요", ""


def save_human_review(path, case_id, evaluation, reason):
    """Save a separately reviewed CSV. Never rewrite the original model output."""
    try:
        if evaluation not in EVAL_CHOICES:
            raise ValueError("평가값은 적절/부적절/추가 검토 필요 중 하나여야 합니다.")
        headers, records = read_review_records(path)
        row = next((r for r in records if r["case_id"] == case_id), None)
        if row is None:
            raise ValueError("결과 CSV에서 사례를 찾지 못했습니다.")
        if evaluation in ("적절", "부적절") and not reason.strip():
            raise ValueError("'적절/부적절'을 선택했다면 구문규칙에 근거한 평가 이유를 입력하세요.")
        if row.get("모델호출상태") != "호출성공" and evaluation != "추가 검토 필요":
            raise ValueError("모델이 응답하지 않은 사례는 정확성을 평가할 수 없습니다.")
        row[EVAL_COL] = evaluation
        row[EVAL_REASON_COL] = reason.strip() or "추가 근거 확인 필요"
        row[EVAL_REVIEWER_COL] = "담당자 수동 평가"
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        destination = LOG_DIR / (Path(path).stem + "_평가완료_" +
                                 dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".csv")
        with destination.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader()
            writer.writerows(records)
        export_csv_to_xlsx(destination)
        return (f"{case_id} 평가 저장: {evaluation} · 원본 CSV 유지 · 새 결과 {destination.name}",
                str(destination), str(destination))
    except Exception as exc:
        return "평가 저장 실패: " + str(exc), None, path


def do_one_with_review(case_id, cases, model):
    base = do_one(case_id, cases, model)
    path = base[-1]
    if not path:
        return (*base, None, "", gr.update())
    return (*base, export_csv_to_xlsx(path), str(path), gr.update(choices=[case_id], value=case_id))


def do_batch_with_review(cases, model, limit):
    base = do_batch(cases, model, limit)
    path = base[-1]
    _, rows = read_review_records(path)
    ids = [r["case_id"] for r in rows]
    return (*base, export_csv_to_xlsx(path), str(path), gr.update(choices=ids, value=ids[0]))


def save_human_review_with_excel(path, case_id, evaluation, reason):
    status, csv_out, review_path = save_human_review(path, case_id, evaluation, reason)
    if csv_out is None:
        return status, None, None, review_path
    return status, csv_out, export_csv_to_xlsx(csv_out), review_path


def convert_existing_csv_to_excel(path):
    """Convert an existing test CSV without calling Ollama or modifying it."""
    try:
        if not path:
            raise ValueError("먼저 결과 CSV를 불러오세요.")
        output = export_csv_to_xlsx(path)
        return "비고 시트 포함 Excel 생성 완료 · 모델 재호출 없음", output
    except Exception as exc:
        return "Excel 생성 실패: " + str(exc), None


def upload_cases(upload):
    if not upload:
        cases = load_cases()
        status = "기본 예시로 복원했습니다."
    else:
        try:
            cases = load_cases(upload)
            status = f"CSV {len(cases)}건을 읽었습니다."
        except Exception as exc:
            return status_message(str(exc)), gr.update(), gr.update(), gr.update()
    return status, cases, display_rows(cases), gr.update(choices=[c["case_id"] for c in cases], value=cases[0]["case_id"])


def status_message(msg):
    return "CSV 불러오기 실패: " + msg


def show_selected(case_id, cases):
    c = get_case(cases, case_id)
    if not c:
        return "", "", "", ""
    return payload_text(c), c["expected_after"], c["classification"], "테스트 사례의 모범답안은 Ollama 요청에 전달되지 않습니다."


def main(preview_only=False):
    cases = load_cases()
    with gr.Blocks(title="DataGuard · Ollama 독립 성능 테스트") as demo:
        gr.Markdown("# DataGuard · Ollama 독립 성능 테스트 v0.1.3\n기존 DataGuard 프로그램과 **완전히 분리**되어 있으며, 현재 예시 CSV에 있는 텍스트만 로컬 Ollama로 전달합니다. 원본 Excel과 라벨링데이터는 수정하지 않습니다.")
        state = gr.State(cases)
        with gr.Row():
            model = gr.Textbox(value="qwen2.5:7b-instruct", label="모델명")
            server_button = gr.Button("로컬 Ollama 연결 확인")
        server_status = gr.Textbox(label="연결 상태", interactive=False)
        server_button.click(check_server, inputs=[model], outputs=[server_status])
        gr.Markdown("## 테스트 사례 (가상 테스트 데이터 · 일부 실제 규칙을 바탕으로 구성)")
        source_upload = gr.File(label="다른 사례 CSV 업로드 (선택)", file_types=[".csv"], type="filepath")
        upload_button = gr.Button("사례 CSV 불러오기")
        upload_status = gr.Textbox(label="CSV 로딩 결과", interactive=False)
        preview = gr.Dataframe(headers=["사례 ID", "구분", "파일명", "프로퍼티", "기대 보완 후"],
                               value=display_rows(cases), interactive=False, label="테스트 목록")
        case_id = gr.Dropdown(choices=[c["case_id"] for c in cases], value=cases[0]["case_id"], label="선택 사례")
        upload_button.click(upload_cases, inputs=[source_upload], outputs=[upload_status, state, preview, case_id])
        show_button = gr.Button("입력 근거와 모범답안 확인")
        prompt_preview = gr.Textbox(label="모델에 전달하는 입력 (모범답안 제외)", lines=12, interactive=False)
        benchmark = gr.Textbox(label="모범답안 (비교용 · 모델에 미전달)", lines=3, interactive=False)
        kind = gr.Textbox(label="사례 유형", interactive=False)
        note = gr.Textbox(label="안내", interactive=False)
        show_button.click(show_selected, inputs=[case_id, state], outputs=[prompt_preview, benchmark, kind, note])
        gr.Markdown("## 단일 사례 시험")
        one_button = gr.Button("선택 사례로 Ollama 실제 호출", variant="primary")
        one_status = gr.Textbox(label="실제 호출 결과", interactive=False)
        suggestion = gr.Textbox(label="Ollama 보완 후 제안", lines=3, interactive=False)
        reason = gr.Textbox(label="Ollama 판단 근거 및 부족한 근거", lines=4, interactive=False)
        gold = gr.Textbox(label="모범답안", lines=3, interactive=False)
        score = gr.Textbox(label="자동 점검 (의미 정확성 보장 아님)", interactive=False)
        one_file = gr.File(label="단일 사례 테스트 기록 CSV")
        one_excel = gr.File(label="단일 사례 테스트 Excel (테스트 결과·비고 시트)")
        # Review selection widgets are defined below, after both test sections.
        gr.Markdown("## 여러 사례 비교 시험")
        max_cases = gr.Slider(minimum=1, maximum=200, value=3, step=1, label="이번에 실행할 사례 수 (처음부터 N건)")
        batch_button = gr.Button("여러 사례 일괄 테스트")
        batch_status = gr.Textbox(label="일괄 실행 결과", interactive=False)
        batch_preview = gr.Dataframe(headers=["사례 ID", "실제 모델 호출", "Ollama 제안", "모범답안", "수정안 정확성 평가", "자동 점검"], interactive=False, label="비교 표")
        batch_file = gr.File(label="전체 테스트 결과 CSV")
        batch_excel = gr.File(label="전체 테스트 결과 Excel (테스트 결과·비고 시트)")
        gr.Markdown("## 수정안 정확성 평가 — 담당자 검토")
        gr.Markdown("새 테스트 결과에는 기본적으로 **추가 검토 필요**가 기록됩니다. Ollama 제안과 구문규칙·모범답안(Ground Truth)을 비교해 **적절 / 부적절 / 추가 검토 필요**를 직접 판정하고, 근거를 저장할 수 있습니다. **모델이 스스로 정확하다고 평가하는 기능은 아닙니다.**")
        review_path = gr.State("")
        review_upload = gr.File(label="예전 결과 CSV 불러오기 (선택)", file_types=[".csv"], type="filepath")
        review_upload_button = gr.Button("기존 결과 CSV 평가 대상으로 불러오기")
        review_status = gr.Textbox(label="평가 진행 상태", interactive=False)
        review_id = gr.Dropdown(choices=[], label="평가할 사례 ID")
        review_detail_button = gr.Button("선택 사례의 제안·근거·모범답안 확인")
        review_detail = gr.Textbox(label="평가용 비교 내용", lines=5, interactive=False)
        review_choice = gr.Dropdown(choices=EVAL_CHOICES, value="추가 검토 필요", label="수정안 정확성 평가")
        review_reason = gr.Textbox(label="평가 근거 (적절/부적절 시 필수)", lines=3,
                                   placeholder="예: 최상위 $.text는 규칙에 없는데 추가하라고 제안하여 부적절")
        review_save_button = gr.Button("평가 저장 CSV 만들기", variant="primary")
        review_download = gr.File(label="평가 완료 CSV 다운로드")
        review_excel = gr.File(label="평가 결과 Excel 다운로드 (비고 시트 포함)")
        existing_excel_button = gr.Button("현재 CSV를 비고 시트 포함 Excel로 변환 (Ollama 재호출 없음)")

        one_button.click(do_one_with_review, inputs=[case_id, state, model],
                         outputs=[one_status, suggestion, reason, gold, score, one_file, one_excel, review_path, review_id])
        batch_button.click(do_batch_with_review, inputs=[state, model, max_cases],
                           outputs=[batch_status, batch_preview, batch_file, batch_excel, review_path, review_id])
        review_upload_button.click(review_select_source, inputs=[review_upload],
                                   outputs=[review_status, review_path, review_id])
        review_detail_button.click(review_case_detail, inputs=[review_path, review_id],
                                   outputs=[review_detail, review_choice, review_reason])
        existing_excel_button.click(convert_existing_csv_to_excel, inputs=[review_path],
                                    outputs=[review_status, review_excel])
        review_save_button.click(save_human_review_with_excel,
                                 inputs=[review_path, review_id, review_choice, review_reason],
                                 outputs=[review_status, review_download, review_excel, review_path])
        gr.Markdown("**Excel에는 결과와 별도의 비고 시트를 포함합니다. 평가 주체·Ground Truth·Ollama 실행 오류 등 모든 컬럼의 뜻을 설명합니다.**")
        gr.Markdown("**주의:** 모델 답변은 수정 지시 *초안*이며 실제 데이터 수정·삭제 사실을 인증하지 않습니다. 키워드 검사와 호출 성공 여부는 정확성 점수가 아닙니다. **평가는 담당자가 구문규칙과 근거를 확인해 기록**합니다. 테스트 CSV 및 실행 기록은 GitHub에 올리지 마세요.")
    if preview_only:
        return demo
    demo.queue().launch(server_name="127.0.0.1", share=False, inbrowser=True)


if __name__ == "__main__":
    main()
