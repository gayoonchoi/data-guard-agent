"""Standalone DataGuard Case-RAG A/B test interface (no production code changes)."""
from __future__ import annotations

import csv
import datetime as dt
from contextlib import closing
from pathlib import Path

import gradio as gr

from case_rag_engine import (
    ROOT, DB_FILE, bootstrap, connect, load_tests, list_cases, status_counts,
    check_ollama, register_verified_case, run_comparison, save_csv,
)
from excel_export import export_csv_to_xlsx

EVALUATIONS = ["적절", "부적절", "추가 검토 필요"]


def init_kb():
    with closing(connect()) as c:
        bootstrap(c)


def db_summary():
    with closing(connect()) as c:
        counts = status_counts(c)
        data = [[x["case_id"], x["source_kind"], x["project_id"], x["error_type"],
                 x["target_path"], "예" if x["source_verified"] and x["human_approved"] and not x["is_demo"] else "아니요"]
                for x in list_cases(c)]
    return (" | ".join(f"{k}: {v}" for k, v in counts.items()), data)


def load_custom_cases(upload):
    try:
        cases = load_tests(Path(upload)) if upload else load_tests()
        return f"시험 사례 {len(cases)}건 준비됨", cases, gr.update(
            choices=[x["case_id"] for x in cases], value=cases[0]["case_id"])
    except Exception as exc:
        return f"CSV 로딩 실패: {exc}", gr.update(), gr.update()


def case_details(case_id, cases):
    row = next((x for x in cases if x["case_id"] == case_id), None)
    if not row:
        return "선택된 사례가 없습니다.", ""
    # Ground Truth is shown in this box only, never in the prompt.
    prompt = "\n".join([
        "파일명: " + row["file_name"], "오류: " + row["reported_error"],
        "오류 경로: " + row["property"], "적용 규칙: " + row["rule"],
        "확인된 사실(가상 입력이면 가상): " + row["observed_evidence"],
    ])
    return prompt, row["expected_after"]


def run_selected(case_id, cases, model, demo_mode):
    query = next((x for x in cases if x["case_id"] == case_id), None)
    if query is None:
        return "사례를 선택하세요.", [], "", "", "", "", None, None, ""
    with closing(connect()) as c:
        result = run_comparison(c, query, model.strip(), demo=bool(demo_mode))
    row = result["record"]
    csv_path = save_csv(row)
    excel_path = export_csv_to_xlsx(csv_path)
    matches = [[x["case_id"], x["source_kind"], x["error_type"], x["target_path"],
                x["relevance_score"], x["approved_after"]] for x in result["matches"]]
    status = (f"A 실제 호출: {'성공' if result['baseline'] else '실패'}"
              f" · B 실제 호출: {'성공' if result['rag'] else '미호출/실패'}"
              f" · 검색 사례 {len(matches)}건"
              f" · 모드: {'가상 데모 (업무 근거 아님)' if demo_mode else '실제 승인 사례만'}")
    if not matches:
        status += "\n※ RAG용 승인 사례가 없으므로 B 모델 호출을 생략했어. A 결과만 평가해 줘."
    if row["A · Ollama 실행 오류"]:
        status += "\nA 오류: " + row["A · Ollama 실행 오류"]
    if row["B · Ollama 실행 오류"] and matches:
        status += "\nB 오류: " + row["B · Ollama 실행 오류"]
    return (status, matches,
            row["A · Ollama 제안(사례 검색 전)"], row["A · Ollama 근거"],
            row["B · Ollama 제안(Case RAG)"], row["B · Ollama 근거"],
            str(csv_path), excel_path, str(csv_path))


def register_case(*args):
    project, fname, etype, prop, observed, rule, after, evidence, reviewer, c1, c2, c3 = args
    payload = {
        "project_id": project, "filename": fname, "error_type": etype, "target_path": prop,
        "reported_error": observed, "rule_text": rule, "approved_after": after,
        "evidence_reference": evidence, "reviewer": reviewer,
    }
    try:
        with closing(connect()) as c:
            case_id = register_verified_case(c, payload, {
                "source_checked": c1, "rule_checked": c2, "after_approved": c3,
            })
        return f"담당자 확인·승인 사례 등록 완료: {case_id}\n※ 체크한 증빙은 사용자 확인에 따른 것이며 시스템이 실제 라벨링 파일을 자동 검사한 것은 아니야."
    except Exception as exc:
        return "등록 중단: " + str(exc)


def review_result(src, a_eval, b_eval, reason_a, reason_b):
    try:
        if not src:
            raise ValueError("먼저 A/B 시험을 실행하세요.")
        if a_eval not in EVALUATIONS or b_eval not in EVALUATIONS:
            raise ValueError("평가값은 적절/부적절/추가 검토 필요 중 하나여야 합니다.")
        if (a_eval != "추가 검토 필요" and not str(reason_a).strip()) or (b_eval != "추가 검토 필요" and not str(reason_b).strip()):
            raise ValueError("적절·부적절 판정에는 평가 근거를 기입하세요.")
        with Path(src).open(encoding="utf-8-sig", newline="") as f:
            records = list(csv.DictReader(f))
            if len(records) != 1:
                raise ValueError("A/B 실행 기록 1건만 평가할 수 있습니다.")
            record = records[0]
        if record.get("A · Ollama 실행 오류") and a_eval != "추가 검토 필요":
            raise ValueError("A의 모델 호출이 실패했으므로 정확성을 확정할 수 없습니다.")
        if record.get("B · Ollama 실행 오류") and b_eval != "추가 검토 필요":
            raise ValueError("B의 모델 호출이 없거나 실패했으므로 정확성을 확정할 수 없습니다.")
        record["A · 수정안 정확성 평가"] = a_eval
        record["B · 수정안 정확성 평가"] = b_eval
        record["A · 평가 근거"] = str(reason_a).strip()
        record["B · 평가 근거"] = str(reason_b).strip()
        record["평가 주체"] = "담당자 수동 평가"
        reviewed = Path(src).with_name(Path(src).stem + "_담당자평가_" + dt.datetime.now().strftime("%H%M%S%f") + ".csv")
        with reviewed.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=list(record))
            writer.writeheader(); writer.writerow(record)
        return f"A={a_eval}, B={b_eval} · 평가 별도 저장 완료 (원본 유지)", str(reviewed), export_csv_to_xlsx(reviewed)
    except Exception as exc:
        return "평가 저장 실패: " + str(exc), None, None


def main(preview_only=False):
    init_kb()
    cases = load_tests()
    with gr.Blocks(title="DataGuard · Case-Based RAG 독립 A/B 실험") as page:
        gr.Markdown("# DataGuard · Case-Based RAG 독립 실험 v0.1\n**DataGuard 본체와 분리**된 도구야. 기존 ①~⑤ 코드, 검사결과서 및 라벨링 원본을 수정하지 않아. 모든 모델 호출은 PC의 `127.0.0.1:11434` 로컬 Ollama만 사용해.")
        with gr.Tab("① Case RAG 전후 성능 비교"):
            gr.Markdown("테스트 사례의 **모범답안은 모델에 보내지 않으며**, 결과 비교·담당자 평가에만 사용해. **기본 Case KB에는 실제 승인 사례가 0건**이라 최초에는 가상 데모를 선택해서 검색 흐름을 실험할 수 있어.")
            with gr.Row():
                model = gr.Textbox(label="Ollama 모델", value="qwen2.5:7b-instruct")
                test_server = gr.Button("로컬 Ollama 모델 확인")
            server_status = gr.Textbox(label="연결 상태", interactive=False)
            test_server.click(check_ollama, inputs=[model], outputs=[server_status])
            upload = gr.File(label="다른 테스트 사례 CSV 업로드 (선택)", type="filepath", file_types=[".csv"])
            reload_button = gr.Button("테스트 CSV 불러오기")
            csv_status = gr.Textbox(label="테스트 CSV 불러오기 결과", interactive=False)
            state = gr.State(cases)
            choose = gr.Dropdown(label="테스트 사례", choices=[x["case_id"] for x in cases], value="S02")
            reload_button.click(load_custom_cases, inputs=[upload], outputs=[csv_status, state, choose])
            detail_btn = gr.Button("오류·규칙 입력과 모범답안 확인")
            input_data = gr.Textbox(label="모델에게 전달할 현재 오류 근거", lines=6, interactive=False)
            ground_truth = gr.Textbox(label="모범답안_비교용(Ground Truth) · 모델에 미전달", lines=3, interactive=False)
            detail_btn.click(case_details, inputs=[choose, state], outputs=[input_data, ground_truth])
            demo_mode = gr.Checkbox(label="가상 데모 유사 사례로 검색 실험 (실제 승인 사례 아님)", value=True)
            run_btn = gr.Button("현재 사례로 A/B 로컬 Ollama 실행", variant="primary")
            status = gr.Textbox(label="호출 및 검색 결과", lines=3, interactive=False)
            retrieval_table = gr.Dataframe(headers=["사례 ID", "사례 유형", "오류 유형", "과거 경로", "검색 관련도", "사례 수정 방향"], label="검색된 유사 사례", interactive=False)
            with gr.Row():
                with gr.Column():
                    baseline = gr.Textbox(label="A · Ollama 보완 후 제안 (사례 검색 전)", lines=4, interactive=False)
                    baseline_reason = gr.Textbox(label="A · Ollama 근거", lines=4, interactive=False)
                with gr.Column():
                    rag = gr.Textbox(label="B · Ollama 보완 후 제안 (Case RAG)", lines=4, interactive=False)
                    rag_reason = gr.Textbox(label="B · Ollama 근거", lines=4, interactive=False)
            csv_output = gr.File(label="A/B 테스트 CSV 다운로드")
            xlsx_output = gr.File(label="A/B 테스트 Excel 다운로드 (비고 시트 포함)")
            latest = gr.State("")
            run_btn.click(run_selected, inputs=[choose, state, model, demo_mode],
                          outputs=[status, retrieval_table, baseline, baseline_reason, rag, rag_reason,
                                   csv_output, xlsx_output, latest])
            gr.Markdown("### 담당자 정확성 평가\n자동 키워드 검사나 검색 관련도 점수는 정확도 점수가 아니야. 결과를 실제 구문규칙과 대조해 판단한 뒤 별도 저장해.")
            with gr.Row():
                eval_a = gr.Dropdown(label="A · 수정안 정확성 평가", choices=EVALUATIONS, value="추가 검토 필요")
                eval_b = gr.Dropdown(label="B · 수정안 정확성 평가", choices=EVALUATIONS, value="추가 검토 필요")
            with gr.Row():
                reason_a = gr.Textbox(label="A · 평가 근거", lines=3)
                reason_b = gr.Textbox(label="B · 평가 근거", lines=3)
            eval_btn = gr.Button("담당자 A/B 평가 저장")
            eval_status = gr.Textbox(label="평가 저장 결과", interactive=False)
            review_csv = gr.File(label="평가 완료 CSV")
            review_xlsx = gr.File(label="평가 완료 Excel")
            eval_btn.click(review_result, inputs=[latest, eval_a, eval_b, reason_a, reason_b],
                           outputs=[eval_status, review_csv, review_xlsx])

        with gr.Tab("② 실제 승인 사례 등록"):
            gr.Markdown("**중요:** 이 탭은 시스템이 실데이터를 자동 검증하는 기능이 아니야. 검사결과서·원본 데이터·보완결과서를 실제로 대조한 **담당자가 증빙을 기록하고 승인한 사례만** 실사례 RAG 검색에 등록해. 회사 데이터 정책에 따라 로컬에서만 활용해.")
            with gr.Row():
                project = gr.Textbox(label="사업 ID", placeholder="예: 019")
                filename = gr.Textbox(label="대상 파일명 (선택)")
            with gr.Row():
                error_type = gr.Textbox(label="오류 유형", placeholder="검사결과서 기록 그대로")
                property_path = gr.Textbox(label="검사 프로퍼티 / 정확한 계층 경로")
            observed = gr.Textbox(label="검사 당시 실제 오류 상태 (원본 확인 후 기입)", lines=3)
            rule = gr.Textbox(label="해당 사업의 승인된 구문규칙 원문 또는 검증된 규칙 요약", lines=4)
            after = gr.Textbox(label="담당자가 승인한 '보완 후 수정 지시문'", lines=3)
            evidence = gr.Textbox(label="검증 증빙 참조 (원본 파일명/규칙 시트·행/승인 문서 번호 등)", lines=3)
            reviewer = gr.Textbox(label="검증·승인 담당자 식별값")
            c1 = gr.Checkbox(label="실제 라벨링 파일에서 검사 당시 오류와 대상 경로를 대조함", value=False)
            c2 = gr.Checkbox(label="현재 적용·승인된 사업별 구문규칙과 제안 문장이 일치함을 검토함", value=False)
            c3 = gr.Checkbox(label="담당자가 해당 보완 후 수정 지시문을 최종 승인함", value=False)
            register_btn = gr.Button("검증·승인 사례를 Case KB에 등록", variant="primary")
            registered = gr.Textbox(label="등록 결과", lines=3, interactive=False)
            register_btn.click(register_case,
                inputs=[project, filename, error_type, property_path, observed, rule, after,
                        evidence, reviewer, c1, c2, c3], outputs=[registered])

        with gr.Tab("③ Case KB 현황"):
            gr.Markdown("기본 등록 S01~S06은 **가상 평가 데이터**라서 실사례 검색에 사용하지 않아. `DEMO-TOPLEVEL-KEY`도 데모 모드에서만 사용해. 실제 승인된 사례만 업무 모드에서 검색해.")
            refresh = gr.Button("Case KB 현황 새로고침")
            db_status = gr.Textbox(label="KB 상태", interactive=False)
            db_table = gr.Dataframe(headers=["사례 ID", "출처", "사업 ID", "오류 유형", "경로", "실사례 검색 가능"], interactive=False, label="Case KB 목록")
            refresh.click(db_summary, outputs=[db_status, db_table])
            gr.Markdown("**개인정보·사업 데이터 주의:** `lab_data`에는 승인 사례·테스트 로그가 로컬 저장돼. 회사 승인 없이 외부 GitHub나 클라우드에 올리지 마.")
    if preview_only:
        return page
    page.queue().launch(server_name="127.0.0.1", server_port=7862, share=False, inbrowser=True)


if __name__ == "__main__":
    main()
