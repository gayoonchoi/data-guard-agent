# DataGuard Local Correction Agent — 실험 모듈 v0.1

## 역할
- **④ 실제 데이터 연결**에서 생성한 `*_evidence_draft_*.csv`를 읽어, 구문 규칙과 검증 상태에 근거한 **보완 후 수정 제안 문장만** 로컬 Ollama로 작성합니다.
- 수정된 **새 CSV**를 반환하므로 기존 CSV/원본 라벨링데이터/Excel은 변경하지 않습니다.
- 출력 CSV는 기존 **⑤ 보완 전후 Excel 반영** 탭에서 **직접 선택하여 업로드**할 수 있습니다.
- 이 모듈은 *검증·자동 수정·승인* 기능이 아닙니다. 모든 제안은 담당자 검토 필요입니다.

## 설치 (Windows CMD)
1. 회사 PC에 승인된 경우에만 [Ollama](https://ollama.com/download/windows) 설치
2. `ollama pull qwen3:4b`
3. `ollama list`에서 모델 확인
4. `local_correction_agent.py`를 프로젝트의 `1. RuleEngine` 폴더에 복사
5. 다음을 CMD에서 실행합니다. (경로는 본인 CSV에 맞게 수정)

```cmd
cd /d "C:\Users\USER\Desktop\최가윤\data-guard-agent"
python "1. RuleEngine\local_correction_agent.py" --input "outputs\step4_evidence\YOUR_EVIDENCE_DRAFT.csv" --limit 10
```

출력: 같은 폴더에 `YOUR_EVIDENCE_DRAFT_ollama_review.csv`가 생성됩니다. 이를 ⑤의 *검토 CSV* 입력으로 선택합니다.

다른 모델: `--model gemma3:4b` (먼저 `ollama pull gemma3:4b`).
필요한 경우에 한해 `--include-before`를 붙여 원본값 일부를 로컬 모델에 전달할 수 있습니다. **기본값은 원본값 미전송**입니다.

## 운영 주의
- 오직 `127.0.0.1:11434`로만 요청합니다. Cloud 모델이나 외부 API는 사용하지 않습니다.
- 019/1 `imagePath`의 미해결 규칙 충돌은 건너뜁니다. 다른 충돌 규칙은 운영 전에 별도 차단 목록/승인 연동이 필요합니다.
- `--limit 10`으로 먼저 평가하고, 모델 결과 정확도를 확인한 후 범위를 늘리세요.
- 동일 규칙/상태/항목은 실행 내에서 캐시하여 중복 LLM 호출을 줄입니다.
- 회사 데이터·규칙 사전과 생성 CSV는 GitHub에 올리지 마세요.
- 규칙과 보완 전 근거를 바탕으로 한 **문장 제안**만 수행하며, 검증 판정과 승인 책임은 사람/Rule Engine에 남깁니다.
