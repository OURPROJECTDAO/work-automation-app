"""데일리 대시보드 파일 자동 인계 — 세션 인박스(휘발성·in-memory).

이전 워크플로우(파일처리: 오픈마켓 송장출력·천년경영 output)가 산출물을 이 인박스에 push →
데일리 대시보드가 재업로드 없이 바로 사용. PII(송장 수령자/주소/송장번호) 때문에 디스크/repo 미저장
→ 세션 한정(리부트/새 세션 시 비워짐 → 수동 업로드). 상품관리(master)는 reference 라이브라 인박스 불요.

단일 진실원천(슬롯 키): 생산처(1_파일처리)와 소비처(데일리 대시보드)가 이 상수를 공유해 키 드리프트 방지.
"""
INBOX_KEY = "daily_inbox"
SLOT_CHEONNYEON = "천년경영"
SLOT_INVOICE = "송장출력"
SLOTS = (SLOT_CHEONNYEON, SLOT_INVOICE)


def push(session_state, slot: str, data, name: str, ts: str) -> None:
    """세션 인박스 슬롯에 파일 바이트 적재(덮어씀). data=bytes."""
    box = session_state.setdefault(INBOX_KEY, {})
    box[slot] = {"bytes": bytes(data), "name": name, "ts": ts}


def get(session_state, slot: str):
    """슬롯 내용 {bytes,name,ts} 또는 None."""
    return session_state.get(INBOX_KEY, {}).get(slot)
