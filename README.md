# work-automation-app

오픈마켓 발주 정제 자동화 시스템.

## 구조
- `core/` — 처리 코어 (프레임워크 무관 Python)
- `reference/` — 참조 데이터 csv (도서산간/필터링/미배송)
- `tests/` — 테스트 + 골든 파일
- `app/` — Streamlit 앱

## 실행
```bash
pip install -r requirements.txt
streamlit run app/streamlit_app.py
```

## 개발 원칙
- 처리 로직은 `core/`에만. `app/`은 UI + 호출만.
- 주소 매칭 전 NFC 정규화 필수 (`normalize_kr()`).
- 참조 데이터 csv는 UTF-8-sig(BOM).
- 새 템플릿 = `core/workflows/`에 모듈 1개 추가.
