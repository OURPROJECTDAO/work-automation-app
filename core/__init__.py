"""core 패키지 초기화.

방탄 가드: pandas 3.0 + pyarrow arrow-backed 문자열(read_csv)이 Cloud 최신 네이티브
stack(pandas3.0+pyarrow25+numpy2.x)에서 SIGSEGV(`string_arrow._from_sequence`) →
문자열 저장소를 python 백엔드로 고정해 크래시 경로를 원천 회피한다.
core 를 import 하는 모든 페이지에서 read_csv 전에 실행되므로 실행 순서에 무관.
(streamlit_app.py 엔트리에도 동일 가드. pitfalls 2026-07-13.)
"""
try:
    import pandas as _pd
    _pd.set_option("mode.string_storage", "python")
except Exception:
    pass
