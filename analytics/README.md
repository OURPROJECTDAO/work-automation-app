# analytics — DuckDB SQL 분석 층

work-automation-data 의 parquet(매출 42파티션·매입 54파티션·거래처그룹)를 **파일 그대로** SQL 뷰로 얹어,
대시보드 핵심 지표를 표준 SQL로 재현·임시분석하는 층. (pandas 파이프라인의 SQL 대응물)

## 사용
```bash
pip install duckdb
python run_sql.py --list            # 쿼리 목록
python run_sql.py 01_online_yearly  # 단일 실행
python run_sql.py all               # 전체
# 데이터 루트: --data <path> 또는 WA_DATA (기본 ../data, 하위 master/ purchases/ groups/)
```

## 뷰 (duck_views.sql)
- `sales` / `purchases` : read_parquet 글롭 + 연도·연월 파생
- `store_groups`, `online_sales` : 거래처 그룹 조인(온라인 필터)

## 쿼리 세트
| 쿼리 | 내용 | SQL 요소 |
|---|---|---|
| 01_online_yearly | 연도별 매출·이익·이익률·활동채널 | GROUP BY 집계 |
| 02_h1_yoy | 상반기 YoY 성장률 | LAG 윈도우 |
| 03_monthly_trend | 월별 추이·전년동월·3M 이동평균 | LAG(12)·프레임 윈도우 |
| 04_channel_yearly | 채널×연도 매출·순위·비중 | RANK·PARTITION·QUALIFY |
| 05_product_margin_rank | 12M 상품 매출순위·누적기여율 | RANK·누적 SUM OVER |
| 06_margin_erosion | 최근3M vs 직전3M 마진율 하락 | 조건부 집계·HAVING |
| 07_latest_purchase_price | 관리코드별 최신 매입단가 | ROW_NUMBER 최신 1건 |
| 08_price_vs_cost | 판매단가×최신매입가 마진 검수 | CTE 조인 |

## 검증
2026-07-21 골든 대조 — pandas 파이프라인과 **원 단위 일치** 확인:
2025 온라인 매출 11,103,105,175 · 2026 H1 6,106,986,665 (양쪽 동일).
