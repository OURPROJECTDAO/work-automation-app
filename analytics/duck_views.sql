-- ============================================================
-- duck_views.sql — work-automation-data parquet 위 DuckDB 뷰 층
-- {DATA} 는 run_sql.py 가 데이터 루트 경로로 치환 (기본: data/)
-- ============================================================
CREATE OR REPLACE VIEW sales AS
SELECT 거래일자, 상호명, 관리코드, 상품명, 규격, 상품분류, 수량, 판매금액, 판매이익,
       CAST(strftime(거래일자,'%Y') AS INT) AS 연도,
       strftime(거래일자,'%Y-%m')          AS 연월
FROM read_parquet('{DATA}/master/sales_*.parquet');

CREATE OR REPLACE VIEW purchases AS
SELECT 기준일, 관리코드, 상품명, 규격, 박스내품, 박스, 수량, 단가, 박스단가,
       공급가액, 부가세, 할인, 합계액, 대분류, 중분류, 소분류,
       strftime(기준일,'%Y-%m') AS 연월
FROM read_parquet('{DATA}/purchases/buyin_*.parquet');

CREATE OR REPLACE VIEW store_groups AS
SELECT 상호명, 그룹 FROM read_csv_auto('{DATA}/groups/store_groups.csv');

-- 온라인 그룹 매출만 (채널 분석 기본 뷰)
CREATE OR REPLACE VIEW online_sales AS
SELECT s.* FROM sales s JOIN store_groups g USING(상호명)
WHERE g.그룹 = '온라인';
