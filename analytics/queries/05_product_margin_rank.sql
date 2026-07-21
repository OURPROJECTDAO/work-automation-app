-- 최근 12개월 상품(관리코드)별 매출·이익·마진율 + 매출 순위·누적 기여율 (윈도우 조합)
WITH p AS (
  SELECT 관리코드, any_value(상품명) AS 상품명,
         SUM(판매금액) AS 매출, SUM(판매이익) AS 이익
  FROM online_sales
  WHERE 거래일자 >= current_date - INTERVAL 12 MONTH
    AND 관리코드 IS NOT NULL AND 관리코드 <> '00-12'
  GROUP BY 관리코드
)
SELECT 관리코드, 상품명,
       ROUND(매출/1e6,1) AS 매출_백만,
       ROUND(이익/매출*100,2) AS 마진율_pct,
       RANK() OVER (ORDER BY 매출 DESC) AS 매출순위,
       ROUND(SUM(매출) OVER (ORDER BY 매출 DESC) / SUM(매출) OVER () * 100, 1) AS 누적기여_pct
FROM p
QUALIFY 매출순위 <= 30
ORDER BY 매출순위;
