-- 상반기(1~6월) 매출 YoY — LAG 윈도우로 전년 대비 성장률
WITH h1 AS (
  SELECT 연도, SUM(판매금액) AS 매출
  FROM online_sales
  WHERE CAST(strftime(거래일자,'%m') AS INT) <= 6
  GROUP BY 연도
)
SELECT 연도,
       ROUND(매출/1e8,1) AS 상반기매출_억,
       ROUND(LAG(매출) OVER (ORDER BY 연도)/1e8,1) AS 전년동기_억,
       ROUND((매출/LAG(매출) OVER (ORDER BY 연도)-1)*100,1) AS YoY_pct
FROM h1 ORDER BY 연도;
