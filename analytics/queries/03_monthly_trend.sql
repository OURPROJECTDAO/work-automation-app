-- 월별 매출 추이 + 전년동월 대비 (LAG 12) + 3개월 이동평균 (프레임 윈도우)
WITH m AS (
  SELECT 연월, SUM(판매금액) AS 매출 FROM online_sales GROUP BY 연월
)
SELECT 연월,
       ROUND(매출/1e8,2) AS 매출_억,
       ROUND(LAG(매출,12) OVER (ORDER BY 연월)/1e8,2) AS 전년동월_억,
       ROUND((매출/NULLIF(LAG(매출,12) OVER (ORDER BY 연월),0)-1)*100,1) AS YoY_pct,
       ROUND(AVG(매출) OVER (ORDER BY 연월 ROWS BETWEEN 2 PRECEDING AND CURRENT ROW)/1e8,2) AS 이동평균3M_억
FROM m ORDER BY 연월;
