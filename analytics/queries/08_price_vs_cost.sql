-- 판매 실적 × 최신 매입가 조인: 최근 30일 평균 판매단가 vs 최신 박스단가 (마진 이중검수 축소판)
WITH latest_cost AS (
  SELECT 관리코드, 박스단가
  FROM (SELECT *, ROW_NUMBER() OVER (PARTITION BY 관리코드 ORDER BY 기준일 DESC) rn FROM purchases)
  WHERE rn = 1
), recent_sales AS (
  SELECT 관리코드, any_value(상품명) AS 상품명,
         SUM(판매금액)/NULLIF(SUM(수량),0) AS 평균판매단가, SUM(수량) AS 판매수량
  FROM online_sales
  WHERE 거래일자 >= current_date - INTERVAL 30 DAY
    AND 관리코드 IS NOT NULL AND 관리코드 <> '00-12'
  GROUP BY 관리코드 HAVING SUM(수량) >= 10
)
SELECT s.관리코드, s.상품명, s.판매수량,
       ROUND(s.평균판매단가) AS 평균판매단가,
       c.박스단가,
       ROUND((s.평균판매단가 - c.박스단가)/NULLIF(s.평균판매단가,0)*100,2) AS 단순마진율_pct
FROM recent_sales s JOIN latest_cost c USING(관리코드)
ORDER BY 단순마진율_pct ASC LIMIT 30;
