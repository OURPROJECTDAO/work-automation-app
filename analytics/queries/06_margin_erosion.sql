-- 마진 침식 탐지: 최근 3개월 vs 직전 3개월 마진율 하락 상품 (조건부 집계 + HAVING)
WITH base AS (
  SELECT 관리코드, any_value(상품명) AS 상품명,
    SUM(CASE WHEN 거래일자 >= current_date - INTERVAL 3 MONTH THEN 판매금액 END) AS 매출_최근,
    SUM(CASE WHEN 거래일자 >= current_date - INTERVAL 3 MONTH THEN 판매이익 END) AS 이익_최근,
    SUM(CASE WHEN 거래일자 <  current_date - INTERVAL 3 MONTH
              AND 거래일자 >= current_date - INTERVAL 6 MONTH THEN 판매금액 END) AS 매출_직전,
    SUM(CASE WHEN 거래일자 <  current_date - INTERVAL 3 MONTH
              AND 거래일자 >= current_date - INTERVAL 6 MONTH THEN 판매이익 END) AS 이익_직전
  FROM online_sales
  WHERE 관리코드 IS NOT NULL AND 관리코드 <> '00-12'
  GROUP BY 관리코드
  HAVING 매출_최근 > 3e6 AND 매출_직전 > 3e6
)
SELECT 관리코드, 상품명,
       ROUND(이익_직전/매출_직전*100,2) AS 마진율_직전3M,
       ROUND(이익_최근/매출_최근*100,2) AS 마진율_최근3M,
       ROUND(이익_최근/매출_최근*100 - 이익_직전/매출_직전*100, 2) AS 변화_pp
FROM base
WHERE 이익_최근/매출_최근 < 이익_직전/매출_직전
ORDER BY 변화_pp ASC LIMIT 25;
