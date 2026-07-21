-- 채널(거래처)×연도 매출 + 연도 내 채널 순위 (RANK 윈도우)
SELECT 연도, 상호명 AS 채널,
       ROUND(SUM(판매금액)/1e8,2) AS 매출_억,
       RANK() OVER (PARTITION BY 연도 ORDER BY SUM(판매금액) DESC) AS 연도내순위,
       ROUND(SUM(판매금액)/SUM(SUM(판매금액)) OVER (PARTITION BY 연도)*100,1) AS 비중_pct
FROM online_sales
GROUP BY 연도, 상호명
QUALIFY 연도내순위 <= 8
ORDER BY 연도, 연도내순위;
