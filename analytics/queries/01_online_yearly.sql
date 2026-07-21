-- 온라인 연도별 매출·이익·활동 채널 수 (대시보드 연간 KPI 재현)
SELECT 연도,
       ROUND(SUM(판매금액)/1e8, 1) AS 매출_억,
       ROUND(SUM(판매이익)/1e8, 2) AS 이익_억,
       ROUND(SUM(판매이익)/SUM(판매금액)*100, 2) AS 이익률_pct,
       COUNT(DISTINCT 상호명) AS 활동채널수
FROM online_sales
GROUP BY 연도 ORDER BY 연도;
