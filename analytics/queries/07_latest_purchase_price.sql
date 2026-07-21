-- 관리코드별 최신 매입 단가 (ROW_NUMBER 윈도우로 최신 1건 선택)
SELECT 관리코드, 상품명, 기준일 AS 최근매입일, 박스단가, 단가 AS 낱개단가
FROM (
  SELECT *, ROW_NUMBER() OVER (PARTITION BY 관리코드 ORDER BY 기준일 DESC) AS rn
  FROM purchases WHERE 관리코드 IS NOT NULL
)
WHERE rn = 1
ORDER BY 최근매입일 DESC LIMIT 50;
