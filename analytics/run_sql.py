#!/usr/bin/env python3
"""DuckDB 분석 러너 — work-automation-data parquet 위 SQL 쿼리 실행.
사용:
  python run_sql.py --list                # 쿼리 목록
  python run_sql.py 01_online_yearly      # 단일 실행
  python run_sql.py all                   # 전체 실행
데이터 루트: --data 또는 환경변수 WA_DATA (기본 ./data, 구조: master/ purchases/ groups/)
"""
import argparse, os, sys, glob
import duckdb

HERE = os.path.dirname(os.path.abspath(__file__))
QDIR = os.path.join(HERE, "queries")

def connect(data_root: str) -> duckdb.DuckDBPyConnection:
    con = duckdb.connect()
    views = open(os.path.join(HERE, "duck_views.sql"), encoding="utf-8").read()
    con.execute(views.replace("{DATA}", data_root.rstrip("/")))
    return con

def query_names():
    return sorted(os.path.splitext(os.path.basename(p))[0]
                  for p in glob.glob(os.path.join(QDIR, "*.sql")))

def run(con, name: str):
    path = os.path.join(QDIR, f"{name}.sql")
    if not os.path.exists(path):
        sys.exit(f"쿼리 없음: {name} (--list 참고)")
    sql = open(path, encoding="utf-8").read()
    df = con.sql(sql).df()
    print(f"\n===== {name} ({len(df)} rows) =====")
    print(df.to_string(index=False))
    return df

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("query", nargs="?", help="쿼리명 또는 all")
    ap.add_argument("--data", default=os.environ.get("WA_DATA", os.path.join(HERE, "..", "data")))
    ap.add_argument("--list", action="store_true")
    a = ap.parse_args()
    if a.list or not a.query:
        print("\n".join(query_names())); return
    con = connect(a.data)
    for name in (query_names() if a.query == "all" else [a.query]):
        run(con, name)

if __name__ == "__main__":
    main()
