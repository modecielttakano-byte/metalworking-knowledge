#!/usr/bin/env python3
"""Local Japanese-first search UI for metalworking process knowledge."""

from __future__ import annotations

import json
import re
import sqlite3
from pathlib import Path

from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from ingest import DEFAULT_DB, to_ngrams

ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
DB_PATH = DEFAULT_DB

SOURCE_FILTERS = {
    "internal_log": "社内ログ",
    "internal_ts_tool": "社内TS・工具",
    "external_qa": "外部Q&A",
}

FTS_UNSAFE = re.compile(r'["^:*()\\]')
STOP_OPS = {"AND", "OR", "NOT", "NEAR"}

app = FastAPI(title="加工方法ナレッジ検索", docs_url=None, redoc_url=None)
app.mount("/static", StaticFiles(directory=str(STATIC)), name="static")


def db() -> sqlite3.Connection:
    if not DB_PATH.exists():
        raise HTTPException(
            status_code=503,
            detail="インデックスがありません。先に python ingest.py を実行してください。",
        )
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def fts_query(q: str) -> str | None:
    q = (q or "").strip()
    if not q:
        return None
    terms = [t for t in re.split(r"\s+", q) if t]
    groups: list[str] = []
    for term in terms:
        term = FTS_UNSAFE.sub(" ", term).strip()
        if not term or term.upper() in STOP_OPS:
            continue
        ngrams = to_ngrams(term).split()
        if not ngrams:
            continue
        # Prefer 2-grams when the term is 2+ chars; keep 1-gram only for short queries.
        bigrams = [g for g in ngrams if len(g) >= 2]
        use = bigrams or ngrams
        # Limit tokens so very long pastes stay usable
        use = use[:12]
        quoted = []
        for g in use:
            if g.upper() in STOP_OPS:
                continue
            quoted.append('"' + g.replace('"', "") + '"')
        if quoted:
            groups.append("(" + " AND ".join(quoted) + ")")
    if not groups:
        return None
    return " AND ".join(groups)


def snippet(body: str, query: str, width: int = 140) -> str:
    text = re.sub(r"\s+", " ", body or "").strip()
    if not text:
        return ""
    q = (query or "").strip()
    hit_at = -1
    if q:
        for term in re.split(r"\s+", q):
            if not term:
                continue
            hit_at = text.lower().find(term.lower())
            if hit_at >= 0:
                break
    if hit_at < 0:
        return text[:width] + ("…" if len(text) > width else "")
    start = max(0, hit_at - 40)
    end = min(len(text), start + width)
    prefix = "…" if start > 0 else ""
    suffix = "…" if end < len(text) else ""
    return prefix + text[start:end] + suffix


def rank_row(row: sqlite3.Row, query: str) -> float:
    score = 0.0
    src = row["source_type"]
    title = row["title"] or ""
    body = row["body"] or ""
    q = (query or "").strip()
    if src == "internal_log":
        score += 90
        extra = {}
        try:
            extra = json.loads(row["extra_json"] or "{}")
        except json.JSONDecodeError:
            extra = {}
        if extra.get("has_countermeasure"):
            score += 25
        if row["has_result"]:
            score += 35
    elif src == "internal_ts_tool":
        score += 75
        if row["has_result"]:
            score += 15
    else:
        score += 18
        if row["has_answer"]:
            score += 45
        else:
            score -= 8
    if q:
        for term in re.split(r"\s+", q):
            if not term:
                continue
            tl = term.lower()
            if tl in title.lower():
                score += 28
            elif tl in body.lower():
                score += 8
    return score


@app.get("/")
def index() -> FileResponse:
    return FileResponse(STATIC / "index.html")


@app.get("/api/health")
def health() -> dict:
    ready = DB_PATH.exists()
    payload: dict = {"ok": ready, "db": str(DB_PATH.relative_to(ROOT))}
    if not ready:
        payload["hint"] = "python ingest.py で data/internal.xlsx と data/qa.xlsx を取り込んでください"
        return payload
    conn = db()
    try:
        row = conn.execute("SELECT value FROM meta WHERE key='stats'").fetchone()
        payload["stats"] = json.loads(row["value"]) if row else {}
        built = conn.execute("SELECT value FROM meta WHERE key='built_at'").fetchone()
        payload["built_at"] = built["value"] if built else None
    finally:
        conn.close()
    return payload


@app.get("/api/facets")
def facets() -> dict:
    conn = db()
    try:
        sources = [
            dict(r)
            for r in conn.execute(
                "SELECT source_type, source_label, COUNT(*) AS n FROM documents GROUP BY source_type"
            )
        ]
        processes = [
            dict(r)
            for r in conn.execute(
                """
                SELECT process, COUNT(*) AS n FROM documents
                WHERE process IS NOT NULL AND process != ''
                GROUP BY process ORDER BY n DESC, process
                """
            )
        ]
        return {"sources": sources, "processes": processes, "source_filters": SOURCE_FILTERS}
    finally:
        conn.close()


@app.get("/api/search")
def search(
    q: str = Query("", description="日本語キーワード"),
    source: str = Query("", description="comma-separated source_type"),
    process: str = Query("", description="工程フィルタ"),
    answered_only: bool = Query(False, description="外部Q&Aは回答ありのみ"),
    result_only: bool = Query(False, description="社内ログは結果ありのみ"),
    limit: int = Query(30, ge=1, le=100),
    offset: int = Query(0, ge=0),
) -> dict:
    conn = db()
    try:
        sources = [s for s in source.split(",") if s in SOURCE_FILTERS]
        match = fts_query(q)
        where = ["1=1"]
        params: list = []
        if match:
            where.append("documents.id IN (SELECT rowid FROM documents_fts WHERE documents_fts MATCH ?)")
            params.append(match)
        if sources:
            where.append(f"source_type IN ({','.join('?' * len(sources))})")
            params.extend(sources)
        if process:
            where.append("process = ?")
            params.append(process)
        if answered_only:
            where.append("(source_type != 'external_qa' OR has_answer = 1)")
        if result_only:
            where.append("(source_type != 'internal_log' OR has_result = 1)")
        where_sql = " AND ".join(where)
        order_sql = """
            CASE source_type
                WHEN 'internal_log' THEN 0
                WHEN 'internal_ts_tool' THEN 1
                ELSE 2 END,
            has_result DESC,
            has_answer DESC,
            id
        """
        try:
            total = conn.execute(
                f"SELECT COUNT(*) FROM documents WHERE {where_sql}", params
            ).fetchone()[0]
            if match:
                rows = conn.execute(
                    f"SELECT * FROM documents WHERE {where_sql}", params
                ).fetchall()
                ranked = sorted(rows, key=lambda r: rank_row(r, q), reverse=True)
                page = ranked[offset : offset + limit]
            else:
                page = conn.execute(
                    f"SELECT * FROM documents WHERE {where_sql} ORDER BY {order_sql} LIMIT ? OFFSET ?",
                    (*params, limit, offset),
                ).fetchall()
        except sqlite3.OperationalError:
            like = f"%{q.strip()}%"
            like_where = "(title LIKE ? OR body LIKE ?)"
            like_params: list = [like, like]
            if sources:
                like_where += f" AND source_type IN ({','.join('?' * len(sources))})"
                like_params.extend(sources)
            if process:
                like_where += " AND process = ?"
                like_params.append(process)
            total = conn.execute(
                f"SELECT COUNT(*) FROM documents WHERE {like_where}", like_params
            ).fetchone()[0]
            rows = conn.execute(
                f"SELECT * FROM documents WHERE {like_where}", like_params
            ).fetchall()
            ranked = sorted(rows, key=lambda r: rank_row(r, q), reverse=True)
            page = ranked[offset : offset + limit]
        hits = []
        for r in page:
            extra = {}
            try:
                extra = json.loads(r["extra_json"] or "{}")
            except json.JSONDecodeError:
                extra = {}
            hits.append(
                {
                    "id": r["id"],
                    "title": r["title"],
                    "snippet": snippet(r["body"], q),
                    "source_type": r["source_type"],
                    "source_label": r["source_label"],
                    "trust_label": r["trust_label"],
                    "process": r["process"],
                    "model": r["model"],
                    "machine": r["machine"],
                    "factory": r["factory"],
                    "tool": r["tool"],
                    "url": r["url"],
                    "has_result": bool(r["has_result"]),
                    "has_answer": bool(r["has_answer"]),
                    "has_countermeasure": bool(extra.get("has_countermeasure")),
                    "score": rank_row(r, q),
                }
            )
        return {"q": q, "total": total, "offset": offset, "limit": limit, "hits": hits}
    finally:
        conn.close()


@app.get("/api/docs/{doc_id}")
def doc_detail(doc_id: int) -> dict:
    conn = db()
    try:
        r = conn.execute("SELECT * FROM documents WHERE id = ?", (doc_id,)).fetchone()
        if not r:
            raise HTTPException(status_code=404, detail="見つかりません")
        extra = {}
        try:
            extra = json.loads(r["extra_json"] or "{}")
        except json.JSONDecodeError:
            extra = {}
        return {
            "id": r["id"],
            "title": r["title"],
            "body": r["body"],
            "source_type": r["source_type"],
            "source_label": r["source_label"],
            "trust_label": r["trust_label"],
            "process": r["process"],
            "sheet_name": r["sheet_name"],
            "model": r["model"],
            "machine": r["machine"],
            "factory": r["factory"],
            "tool": r["tool"],
            "url": r["url"],
            "has_result": bool(r["has_result"]),
            "has_answer": bool(r["has_answer"]),
            "extra": extra,
        }
    finally:
        conn.close()


def main() -> None:
    import uvicorn

    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=False)


if __name__ == "__main__":
    main()
