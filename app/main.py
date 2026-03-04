"""
月次受講人数集計 Web アプリ
FastAPI + htmx
ポータル経由: BEHIND_PORTAL=true + X-Portal-Role ヘッダーで認証スキップ
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from fastapi import FastAPI, File, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from services.aggregator import (
    aggregate,
    build_pivot,
    load_excel,
    parse_target_month,
    save_monthly_result,
    to_excel_bytes,
)

RESULTS_DIR = PROJECT_ROOT / "outputs" / "results"
RESULTS_DIR.mkdir(parents=True, exist_ok=True)

templates = Jinja2Templates(directory=str(PROJECT_ROOT / "templates"))

app = FastAPI(title="月次受講人数集計")

# グローバルキャッシュ（シングルユーザー前提）
_last_pivot: object = None


@app.middleware("http")
async def portal_auth(request: Request, call_next):
    if os.environ.get("BEHIND_PORTAL") == "true" and request.headers.get("X-Portal-Role"):
        return await call_next(request)
    # ポータル外からのアクセスはそのまま通す（Caddyのforward_authに委任）
    return await call_next(request)


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    pivot = build_pivot(RESULTS_DIR)
    months = [c for c in pivot.columns if c not in ["学年", "教室", "講座名", "M/C", "担当"]] if not pivot.empty else []
    return templates.TemplateResponse("index.html", {
        "request": request,
        "has_data": not pivot.empty,
        "months": months,
        "total_rows": len(pivot) if not pivot.empty else 0,
    })


@app.post("/upload", response_class=HTMLResponse)
async def upload(request: Request, file: UploadFile = File(...)):
    global _last_pivot

    contents = await file.read()
    MAX_UPLOAD_SIZE = 20 * 1024 * 1024  # 20MB
    if len(contents) > MAX_UPLOAD_SIZE:
        return templates.TemplateResponse("result.html", {
            "request": request,
            "error": "ファイルサイズが上限（20MB）を超えています",
        })
    filename = file.filename or ""

    target_month = parse_target_month(filename)
    if target_month is None:
        return templates.TemplateResponse("result.html", {
            "request": request,
            "error": f"ファイル名からターゲット月を判定できません: {filename}（例: *_2504.xlsx）",
        })

    try:
        df = load_excel(contents)
        result = aggregate(df, target_month)
    except Exception as e:
        return templates.TemplateResponse("result.html", {
            "request": request,
            "error": f"集計エラー: {e}",
        })

    if result is None or result.empty:
        return templates.TemplateResponse("result.html", {
            "request": request,
            "error": f"{target_month}: 集計対象データがありませんでした",
        })

    save_monthly_result(result, target_month, RESULTS_DIR)

    pivot = build_pivot(RESULTS_DIR)
    _last_pivot = pivot

    months = [c for c in pivot.columns if c not in ["学年", "教室", "講座名", "M/C", "担当"]]
    return templates.TemplateResponse("result.html", {
        "request": request,
        "month": str(target_month),
        "rows": len(result),
        "months": months,
        "total_rows": len(pivot),
    })


@app.get("/download")
async def download():
    global _last_pivot
    if _last_pivot is None:
        _last_pivot = build_pivot(RESULTS_DIR)
    if _last_pivot is None or (hasattr(_last_pivot, "empty") and _last_pivot.empty):
        return Response("データがありません", status_code=404)
    excel_bytes = to_excel_bytes(_last_pivot)
    return Response(
        content=excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=monthly_stats.xlsx"},
    )
