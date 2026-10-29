"""共有の医用画像ライブラリ（~/src/medical-image-library）を参照する。

厳密な解剖図が必要なスライドは、AIに描かせず検証済みのストック図を差し込む。
その図をこのモジュールで検索・解決する。ライブラリ本体は別リポジトリで管理し、
このプロジェクトには画像をコピーしない（同じ図を各所に増やさないため）。
"""

from __future__ import annotations

import json
import os
from functools import lru_cache
from pathlib import Path

LIBRARY_DIR = Path(
    os.environ.get("MEDICAL_IMAGE_LIBRARY", Path.home() / "src/medical-image-library")
)

CATALOG_PATH = LIBRARY_DIR / "catalog.json"
WEB_DIR = LIBRARY_DIR / "web"


def available() -> bool:
    return CATALOG_PATH.exists()


@lru_cache(maxsize=1)
def _load(mtime: float) -> list[dict]:
    data = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    return data.get("images", [])


def catalog() -> list[dict]:
    """カタログを読む。ライブラリ側が更新されたら自動で読み直す。"""
    if not available():
        return []
    return _load(CATALOG_PATH.stat().st_mtime)


def resolve(rel_path: str) -> Path | None:
    """カタログ上の相対パス（例 anatomy/shoulder-xxx.jpg）を実ファイルに解決する。"""
    if not rel_path:
        return None
    # ライブラリの外を指す指定を弾く
    p = (WEB_DIR / rel_path).resolve()
    try:
        p.relative_to(WEB_DIR.resolve())
    except ValueError:
        return None
    return p if p.exists() else None


def search(
    query: str = "",
    *,
    region: str | None = None,
    category: str | None = None,
    text_in_image: str | None = None,
    limit: int = 60,
) -> list[dict]:
    """キーワード（空白区切りのAND）と属性で絞り込む。"""
    terms = [t for t in (query or "").split() if t]
    hits = []
    for e in catalog():
        if region and e.get("region") != region:
            continue
        if category and e.get("category") != category:
            continue
        if text_in_image and e.get("text_in_image") != text_in_image:
            continue
        if terms:
            haystack = " ".join(
                [e.get("title", ""), e.get("description", ""), e.get("path", ""), *e.get("tags", [])]
            )
            if not all(t in haystack for t in terms):
                continue
        hits.append(e)
        if len(hits) >= limit:
            break
    return hits


def entry(rel_path: str) -> dict | None:
    for e in catalog():
        if e.get("path") == rel_path:
            return e
    return None


def regions() -> list[dict]:
    """UI のフィルタ用に、部位と件数の一覧を返す。"""
    counts: dict[str, int] = {}
    for e in catalog():
        counts[e.get("region", "none")] = counts.get(e.get("region", "none"), 0) + 1
    labels = {
        "shoulder": "肩", "elbow": "肘", "wrist-hand": "手", "neck": "頸部",
        "back": "背部", "low-back": "腰部", "hip": "股関節", "knee": "膝",
        "ankle-foot": "足", "head-face": "頭頸部", "chest": "胸部",
        "abdomen": "腹部", "whole-body": "全身", "systemic": "全身性", "none": "その他",
    }
    order = list(labels.keys())
    return [
        {"value": r, "label": labels.get(r, r), "count": counts[r]}
        for r in order
        if counts.get(r)
    ]
