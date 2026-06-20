from __future__ import annotations

import re
from datetime import datetime

from flask import Blueprint, jsonify, render_template, request, send_file

from . import slides as slides_svc

bp = Blueprint("main", __name__)

OUTPUT_ROOT = slides_svc.OUTPUT_ROOT


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _valid_folder_name(name: str) -> bool:
    if not name or name.startswith("."):
        return False
    if "/" in name or "\\" in name or ".." in name:
        return False
    return True


def _list_projects() -> list[str]:
    if not OUTPUT_ROOT.exists():
        return []
    return sorted(
        d.name
        for d in OUTPUT_ROOT.iterdir()
        if d.is_dir()
        and not d.name.startswith(".")
        and (d / slides_svc.MARKDOWN_FILE).exists()
    )


def _next_project_name() -> str:
    existing = set(_list_projects())
    n = 1
    while f"プロジェクト{n}" in existing:
        n += 1
    return f"プロジェクト{n}"


def _project_status(project_name: str) -> dict:
    plan = slides_svc.load_plan(project_name)
    generated = [slides_svc.slide_png_path(project_name, i).exists() for i in range(len(plan))]
    return {
        "project": project_name,
        "markdown": slides_svc.read_markdown(project_name),
        "prompt": slides_svc.read_prompt(project_name),
        "has_reference": slides_svc.reference_path(project_name).exists(),
        "plan": plan,
        "generated": generated,
        "pptx_exists": slides_svc.pptx_path(project_name).exists(),
    }


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

@bp.route("/")
def index():
    return render_template("index.html")


# ---------------------------------------------------------------------------
# プロジェクト管理（ペイン1）
# ---------------------------------------------------------------------------

@bp.route("/api/projects")
def api_projects():
    if not OUTPUT_ROOT.exists():
        return jsonify({"projects": [], "error": f"出力フォルダがありません: {OUTPUT_ROOT}"}), 200
    return jsonify({"projects": _list_projects()})


@bp.route("/api/prompt_templates")
def api_prompt_templates():
    """新規作成時に選べるプロンプト雛形を返す。"""
    return jsonify({
        "templates": slides_svc.PROMPT_TEMPLATES,
        "default": slides_svc.DEFAULT_PROMPT,
        # デフォルト見本画像を持つテンプレート名（フロントでプレビュー表示に使う）
        "with_reference": slides_svc.templates_with_reference(),
    })


@bp.route("/api/template_reference")
def api_template_reference():
    """テンプレートに同梱されたデフォルト見本画像を返す（プレビュー用）。"""
    template = request.args.get("template")
    p = slides_svc.template_reference_path(template or "")
    if p is None:
        return jsonify({"error": "not found"}), 404
    return send_file(str(p), mimetype="image/png")


@bp.route("/api/suggest_name")
def api_suggest_name():
    """自動採番のプロジェクト名候補を返す。"""
    return jsonify({"name": _next_project_name()})


@bp.route("/api/create_project", methods=["POST"])
def api_create_project():
    """新規プロジェクトを作成し、Markdown・プロンプトを保存する。"""
    data = request.get_json() or {}
    folder = (data.get("name") or "").strip() or _next_project_name()
    if not _valid_folder_name(folder):
        return jsonify({"error": "プロジェクト名に使えない文字が含まれています"}), 400
    if (OUTPUT_ROOT / folder).exists() and slides_svc.markdown_path(folder).exists():
        return jsonify({"error": f"「{folder}」は既に存在します"}), 400

    markdown = data.get("markdown", "") or ""
    prompt = (data.get("prompt") or "").strip() or slides_svc.DEFAULT_PROMPT
    template = (data.get("template") or "").strip()

    slides_svc.project_dir(folder)
    slides_svc.markdown_path(folder).write_text(markdown, encoding="utf-8")
    slides_svc.prompt_path(folder).write_text(prompt, encoding="utf-8")

    # テンプレートの初期化（見本画像コピー＋右下ネガティブスペース設定の保存）。
    # 見本画像はこの後アップロードされれば upload_reference 側で上書きされる。
    init = slides_svc.init_project_from_template(folder, template)

    return jsonify({
        "status": "ok",
        "project": folder,
        "applied_template_reference": init["applied_reference"],
    })


@bp.route("/api/project/<project_name>")
def api_project_detail(project_name: str):
    return jsonify(_project_status(project_name))


@bp.route("/api/delete_project", methods=["POST"])
def api_delete_project():
    """プロジェクトをアプリ一覧から非表示にする（ソフト削除）。

    ファイル（PNG・PPTX・見本画像など）は OUTPUT_ROOT 内に残したまま、
    markdown.md を __markdown.md にリネームするだけ。一覧は markdown.md の有無で
    判定しているため、リネームすれば一覧から消える。__markdown.md を元に戻せば復元可能。
    """
    data = request.get_json() or {}
    name = (data.get("project") or "").strip()
    if not _valid_folder_name(name):
        return jsonify({"error": "プロジェクト名が不正です"}), 400

    target = (OUTPUT_ROOT / name).resolve()
    # OUTPUT_ROOT の直下のみ対象（パストラバーサル防止）
    if target.parent != OUTPUT_ROOT.resolve() or not target.is_dir():
        return jsonify({"error": "対象プロジェクトが見つかりません"}), 404

    md = target / slides_svc.MARKDOWN_FILE
    if not md.exists():
        return jsonify({"error": "対象プロジェクトが見つかりません"}), 404
    md.rename(target / f"__{slides_svc.MARKDOWN_FILE}")

    return jsonify({"status": "ok", "deleted": name})


@bp.route("/api/save_markdown", methods=["POST"])
def api_save_markdown():
    """ペイン2の編集内容（Markdown・プロンプト）を保存する。"""
    data = request.get_json() or {}
    project_name = data.get("project")
    if not project_name:
        return jsonify({"error": "project is required"}), 400
    if "markdown" in data:
        slides_svc.markdown_path(project_name).write_text(data["markdown"] or "", encoding="utf-8")
    if "prompt" in data and (data.get("prompt") or "").strip():
        slides_svc.prompt_path(project_name).write_text(data["prompt"].strip(), encoding="utf-8")
    return jsonify({"status": "ok"})


@bp.route("/api/upload_reference", methods=["POST"])
def api_upload_reference():
    """見本デザイン画像をアップロードし _reference.png として保存する。"""
    project_name = request.form.get("project")
    if not project_name:
        return jsonify({"error": "project is required"}), 400
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "画像ファイルがありません"}), 400
    try:
        from PIL import Image
        img = Image.open(file.stream).convert("RGB")
    except Exception as e:
        return jsonify({"error": f"画像として読み込めませんでした: {e}"}), 400
    slides_svc.project_dir(project_name)
    img.save(slides_svc.reference_path(project_name), "PNG")
    return jsonify({"status": "ok", "message": "見本画像を保存しました"})


@bp.route("/api/reference_image")
def api_reference_image():
    project_name = request.args.get("project")
    if not project_name:
        return jsonify({"error": "project is required"}), 400
    p = slides_svc.reference_path(project_name)
    if not p.exists():
        return jsonify({"error": "not found"}), 404
    return send_file(str(p), mimetype="image/png")


# ---------------------------------------------------------------------------
# スライド生成（ペイン2 → ペイン3）
# ---------------------------------------------------------------------------

@bp.route("/api/slide_plan", methods=["POST"])
def api_slide_plan():
    """Markdown の見出し1/2/3ごとに構成案を生成する。"""
    data = request.get_json() or {}
    project_name = data.get("project")
    if not project_name:
        return jsonify({"error": "project is required"}), 400
    # 直前の編集を保存してから plan 化
    if "markdown" in data:
        slides_svc.markdown_path(project_name).write_text(data["markdown"] or "", encoding="utf-8")
    plan = slides_svc.generate_slide_plan(project_name)
    if not plan:
        return jsonify({"error": "見出し（# / ## / ###）が見つかりませんでした"}), 400
    return jsonify({"plan": plan, "count": len(plan)})


@bp.route("/api/generate_slide", methods=["POST"])
def api_generate_slide():
    """構成案の1スライドを画像生成する（index指定・任意の修正指示）。"""
    data = request.get_json() or {}
    project_name = data.get("project")
    index = data.get("index")
    fix = data.get("fix", "") or ""
    if not project_name or index is None:
        return jsonify({"error": "project / index が必要です"}), 400

    plan = slides_svc.load_plan(project_name)
    if index < 0 or index >= len(plan):
        return jsonify({"error": "index が範囲外です"}), 400

    project_prompt = slides_svc.read_prompt(project_name)
    try:
        slides_svc.generate_slide_image(
            project_name, index, plan[index], len(plan), project_prompt, fix=fix,
        )
    except Exception as e:
        return jsonify({"error": f"スライド生成エラー: {e}"}), 500

    return jsonify({
        "status": "ok",
        "index": index,
        "image_url": f"/api/slide_image?project={project_name}&index={index}&t={datetime.now().timestamp()}",
    })


@bp.route("/api/regenerate_all", methods=["POST"])
def api_regenerate_all():
    """全スライドを（任意の修正指示つきで）まとめて再生成する。

    fixes: {index: 修正テキスト} の辞書（無いものは空指示）
    """
    data = request.get_json() or {}
    project_name = data.get("project")
    fixes = data.get("fixes", {}) or {}
    if not project_name:
        return jsonify({"error": "project is required"}), 400

    plan = slides_svc.load_plan(project_name)
    if not plan:
        return jsonify({"error": "構成案がありません。先にスライド生成してください。"}), 400

    project_prompt = slides_svc.read_prompt(project_name)
    errors = []
    for i in range(len(plan)):
        fix = fixes.get(str(i)) or fixes.get(i) or ""
        try:
            slides_svc.generate_slide_image(
                project_name, i, plan[i], len(plan), project_prompt, fix=fix,
            )
        except Exception as e:
            errors.append({"index": i, "error": str(e)})

    return jsonify({"status": "ok", "count": len(plan), "errors": errors})


@bp.route("/api/slide_image")
def api_slide_image():
    project_name = request.args.get("project")
    index = request.args.get("index", type=int)
    if not project_name or index is None:
        return jsonify({"error": "project / index が必要です"}), 400
    png = slides_svc.slide_png_path(project_name, index)
    if not png.exists():
        return jsonify({"error": "not found"}), 404
    return send_file(str(png), mimetype="image/png")


# ---------------------------------------------------------------------------
# PPTX 書き出し
# ---------------------------------------------------------------------------

@bp.route("/api/build_pptx", methods=["POST"])
def api_build_pptx():
    data = request.get_json() or {}
    project_name = data.get("project")
    if not project_name:
        return jsonify({"error": "project is required"}), 400
    plan = slides_svc.load_plan(project_name)
    if not plan:
        return jsonify({"error": "構成案がありません"}), 400
    try:
        out = slides_svc.build_pptx(project_name, len(plan))
    except Exception as e:
        return jsonify({"error": f"PPTX組立エラー: {e}"}), 500
    return jsonify({
        "status": "ok",
        "message": "PPTX を作成しました",
        "pptx_path": str(out),
        "download_url": f"/api/download_pptx?project={project_name}",
    })


@bp.route("/api/download_pptx")
def api_download_pptx():
    project_name = request.args.get("project")
    if not project_name:
        return jsonify({"error": "project is required"}), 400
    pptx = slides_svc.pptx_path(project_name)
    if not pptx.exists():
        return jsonify({"error": "not found"}), 404
    return send_file(str(pptx), as_attachment=True, download_name=f"{project_name}.pptx")
