from __future__ import annotations

import os
import re
import tempfile
from datetime import datetime
from pathlib import Path

from flask import Blueprint, jsonify, render_template, request, send_file

from . import medical_images as medimg_svc
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
    materials = [slides_svc.list_material_indexes(project_name, i) for i in range(len(plan))]
    library_materials = [
        sorted(int(k) for k in slides_svc.library_materials(s) if k.isdigit()) for s in plan
    ]
    config = slides_svc.read_project_config(project_name)
    # ドキュメント画像は Markdown の参照のうち、実ファイルが残っているものだけ返す
    doc_images = [
        [
            rel for rel in (s.get("images") or [])
            if slides_svc.resolve_doc_image(project_name, rel) is not None
        ]
        for s in plan
    ]
    return {
        "project": project_name,
        "markdown": slides_svc.read_markdown(project_name),
        "prompt": slides_svc.read_prompt(project_name),
        "has_reference": slides_svc.reference_path(project_name).exists(),
        "plan": plan,
        "generated": generated,
        "materials": materials,
        "library_materials": library_materials,
        "doc_images": doc_images,
        "source": config.get("source") or {},
        "condense": bool(config.get("condense", True)),
        "pptx_notes": bool(config.get("pptx_notes", True)),
        "pptx_exists": slides_svc.pptx_path(project_name).exists(),
        "reference_check": slides_svc.reference_check(plan),
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
    """新規作成時に選べるプロンプト雛形を返す（組み込み＋自作）。"""
    return jsonify({
        "templates": slides_svc.all_templates(),
        "default": slides_svc.DEFAULT_PROMPT,
        # デフォルト見本画像を持つテンプレート名（フロントでプレビュー表示に使う）
        "with_reference": slides_svc.templates_with_reference(),
        # 自作テンプレート名（削除ボタンの出し分けに使う）
        "user": sorted(slides_svc.list_user_templates()),
    })


@bp.route("/api/save_template", methods=["POST"])
def api_save_template():
    """生成済みスライドを見本画像にして、自作テンプレートとして保存する。"""
    data = request.get_json() or {}
    project_name = (data.get("project") or "").strip()
    name = (data.get("name") or "").strip()
    prompt = data.get("prompt") or ""
    overwrite = bool(data.get("overwrite"))
    index = data.get("index")

    if not project_name or not _valid_folder_name(project_name):
        return jsonify({"error": "project is required"}), 400
    try:
        index = int(index)
    except (TypeError, ValueError):
        return jsonify({"error": "index が不正です"}), 400

    # 元プロジェクトの右下余白設定を引き継ぐ（ロゴ用スペースの有無も見た目の一部）
    negative_space = slides_svc.read_project_config(project_name).get("negative_space")

    try:
        result = slides_svc.save_user_template(
            name,
            prompt,
            slides_svc.slide_png_path(project_name, index),
            negative_space=negative_space,
            overwrite=overwrite,
        )
    except slides_svc.TemplateError as e:
        # 既存名との衝突だけは、フロントで上書き確認を出せるように区別して返す
        exists = slides_svc.user_template_dir(name).exists() and not overwrite
        return jsonify({"error": str(e), "exists": exists}), 409 if exists else 400
    except Exception as e:
        return jsonify({"error": f"テンプレート保存エラー: {e}"}), 500

    return jsonify({
        "status": "ok",
        "name": result["name"],
        "message": f"テンプレート「{result['name']}」を保存しました",
    })


@bp.route("/api/save_prompt_template", methods=["POST"])
def api_save_prompt_template():
    """編集中のデザインプロンプトを、その場で自作テンプレートとして保存する。

    新規作成モーダルからスライド生成前に呼ばれるので、見本デザイン画像は
    「添付されたファイル」→「元にした雛形の見本画像」の順で引き継ぎ、
    どちらも無ければ見本画像なしのテンプレートにする。
    """
    name = (request.form.get("name") or "").strip()
    prompt = request.form.get("prompt") or ""
    base = (request.form.get("base") or "").strip()
    overwrite = request.form.get("overwrite") == "1"

    reference: Path | None = None
    tmp_ref: Path | None = None
    file = request.files.get("file")
    if file and file.filename:
        try:
            from PIL import Image

            img = Image.open(file.stream).convert("RGB")
        except Exception as e:
            return jsonify({"error": f"画像として読み込めませんでした: {e}"}), 400
        fd, tmp_name = tempfile.mkstemp(suffix=".png")
        os.close(fd)
        tmp_ref = Path(tmp_name)
        img.save(tmp_ref, "PNG")
        reference = tmp_ref
    else:
        reference = slides_svc.template_reference_path(base)

    try:
        result = slides_svc.save_user_template(
            name,
            prompt,
            reference,
            negative_space=slides_svc.template_negative_space(base),
            overwrite=overwrite,
        )
    except slides_svc.TemplateError as e:
        exists = slides_svc.user_template_dir(name).exists() and not overwrite
        return jsonify({"error": str(e), "exists": exists}), 409 if exists else 400
    except Exception as e:
        return jsonify({"error": f"テンプレート保存エラー: {e}"}), 500
    finally:
        if tmp_ref is not None:
            tmp_ref.unlink(missing_ok=True)

    return jsonify({
        "status": "ok",
        "name": result["name"],
        "message": f"テンプレート「{result['name']}」を保存しました",
    })


@bp.route("/api/delete_template", methods=["POST"])
def api_delete_template():
    """自作テンプレートを削除する。"""
    data = request.get_json() or {}
    name = (data.get("name") or "").strip()
    try:
        slides_svc.delete_user_template(name)
    except slides_svc.TemplateError as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"status": "ok", "deleted": name})


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
    doc_url = (data.get("doc_url") or "").strip()

    slides_svc.project_dir(folder)
    slides_svc.markdown_path(folder).write_text(markdown, encoding="utf-8")
    slides_svc.prompt_path(folder).write_text(prompt, encoding="utf-8")

    # テンプレートの初期化（見本画像コピー＋右下ネガティブスペース設定の保存）。
    # 見本画像はこの後アップロードされれば upload_reference 側で上書きされる。
    init = slides_svc.init_project_from_template(folder, template)

    # ドキュメント URL が指定されていれば、その場で取り込んで Markdown を作る。
    # 取り込みに失敗してもプロジェクト自体は残し、原因を返して手直しできるようにする。
    imported = None
    import_error = ""
    if doc_url:
        try:
            imported = slides_svc.import_document(folder, url=doc_url)
        except slides_svc.DocumentImportError as e:
            import_error = str(e)

    return jsonify({
        "status": "ok",
        "project": folder,
        "applied_template_reference": init["applied_reference"],
        "imported": imported,
        "import_error": import_error,
    })


@bp.route("/api/import_document", methods=["POST"])
def api_import_document():
    """Google ドキュメントの URL、または .docx ファイルを取り込んで Markdown 化する。

    JSON で {"project", "url"} を送るか、multipart で project ＋ file を送る。
    既存の Markdown は取り込み結果で置き換わる。
    """
    if request.files.get("file"):
        project_name = (request.form.get("project") or "").strip()
        url = ""
        docx_bytes = request.files["file"].read()
    else:
        data = request.get_json() or {}
        project_name = (data.get("project") or "").strip()
        url = (data.get("url") or "").strip()
        docx_bytes = b""
        if not url:
            return jsonify({"error": "ドキュメントの URL を入力してください"}), 400

    if not project_name or not _valid_folder_name(project_name):
        return jsonify({"error": "project is required"}), 400

    try:
        result = slides_svc.import_document(project_name, url=url, docx_bytes=docx_bytes)
    except slides_svc.DocumentImportError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"取り込みエラー: {e}"}), 500

    return jsonify({
        "status": "ok",
        "markdown": result["markdown"],
        "headings": result["headings"],
        "images": result["images"],
        "warning": result["warning"],
        "message": (
            f"取り込みました（見出し {result['headings']}件 / 画像 {result['images']}枚）"
        ),
    })


@bp.route("/api/doc_image")
def api_doc_image():
    """取り込んだドキュメント画像を返す（プレビュー用）。"""
    project_name = request.args.get("project")
    rel = request.args.get("path")
    if not project_name or not rel:
        return jsonify({"error": "project / path が必要です"}), 400
    p = slides_svc.resolve_doc_image(project_name, rel)
    if p is None:
        return jsonify({"error": "not found"}), 404
    return send_file(str(p), mimetype="image/png")


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
# 素材画像（スライドごとにユーザーが貼り込む画像）
# ---------------------------------------------------------------------------

@bp.route("/api/upload_material", methods=["POST"])
def api_upload_material():
    """スライドに紐づく素材画像をアップロードして保存する。"""
    project_name = request.form.get("project")
    if not project_name or not _valid_folder_name(project_name):
        return jsonify({"error": "project is required"}), 400
    try:
        index = int(request.form.get("index", ""))
    except (TypeError, ValueError):
        return jsonify({"error": "index が不正です"}), 400
    if index < 0:
        return jsonify({"error": "index が不正です"}), 400

    if len(slides_svc.list_material_indexes(project_name, index)) >= slides_svc.MAX_MATERIALS_PER_SLIDE:
        return jsonify({"error": f"素材画像は1スライドあたり最大{slides_svc.MAX_MATERIALS_PER_SLIDE}枚までです"}), 400

    file = request.files.get("file")
    if not file:
        return jsonify({"error": "画像ファイルがありません"}), 400
    try:
        k = slides_svc.save_material(project_name, index, file.stream)
    except Exception as e:
        return jsonify({"error": f"画像として読み込めませんでした: {e}"}), 400

    return jsonify({
        "status": "ok",
        "index": index,
        "mindex": k,
        "materials": slides_svc.list_material_indexes(project_name, index),
    })


@bp.route("/api/delete_material", methods=["POST"])
def api_delete_material():
    """スライドに紐づく素材画像を削除する。"""
    data = request.get_json() or {}
    project_name = (data.get("project") or "").strip()
    if not project_name or not _valid_folder_name(project_name):
        return jsonify({"error": "project is required"}), 400
    try:
        index = int(data.get("index"))
        k = int(data.get("mindex"))
    except (TypeError, ValueError):
        return jsonify({"error": "index / mindex が不正です"}), 400
    if index < 0 or k < 0:
        return jsonify({"error": "index / mindex が不正です"}), 400

    slides_svc.delete_material(project_name, index, k)
    slides_svc.clear_library_material(project_name, index, k)
    return jsonify({
        "status": "ok",
        "index": index,
        "materials": slides_svc.list_material_indexes(project_name, index),
        "library_materials": _library_materials(project_name, index),
    })


@bp.route("/api/material_image")
def api_material_image():
    project_name = request.args.get("project")
    index = request.args.get("index", type=int)
    k = request.args.get("mindex", type=int)
    if not project_name or index is None or k is None:
        return jsonify({"error": "project / index / mindex が必要です"}), 400
    p = slides_svc.material_path(project_name, index, k)
    if not p.exists():
        return jsonify({"error": "not found"}), 404
    return send_file(str(p), mimetype="image/png")


# ---------------------------------------------------------------------------
# 医用画像ライブラリ（購入済みの解剖図を素材として取り込む）
# ---------------------------------------------------------------------------

@bp.route("/api/medical_library_status")
def api_medical_library_status():
    """ライブラリが使えるかと、絞り込み用の部位一覧を返す。"""
    if not medimg_svc.available():
        return jsonify({"available": False, "path": str(medimg_svc.LIBRARY_DIR)})
    return jsonify({
        "available": True,
        "path": str(medimg_svc.LIBRARY_DIR),
        "count": len(medimg_svc.catalog()),
        "regions": medimg_svc.regions(),
    })


@bp.route("/api/medical_images")
def api_medical_images():
    """ライブラリをキーワードと属性で検索する。"""
    if not medimg_svc.available():
        return jsonify({"error": "医用画像ライブラリが見つかりません"}), 404
    hits = medimg_svc.search(
        request.args.get("q", "").strip(),
        region=request.args.get("region") or None,
        category=request.args.get("category") or None,
        text_in_image=request.args.get("text_in_image") or None,
    )
    return jsonify({"images": hits})


@bp.route("/api/medical_image_file")
def api_medical_image_file():
    """ライブラリの画像を選択用プレビューとして配信する。"""
    p = medimg_svc.resolve(request.args.get("path", ""))
    if p is None:
        return jsonify({"error": "not found"}), 404
    return send_file(str(p))


@bp.route("/api/add_material_from_library", methods=["POST"])
def api_add_material_from_library():
    """ライブラリの図をスライドの素材画像としてプロジェクトに取り込む。

    以降はアップロードした素材と完全に同じ扱いになるため、削除も生成時の
    添付も既存の仕組みがそのまま効く。
    """
    data = request.get_json() or {}
    project_name = (data.get("project") or "").strip()
    if not project_name or not _valid_folder_name(project_name):
        return jsonify({"error": "project is required"}), 400
    try:
        index = int(data.get("index"))
    except (TypeError, ValueError):
        return jsonify({"error": "index が不正です"}), 400
    if index < 0:
        return jsonify({"error": "index が不正です"}), 400

    if len(slides_svc.list_material_indexes(project_name, index)) >= slides_svc.MAX_MATERIALS_PER_SLIDE:
        return jsonify({"error": f"素材画像は1スライドあたり最大{slides_svc.MAX_MATERIALS_PER_SLIDE}枚までです"}), 400

    src = medimg_svc.resolve((data.get("path") or "").strip())
    if src is None:
        return jsonify({"error": "ライブラリにその画像がありません"}), 404

    rel = (data.get("path") or "").strip()
    try:
        with src.open("rb") as fh:
            k = slides_svc.save_material(project_name, index, fh)
    except Exception as e:
        return jsonify({"error": f"画像として読み込めませんでした: {e}"}), 400

    # ライブラリ由来だと覚えておく。生成時にAIへ別の図解を描かせないため
    slides_svc.set_library_material(project_name, index, k, rel)

    return jsonify({
        "status": "ok",
        "index": index,
        "mindex": k,
        "materials": slides_svc.list_material_indexes(project_name, index),
        "library_materials": _library_materials(project_name, index),
    })


def _library_materials(project_name: str, index: int) -> list[int]:
    """スライドの素材のうち、ライブラリ由来のものの k を返す。"""
    plan = slides_svc.load_plan(project_name)
    if not (0 <= index < len(plan)):
        return []
    return sorted(int(k) for k in slides_svc.library_materials(plan[index]) if k.isdigit())


# ---------------------------------------------------------------------------
# スライド生成（ペイン2 → ペイン3）
# ---------------------------------------------------------------------------

@bp.route("/api/slide_plan", methods=["POST"])
def api_slide_plan():
    """Markdown の見出し1/2/3ごとに構成案を生成する。

    condense=True のときは本文をテキストLLMでキーワードに絞り込む。
    要約だけ失敗した場合は構成案を返しつつ warning で知らせる（生成は続行可能）。
    """
    data = request.get_json() or {}
    project_name = data.get("project")
    if not project_name:
        return jsonify({"error": "project is required"}), 400
    # 直前の編集を保存してから plan 化
    if "markdown" in data:
        slides_svc.markdown_path(project_name).write_text(data["markdown"] or "", encoding="utf-8")

    condense = bool(data.get("condense", True))
    cfg = slides_svc.read_project_config(project_name)
    cfg["condense"] = condense
    slides_svc.write_project_config(project_name, cfg)

    warning = ""
    try:
        plan = slides_svc.generate_slide_plan(project_name, condense=condense)
    except slides_svc.CondenseWarning as e:
        plan = e.plan
        warning = f"本文の要約に失敗したため、本文をそのまま素材に使います: {e}"

    if not plan:
        return jsonify({"error": "見出し（# / ## / ###）が見つかりませんでした"}), 400
    return jsonify({
        "plan": plan,
        "count": len(plan),
        "warning": warning,
        "reference_check": slides_svc.reference_check(plan),
    })


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

    with_notes = bool(data.get("notes", True))
    cfg = slides_svc.read_project_config(project_name)
    cfg["pptx_notes"] = with_notes
    slides_svc.write_project_config(project_name, cfg)

    try:
        out = slides_svc.build_pptx(project_name, len(plan), with_notes=with_notes)
    except Exception as e:
        return jsonify({"error": f"PPTX組立エラー: {e}"}), 500

    noted = sum(1 for s in plan if slides_svc.slide_note(s)) if with_notes else 0
    message = (
        f"PPTX を作成しました（{noted}枚のノートに台本を転記）" if noted
        else "PPTX を作成しました"
    )
    check = slides_svc.reference_check(plan)
    if not check["ok"]:
        message += f" ※{check['message']}"
    return jsonify({
        "status": "ok",
        "message": message,
        "pptx_path": str(out),
        "noted": noted,
        "download_url": f"/api/download_pptx?project={project_name}",
        "reference_check": check,
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
