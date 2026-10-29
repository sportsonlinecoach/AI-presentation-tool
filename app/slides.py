"""
スライド生成サービス層

方針:
0. 入力は「Markdown を直接書く」か「Google ドキュメント / .docx を取り込む」の2通り
   - 取り込み時は見出し・本文・箇条書き・画像を Markdown へ正規化し、
     画像はプロジェクト内の _doc/ に保存して Markdown から参照する
     （Markdown を唯一の正とすることで、取り込み後の手直しがそのまま効く）
1. Markdown の見出し1/2/3ごとに1スライドへ分解（決定論パーサ・LLM不要）
   - h4以降・箇条書き・段落は直近スライドの本文に束ねる
   - 画像参照は直近スライドの「ドキュメント画像」として紐づける
2. 任意で本文をテキストLLMに要約させ、キーワード＋図解指示に絞る（condense.py）
   - 長い本文をそのまま描かせると文字量が多すぎて読めないため
3. 各スライドを GPT Image 2.0 で1枚絵として生成
   - プロジェクトの見本画像があれば images.edit の参照に使う（無ければ images.generate）
   - ドキュメント画像・素材画像があれば添付し、主役として大きく配置させる
   - プロジェクトごとの生成プロンプトをスタイル指示に流用
   - 任意で「修正指示(fix)」をプロンプト末尾に追加して再生成
4. PNG 群を python-pptx で「1枚画像=1スライド（16:9 全面）」の PPTX に組立

出力先は OUTPUT_ROOT/<プロジェクト名>/ （既定は iCloud Drive の Downloads）。
"""
from __future__ import annotations

import os
import io
import re
import json
import base64
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches

from . import condense as condense_svc
from .gdoc import DocumentImportError, ParsedDocument, parse_docx, parse_google_doc

# ---------------------------------------------------------------------------
# パス設定
# ---------------------------------------------------------------------------

OUTPUT_ROOT = Path(
    os.environ.get(
        "OUTPUT_ROOT",
        Path.home() / "Library/Mobile Documents/com~apple~CloudDocs/Downloads",
    )
)

SLIDE_IMAGE_MODEL = os.environ.get("SLIDE_IMAGE_MODEL", "gpt-image-2")

# gpt-image の landscape サイズ（3:2）。生成後に 16:9 へ縦圧縮フィットする。
GEN_SIZE = "1536x1024"
TARGET_W, TARGET_H = 1536, 864  # 16:9

# images.edit に一度に渡せる画像の上限（API仕様）。参照画像もこの枚数に含まれる。
MAX_ATTACHED_IMAGES = 16

# 1スライドに追加できる素材画像の上限。
# API の上限内であれば制限する理由はないので、参照画像1枚分だけ余らせた枚数にする。
MAX_MATERIALS_PER_SLIDE = MAX_ATTACHED_IMAGES - 1

# 取り込んだドキュメント画像を保存する長辺の上限。
# 1スライドに多数の画像を添付するとリクエストが大きくなり失敗しやすいため、
# スライド幅(1536px)に対して十分な解像度を保ちつつ軽くしておく。
DOC_IMAGE_MAX_SIDE = 1600

# プロジェクト内の固定ファイル名
MARKDOWN_FILE = "markdown.md"
PROMPT_FILE = "_prompt.txt"
REFERENCE_FILE = "_reference.png"
PLAN_FILE = "plan.json"
CONFIG_FILE = "_config.json"

# 取り込んだドキュメント画像の置き場（プロジェクト直下のサブフォルダ）
DOC_ASSETS_DIRNAME = "_doc"


# ---------------------------------------------------------------------------
# プロンプトテンプレート（新規作成時に選択・編集できる雛形）
# ---------------------------------------------------------------------------

# 全テンプレート共通で末尾に付与する厳守ルール。
# 見本デザイン画像はあくまで「見た目のお手本」であって、
# そこに写っている文字・トピックなどの内容は一切流用させない。
REFERENCE_TEXT_RULE = (
    "\n【見本デザイン画像の扱い・全体共通／最重要・厳守】\n"
    "・見本（参照）デザイン画像は、配色・フォントの雰囲気・レイアウト・装飾の"
    "テイストを参考にするための『デザインの見本』にすぎない。\n"
    "・見本画像に写っている見出し・本文・キャプション・ロゴ・数字などの"
    "テキスト内容は一切引っ張ってこない・流用しない・転記しない・再現しないこと。\n"
    "・見本画像のトピックや文言は、このスライドの内容とは無関係だとみなすこと。\n"
    "・スライドに描画してよい文字は、別途指定される見出し・本文の素材だけ。"
    "それ以外の言葉を勝手に足さない。"
)

_PROMPT_TEMPLATES_BASE: dict[str, str] = {
    "SHI-GYO": (
        "あなたは最高のプレゼンテーションスライドメーカーです。\n"
        "参照画像のスライドのテイスト（デザイン・色味）を踏襲して、"
        "情報発信プロジェクト「Dr.SHI-GYO」に関連するスライドを生成してください。\n"
        "・文字数はできるだけ減らす（キーワード中心）\n"
        "・高齢者も見るセミナーなので小さい文字は禁止（引用文献情報のみ小さく）\n"
        "・少ない文字数・大きな文字で、イラスト・図表・写真も使って直感的に理解させる\n"
        "【右下にロゴ用スペースを空けるルール・厳守】\n"
        "・右下の角に、ロゴやワイプを置くための小さな空きスペース"
        "（背景のまま・文字も図版も置かない／目安は横幅の約15%・高さの約20%）を必ず確保する。\n"
        "・カードや図版・本文は、その右下の角を避けて配置する。"
        "右下にカードが重なる場合は、最後の項目を左に寄せる・カードを一段上げる・"
        "下段を3列ではなく左寄せにするなどして、右下の角だけは背景を残すこと。\n"
        "・ただし空けるのは右下の角だけ。右端を縦一列にまるごと空けたり、"
        "画面の右3〜4割を縦長の空白にするのは絶対に禁止。\n"
        "・左・中央・上部、そして右上〜右中央は画面いっぱいに使い、横幅を端近くまで使い切ること。"
    ),
    "シンプル": (
        "クリーンでミニマルなビジネスプレゼンのスライド。"
        "白基調の背景、余白を広めに取り、見出しを大きく、本文は読みやすい階層で配置。"
        "アクセントカラーを1色だけ使う。装飾は控えめに。"
    ),
    "ポップ": (
        "明るくカラフルでポップなプレゼンスライド。"
        "親しみやすい配色、丸みのある図形、わかりやすいアイコン的イラストを添える。"
        "見出しは目を引く大きさで。"
    ),
    "ダーク": (
        "ダークテーマのモダンなプレゼンスライド。"
        "濃紺〜黒の背景に明るい文字、ネオン的なアクセント1色。"
        "テック系の洗練された印象。"
    ),
    "和風・落ち着き": (
        "落ち着いた和モダンのプレゼンスライド。"
        "生成り・墨・若干の差し色。余白を活かし、上品で読みやすいレイアウト。"
    ),
}

# 全テンプレートに共通ルールを付与して確定させる。
PROMPT_TEMPLATES: dict[str, str] = {
    name: base + REFERENCE_TEXT_RULE for name, base in _PROMPT_TEMPLATES_BASE.items()
}

DEFAULT_PROMPT = PROMPT_TEMPLATES["SHI-GYO"]

# テンプレートに同梱する「デフォルト見本デザイン画像」。
# キーは PROMPT_TEMPLATES のテンプレート名、値は app/template_assets/ 内のファイル名。
# 新規作成時に見本画像が未アップロードなら、ここで指定した画像が初期見本になる。
# （プロジェクトで見本画像をアップロードすると、そのプロジェクトの見本に差し替わる）
TEMPLATE_ASSETS_DIR = Path(__file__).resolve().parent / "template_assets"

TEMPLATE_REFERENCES: dict[str, str] = {
    "SHI-GYO": "shigyo_reference.png",
}

# テンプレートごとの「右下ネガティブスペース（ロゴ/ワイプ用の余白）」設定。
# プロンプトで余白を空けさせたうえで、生成後に下記の矩形を背景色で軽く均して
# きれいな余白に仕上げる（端のフレームは inset 分だけ残す）。
#   right / bottom : 画像の右端・下端からの割合（空ける領域の幅・高さ）
#   inset          : 外周フレームを残すための内側マージン（幅に対する割合）
#   feather        : 合成マスクのぼかし量（幅に対する割合）
TEMPLATE_NEGATIVE_SPACE: dict[str, dict] = {
    "SHI-GYO": {"right": 0.15, "bottom": 0.20, "inset": 0.014, "feather": 0.018},
}


# ---------------------------------------------------------------------------
# ユーザー定義テンプレート
#
# 生成できたスライドを見本画像として登録し、自分用のテンプレートを増やせるようにする。
# 組み込みテンプレート（上記の定数）と違ってコード変更を伴わないので、
# OUTPUT_ROOT 配下に置いて他のユーザーデータと一緒に扱う（iCloud で同期される）。
#
#   OUTPUT_ROOT/_slide_templates/<テンプレート名>/
#       prompt.txt     デザインプロンプト
#       reference.png  見本デザイン画像（元になったスライド）
#       config.json    右下ネガティブスペースなどの設定
# ---------------------------------------------------------------------------

USER_TEMPLATES_DIRNAME = "_slide_templates"
USER_TEMPLATE_PROMPT_FILE = "prompt.txt"
USER_TEMPLATE_REFERENCE_FILE = "reference.png"
USER_TEMPLATE_CONFIG_FILE = "config.json"


class TemplateError(Exception):
    """テンプレート保存・削除の失敗（そのままユーザーに見せる文言）。"""


def user_templates_root() -> Path:
    return OUTPUT_ROOT / USER_TEMPLATES_DIRNAME


def valid_template_name(name: str) -> bool:
    """フォルダ名として安全なテンプレート名か。"""
    if not name or name.startswith(".") or len(name) > 40:
        return False
    return not any(c in name for c in ("/", "\\", ":", ".."))


def user_template_dir(name: str) -> Path:
    return user_templates_root() / name


def list_user_templates() -> dict:
    """ユーザー定義テンプレートの {名前: プロンプト} を返す。"""
    root = user_templates_root()
    if not root.exists():
        return {}

    result = {}
    for d in sorted(root.iterdir()):
        if not d.is_dir() or d.name.startswith("."):
            continue
        prompt_file = d / USER_TEMPLATE_PROMPT_FILE
        if not prompt_file.exists():
            continue
        try:
            result[d.name] = prompt_file.read_text(encoding="utf-8")
        except OSError:
            continue
    return result


def all_templates() -> dict:
    """組み込み＋ユーザー定義のテンプレート一覧。

    同名の場合はユーザー定義を優先せず、保存時に組み込み名を弾いて衝突を防ぐ。
    """
    templates = dict(PROMPT_TEMPLATES)
    templates.update(list_user_templates())
    return templates


def template_reference_path(template_name: str) -> Path | None:
    """テンプレートの見本画像のパスを返す（無ければ None）。

    組み込みテンプレートは app/template_assets/ を、
    ユーザー定義テンプレートは OUTPUT_ROOT/_slide_templates/ を見る。
    """
    name = template_name or ""

    fname = TEMPLATE_REFERENCES.get(name)
    if fname:
        p = TEMPLATE_ASSETS_DIR / fname
        if p.exists():
            return p

    if valid_template_name(name):
        p = user_template_dir(name) / USER_TEMPLATE_REFERENCE_FILE
        if p.exists():
            return p
    return None


def templates_with_reference() -> list[str]:
    """見本画像を持つテンプレート名の一覧。"""
    names = list(TEMPLATE_REFERENCES) + list(list_user_templates())
    return [name for name in names if template_reference_path(name)]


def template_negative_space(template_name: str) -> dict | None:
    """テンプレートに定義された右下ネガティブスペース設定を返す（無ければ None）。"""
    name = template_name or ""

    ns = TEMPLATE_NEGATIVE_SPACE.get(name)
    if ns:
        return dict(ns)

    if valid_template_name(name):
        cfg_file = user_template_dir(name) / USER_TEMPLATE_CONFIG_FILE
        if cfg_file.exists():
            try:
                cfg = json.loads(cfg_file.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                return None
            ns = cfg.get("negative_space")
            if ns:
                return dict(ns)
    return None


def save_user_template(
    name: str, prompt: str, reference_png: Path | None = None,
    negative_space: dict = None, overwrite: bool = False,
) -> dict:
    """ユーザー定義テンプレートを保存する。

    reference_png を渡すとそれを見本デザイン画像として登録する。
    None の場合は見本画像なし（既存テンプレートの上書き時は今の見本画像を残す）。
    """
    name = (name or "").strip()
    if not valid_template_name(name):
        raise TemplateError(
            "テンプレート名に使えない文字が含まれています"
            "（/ \\ : . で始まる名前は使えません／40文字以内）"
        )
    if name in PROMPT_TEMPLATES:
        raise TemplateError(
            f"「{name}」は最初から入っているテンプレート名です。別の名前にしてください。"
        )
    if not (prompt or "").strip():
        raise TemplateError("デザインプロンプトが空です")
    if reference_png is not None and not reference_png.exists():
        raise TemplateError(
            "元にするスライド画像が見つかりません。先にスライドを生成してください。"
        )

    d = user_template_dir(name)
    if d.exists() and not overwrite:
        raise TemplateError(f"「{name}」は既にあります。上書きしますか？")

    d.mkdir(parents=True, exist_ok=True)
    (d / USER_TEMPLATE_PROMPT_FILE).write_text(prompt.strip(), encoding="utf-8")
    # 見本画像は RGB PNG に正規化する（参照画像として渡す形式を揃える）
    if reference_png is not None:
        Image.open(reference_png).convert("RGB").save(
            d / USER_TEMPLATE_REFERENCE_FILE, "PNG"
        )
    (d / USER_TEMPLATE_CONFIG_FILE).write_text(
        json.dumps({"negative_space": negative_space or None}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return {
        "name": name,
        "negative_space": negative_space or None,
        "has_reference": (d / USER_TEMPLATE_REFERENCE_FILE).exists(),
    }


def delete_user_template(name: str) -> None:
    """ユーザー定義テンプレートを削除する（組み込みは削除できない）。"""
    name = (name or "").strip()
    if name in PROMPT_TEMPLATES:
        raise TemplateError("最初から入っているテンプレートは削除できません。")
    if not valid_template_name(name):
        raise TemplateError("テンプレート名が不正です")

    d = user_template_dir(name)
    if not d.is_dir():
        raise TemplateError(f"「{name}」が見つかりません")

    import shutil

    # OUTPUT_ROOT/_slide_templates 直下のみ削除対象（パストラバーサル防止）
    if d.resolve().parent != user_templates_root().resolve():
        raise TemplateError("テンプレート名が不正です")
    shutil.rmtree(d)


# ---------------------------------------------------------------------------
# プロジェクトのパス
# ---------------------------------------------------------------------------

def project_dir(project_name: str) -> Path:
    d = OUTPUT_ROOT / project_name
    d.mkdir(parents=True, exist_ok=True)
    return d


def markdown_path(project_name: str) -> Path:
    return project_dir(project_name) / MARKDOWN_FILE


def prompt_path(project_name: str) -> Path:
    return project_dir(project_name) / PROMPT_FILE


def reference_path(project_name: str) -> Path:
    return project_dir(project_name) / REFERENCE_FILE


def apply_template_reference(project_name: str, template_name: str) -> bool:
    """テンプレートのデフォルト見本画像をプロジェクトの _reference.png に設定する。

    既にプロジェクト独自の見本画像がある場合は上書きしない。
    コピーした場合は True、対象が無い／既存ありの場合は False を返す。
    """
    src = template_reference_path(template_name)
    if src is None:
        return False
    dst = reference_path(project_name)
    if dst.exists():
        return False
    import shutil

    project_dir(project_name)
    shutil.copyfile(src, dst)
    return True


def config_path(project_name: str) -> Path:
    return project_dir(project_name) / CONFIG_FILE


def read_project_config(project_name: str) -> dict:
    p = config_path(project_name)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def write_project_config(project_name: str, config: dict) -> None:
    config_path(project_name).write_text(
        json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def init_project_from_template(project_name: str, template_name: str) -> dict:
    """テンプレート選択時の初期化（見本画像コピー＋設定保存）をまとめて行う。

    返り値: {"applied_reference": bool, "negative_space": dict | None}
    """
    applied_reference = apply_template_reference(project_name, template_name)
    ns = template_negative_space(template_name)
    if ns:
        cfg = read_project_config(project_name)
        cfg["negative_space"] = ns
        write_project_config(project_name, cfg)
    return {"applied_reference": applied_reference, "negative_space": ns}


def plan_path(project_name: str) -> Path:
    return project_dir(project_name) / PLAN_FILE


def slide_png_path(project_name: str, index: int) -> Path:
    return project_dir(project_name) / f"slide_{index:02d}.png"


# ---------------------------------------------------------------------------
# 素材画像（スライドごとにユーザーが貼り込む画像）
# ---------------------------------------------------------------------------

def material_path(project_name: str, index: int, k: int) -> Path:
    return project_dir(project_name) / f"slide_{index:02d}_mat_{k:02d}.png"


def list_material_indexes(project_name: str, index: int) -> list[int]:
    """スライド index に存在する素材画像の k（ソート済み）を返す。"""
    d = OUTPUT_ROOT / project_name
    if not d.exists():
        return []
    prefix = f"slide_{index:02d}_mat_"
    ks: list[int] = []
    for p in d.iterdir():
        if not p.is_file() or not p.name.startswith(prefix) or not p.name.endswith(".png"):
            continue
        stem = p.name[len(prefix):-len(".png")]
        if stem.isdigit():
            ks.append(int(stem))
    return sorted(ks)


def list_material_paths(project_name: str, index: int) -> list[Path]:
    """スライド index の素材画像パス（k 昇順）を返す。"""
    return [material_path(project_name, index, k) for k in list_material_indexes(project_name, index)]


def next_material_index(project_name: str, index: int) -> int:
    ks = list_material_indexes(project_name, index)
    k = 0
    existing = set(ks)
    while k in existing:
        k += 1
    return k


def save_material(project_name: str, index: int, file_stream) -> int:
    """素材画像を PNG（透過保持）で保存し、割り当てた k を返す。"""
    img = Image.open(file_stream)
    img = img.convert("RGBA") if img.mode in ("RGBA", "LA", "P") else img.convert("RGB")
    project_dir(project_name)
    k = next_material_index(project_name, index)
    img.save(material_path(project_name, index, k), "PNG")
    return k


def delete_material(project_name: str, index: int, k: int) -> bool:
    p = material_path(project_name, index, k)
    if p.exists():
        p.unlink()
        return True
    return False


# ---------------------------------------------------------------------------
# ドキュメント取り込み（Google ドキュメント / .docx → Markdown ＋ _doc/ の画像）
# ---------------------------------------------------------------------------

def doc_assets_dir(project_name: str) -> Path:
    d = project_dir(project_name) / DOC_ASSETS_DIRNAME
    d.mkdir(parents=True, exist_ok=True)
    return d


def resolve_doc_image(project_name: str, rel_path: str) -> Path | None:
    """Markdown の画像参照をプロジェクト内の実ファイルへ解決する。

    Markdown はユーザーが自由に編集できるため、プロジェクトフォルダの外を
    指す参照（../ など）は無視する。存在しない場合も None を返す。
    """
    rel = (rel_path or "").strip()
    if not rel or "://" in rel:
        return None

    root = project_dir(project_name).resolve()
    try:
        target = (root / rel).resolve()
        target.relative_to(root)
    except (ValueError, OSError):
        return None
    return target if target.is_file() else None


def _clear_doc_assets(project_name: str) -> None:
    """再取り込み時に古いドキュメント画像を消す（Markdown ごと作り直すため）。"""
    d = project_dir(project_name) / DOC_ASSETS_DIRNAME
    if not d.exists():
        return
    for p in d.iterdir():
        if p.is_file() and p.suffix.lower() == ".png":
            p.unlink()


def _save_doc_image(project_name: str, seq: int, blob: bytes) -> str:
    """ドキュメント内の画像を _doc/img_NNN.png に保存し、相対パスを返す。

    PNG に正規化しておくことで、画像生成APIに渡す形式を揃えられる。
    大きすぎる画像は長辺 DOC_IMAGE_MAX_SIDE に縮小する。
    """
    img = Image.open(io.BytesIO(blob))
    img = img.convert("RGBA") if img.mode in ("RGBA", "LA", "P") else img.convert("RGB")

    longest = max(img.size)
    if longest > DOC_IMAGE_MAX_SIDE:
        scale = DOC_IMAGE_MAX_SIDE / longest
        img = img.resize(
            (max(1, int(img.width * scale)), max(1, int(img.height * scale))),
            Image.LANCZOS,
        )

    name = f"img_{seq:03d}.png"
    img.save(doc_assets_dir(project_name) / name, "PNG")
    return f"{DOC_ASSETS_DIRNAME}/{name}"


def document_to_markdown(project_name: str, parsed: ParsedDocument) -> str:
    """ブロック列を Markdown に変換し、画像を _doc/ へ保存する。

    画像は Markdown 上で `![](_doc/img_000.png)` として本文中の元の位置に残す。
    こうしておくと、あとで見出しを足したり順番を入れ替えたりしたときにも、
    画像がどのスライドに属するかを Markdown だけで決められる。
    """
    _clear_doc_assets(project_name)

    lines: list[str] = []
    seq = 0
    for block in parsed.blocks:
        if block.kind == "heading":
            level = min(max(block.level, 1), 3)
            lines.extend(["", "#" * level + " " + block.text, ""])
        elif block.kind == "bullet":
            lines.append(f"- {block.text}")
        elif block.kind == "para":
            lines.extend([block.text, ""])
        elif block.kind == "image":
            try:
                rel = _save_doc_image(project_name, seq, block.image_bytes)
            except Exception:
                continue
            seq += 1
            lines.extend(["", f"![]({rel})", ""])

    # 連続した空行を1つに畳んでから確定する
    md_lines: list[str] = []
    for line in lines:
        if not line.strip() and (not md_lines or not md_lines[-1].strip()):
            continue
        md_lines.append(line)
    return "\n".join(md_lines).strip() + "\n"


# 見出しスタイルが使われていないドキュメントは、そのままでは1枚の巨大なスライドに
# なってしまうため、取り込みは成功させつつ直し方を伝える。
NO_HEADING_WARNING = (
    "ドキュメント内に「見出し」スタイルが見つかりませんでした。\n"
    "このままだと全文が1枚のスライドになってしまいます。\n"
    "Google ドキュメント側で、スライドの区切りにしたい行を選んで"
    "「表示形式 → 段落スタイル → 見出し1／見出し2／見出し3」に設定してから"
    "取り込み直してください。\n"
    "（文字を太字・大きくしただけでは見出しとして扱われません）\n"
    "または、取り込んだ台本に `#` `##` `###` を直接書き足しても分割できます。"
)


def import_document(project_name: str, url: str = "", docx_bytes: bytes = b"") -> dict:
    """Google ドキュメント URL か .docx から Markdown を作り、プロジェクトに保存する。

    返り値: {"markdown", "title", "headings", "images", "warning"}
    """
    if docx_bytes:
        parsed = parse_docx(docx_bytes)
    else:
        parsed = parse_google_doc(url)

    markdown = document_to_markdown(project_name, parsed)
    markdown_path(project_name).write_text(markdown, encoding="utf-8")

    # 取り込みで内容が変わったので、前の内容で作られたスライド画像は破棄する。
    # 残しておくと「未生成」表示なのに古い画像が PPTX に入る食い違いが起きる。
    for png in project_dir(project_name).glob("slide_*.png"):
        if re.fullmatch(r"slide_\d+\.png", png.name):
            png.unlink()

    cfg = read_project_config(project_name)
    cfg["source"] = {"type": "gdoc" if url else "docx", "url": url}
    write_project_config(project_name, cfg)

    return {
        "markdown": markdown,
        "title": parsed.title,
        "headings": parsed.heading_count,
        "images": parsed.image_count,
        "warning": NO_HEADING_WARNING if parsed.heading_count == 0 else "",
    }


def pptx_path(project_name: str) -> Path:
    return project_dir(project_name) / f"{project_name}.pptx"


def read_markdown(project_name: str) -> str:
    p = markdown_path(project_name)
    return p.read_text(encoding="utf-8") if p.exists() else ""


def read_prompt(project_name: str) -> str:
    p = prompt_path(project_name)
    return p.read_text(encoding="utf-8") if p.exists() else DEFAULT_PROMPT


def _client() -> OpenAI:
    return OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))


# ---------------------------------------------------------------------------
# ① Markdown → スライド構成（見出し1/2/3ごとに1枚）
# ---------------------------------------------------------------------------

_HEADING_RE = re.compile(r"^(#{1,3})\s+(.*)$")

# Markdown の画像参照 `![alt](path)`。行頭の箇条書き記号は許容する。
_IMAGE_RE = re.compile(r"^(?:[-*+]\s+)?!\[[^\]]*\]\(\s*([^)\s]+)(?:\s+\"[^\"]*\")?\s*\)$")


# ---------------------------------------------------------------------------
# 引用文献スライド
#
# どのテンプレートでも「発表の最後は一次ソースの引用リストで終える」を守る。
# まとめスライドではなく、読んだ人がその論文に必ずたどり着ける情報を並べる。
#
# 引用文献だけは「要約してキーワードにする」対象から外す。文献番号・著者・雑誌名・
# 巻・号・ページ・年は、まとめた瞬間に引用として成立しなくなるため、原文のまま
# 1件1行で並べる（このスライドだけ小さい文字を許す）。
# ---------------------------------------------------------------------------

# この見出しのスライドは逐語リストとして扱う
_REFERENCES_TITLE_RE = re.compile(
    r"(引用|参考)\s*(文献|資料)|文献\s*(一覧|リスト)|出典|references?|bibliography",
    re.IGNORECASE,
)

# 引用1件に必ず含める項目。これが揃っていれば論文を一意に特定できる。
REFERENCE_REQUIRED_FIELDS = "番号・著者名（1人＋et al. で可）・雑誌名・巻（号）・ページ・年"

# 台本に書いてもらう1件の形。UI のヘルプとプロンプトの両方でこれを見せる。
REFERENCE_LINE_EXAMPLE = (
    "1／R. L. Diercks, M. Stevens, J. Shoulder Elbow Surg. 13, 499–502 (2004)."
)

# 1枚に載せる文献の上限。超えたら均等に分割する
MAX_REFERENCES_PER_SLIDE = 14

# 「13／著者..., 雑誌 巻, ページ (年).／URL／論文要約：…」のような行から、
# 先頭の文献番号を取り出す（「13.」「13)」「13：」などの区切りも許す）
_REFERENCE_NUMBER_RE = re.compile(r"^\(?(\d{1,3})\s*[／/.)\]：:、．]\s*")

# 台本には書誌情報のあとに URL や要約メモが続く。スライドに載せるのは書誌情報だけ。
_REFERENCE_SPLIT_RE = re.compile(r"[／|]")


def _reference_citation(line: str) -> str:
    """文献1行から、スライドに載せる書誌情報だけを取り出す。

    台本の1行は「番号／書誌情報／URL／論文要約：…／今回引用すべき理由：…」の形で、
    URL 以降は発表者向けのメモ。スライドに出すのは書誌情報のところだけなので、
    区切り記号で分けて最初のひとかたまりを採る。区切りが無い行はそのまま使う。
    """
    parts = [p.strip() for p in _REFERENCE_SPLIT_RE.split(line) if p.strip()]
    for part in parts:
        if part.isdigit() or part.startswith("http"):
            continue
        # 区切りが無い行では URL の手前で切る
        return re.split(r"\s*https?://", part)[0].strip()
    return line.strip()


def _reference_entries(raw_lines: list[str]) -> list[str]:
    """引用文献スライドの本文から「12. 著者..., 雑誌 巻, ページ (年).」の行を作る。

    元の台本に文献番号があればその番号を使う（他スライドの本文から
    「文献12」と参照されるため、番号を振り直してはいけない）。
    番号が無い行は、直前の番号の続きとして通し番号を振る。
    """
    entries: list[str] = []
    number = 0
    for line in raw_lines:
        text = re.sub(r"^[-*+]\s+", "", line.strip())
        if not text:
            continue
        m = _REFERENCE_NUMBER_RE.match(text)
        if m:
            number = int(m.group(1))
        else:
            number += 1
        citation = _reference_citation(text)
        if citation:
            entries.append(f"{number}. {citation}")
    return entries


# 「(2004)」「（2004）」。書誌情報の末尾に来る発行年。
_CITATION_YEAR_RE = re.compile(r"[(（]\s*(\d{4})\s*[)）]")

# 年の手前に来る「巻（号）, ページ」。ページは 499–502 / 546 / e333–e339 のいずれも許す。
_CITATION_VOLUME_PAGE_RE = re.compile(
    r"\d{1,4}\s*(?:[(（]\s*[^)）]{1,16}\s*[)）])?\s*[,，]\s*"
    r"[A-Za-z]?\d+(?:\s*[-–—~〜]\s*[A-Za-z]?\d+)?\s*[.,，]?\s*$"
)


def reference_field_gaps(entry: str) -> list[str]:
    """引用1件から、論文を特定するのに足りない項目を挙げる（空なら充足）。

    「番号. 著者…, 雑誌名 巻(号), ページ (年).」という並びを後ろから崩していく。
    年 →「巻（号）・ページ」→ 残りをカンマで分けて「著者名」「雑誌名」の2つ以上、
    という順に確かめる。著者名と雑誌名の境目は機械的には決められないので、
    カンマ区切りが2つ以上あるかどうかだけを見る。

    論文以外の一次ソース（書籍・ガイドライン等）は巻・ページを持たないことがあるため、
    ここで挙げるのは「足りない項目」であってエラーではない。
    """
    text = _REFERENCE_NUMBER_RE.sub("", (entry or "").strip())
    if not text:
        return ["著者名", "雑誌名", "巻（号）・ページ", "年"]

    gaps: list[str] = []
    head = text
    year = _CITATION_YEAR_RE.search(text)
    if year:
        head = text[:year.start()]
    else:
        gaps.append("年")

    head = head.strip().rstrip(".,，、")
    volume_page = _CITATION_VOLUME_PAGE_RE.search(head)
    if volume_page:
        head = head[:volume_page.start()].strip().rstrip(".,，、")
    else:
        gaps.append("巻（号）・ページ")

    if len([c for c in re.split(r"[,，]", head) if c.strip()]) < 2:
        gaps.append("著者名・雑誌名")
    return gaps


def _references_last(slides: list[dict]) -> list[dict]:
    """引用文献スライドを必ず末尾へ回す。

    台本が「引用文献 → まとめ」の順で書かれていても、発表の最後は一次ソースの
    一覧で終える。文献スライド同士・それ以外同士の並び順は変えない。
    """
    refs = [s for s in slides if s.get("references")]
    if not refs or slides[-len(refs):] == refs:
        return slides
    return [s for s in slides if not s.get("references")] + refs


def reference_check(plan: list[dict]) -> dict:
    """最後のスライドが一次ソースの引用リストとして成立しているか調べる。

    返り値の gaps は「何件目のどの項目が足りないか」。台本を直すのは人間なので、
    足りない項目を名指しして返す（ここで勝手に補うと存在しない文献になる）。
    """
    ref_slides = [(i, s) for i, s in enumerate(plan) if s.get("references")]
    if not plan:
        return {"ok": True, "count": 0, "gaps": [], "message": ""}

    if not ref_slides:
        return {
            "ok": False,
            "count": 0,
            "gaps": [],
            "message": (
                "最後のスライドが引用文献リストになっていません。台本の末尾に"
                "「# 引用文献」の見出しを作り、1行1件で"
                f"{REFERENCE_REQUIRED_FIELDS}を並べてください"
                f"（例: {REFERENCE_LINE_EXAMPLE}）。"
            ),
        }

    gaps = []
    count = 0
    for index, slide in ref_slides:
        for entry in slide.get("references") or []:
            count += 1
            missing = reference_field_gaps(entry)
            if missing:
                gaps.append({"index": index, "entry": entry, "missing": missing})

    if gaps:
        message = (
            f"引用文献 {count} 件のうち {len(gaps)} 件で、論文を特定できる情報が"
            f"足りません（{REFERENCE_REQUIRED_FIELDS}）。論文以外の一次ソースなら"
            "そのままで構いません。"
        )
    else:
        message = ""
    return {"ok": not gaps, "count": count, "gaps": gaps, "message": message}


def _split_reference_slides(slides: list[dict]) -> list[dict]:
    """文献が多すぎる引用文献スライドを、読める件数まで複数枚に分ける。

    要約しない＝件数がそのまま行数になるので、分割しないと文字が潰れる。
    """
    result: list[dict] = []
    for slide in slides:
        entries = slide.get("references") or []
        if len(entries) <= MAX_REFERENCES_PER_SLIDE:
            result.append(slide)
            continue

        pages = -(-len(entries) // MAX_REFERENCES_PER_SLIDE)  # 切り上げ
        per_page = -(-len(entries) // pages)  # 件数を均等に配る
        chunks = [entries[i:i + per_page] for i in range(0, len(entries), per_page)]

        base_title = slide.get("title") or "引用文献"
        for page, chunk in enumerate(chunks):
            part = dict(slide)
            part["title"] = f"{base_title}（{page + 1}/{len(chunks)}）"
            part["references"] = chunk
            part["lines"] = list(chunk)
            part["script"] = "\n".join(chunk)
            part["content"] = [{"kind": "text", "text": c} for c in chunk]
            # 画像は1枚目にだけ残す（引用文献に図が付くことはまず無いが、
            # 全ページに同じ画像を添付して重複生成させない）
            if page > 0:
                part["images"] = []
                part["image_notes"] = []
            result.append(part)
    return result


def _new_slide(title: str = "") -> dict:
    return {
        "title": title, "content": [], "lines": [], "images": [],
        "illustration": "", "script": "", "_raw": [],
    }


def _finalize_slide(slide: dict) -> dict:
    """content（本文と画像の並び）から lines / images / script を導出する。

    lines・images は「テキストだけ」「画像だけ」を扱う既存処理向けの平坦な形。
    どの本文の間に画像が挟まっていたかという文脈は content 側だけが持つ。
    script は PPTX のスピーカーノートに入れる台本原文。スライドには要約された
    キーワードしか出ないため、読み上げる元の文章を別に持っておく。
    """
    slide["lines"] = [c["text"] for c in slide["content"] if c["kind"] == "text"]
    slide["images"] = [c["path"] for c in slide["content"] if c["kind"] == "image"]

    # 台本は原文のまま（箇条書き記号や番号も落とさない）。空行は段落の区切りとして
    # 1つだけ残し、先頭・末尾の空行は落とす。
    raw = slide.pop("_raw", [])
    lines: list[str] = []
    for line in raw:
        if line or (lines and lines[-1]):
            lines.append(line)
    while lines and not lines[-1]:
        lines.pop()
    slide["script"] = "\n".join(lines)

    # 引用文献スライドは要約せず原文のまま並べる。番号や記号を落とさないよう、
    # 本文化して記号を削ったあとの lines ではなく原文から組み立てる。
    if _REFERENCES_TITLE_RE.search(slide.get("title") or ""):
        entries = _reference_entries(lines)
        if entries:
            slide["references"] = entries
    return slide


def parse_markdown_to_slides(md: str) -> list[dict]:
    """Markdown を見出し1/2/3ごとに1スライドへ分解する。

    - `#` `##` `###` が現れるたびに新スライド境界
    - h4以降・箇条書き・番号付き・段落は直近スライドの本文に束ねる
    - `![](_doc/img_000.png)` の画像参照も、本文との前後関係を保ったまま
      content に並べる。ドキュメント上で画像が置かれた位置は
      「その画像がどの話題のものか」を表す情報なので、順序を落とさない。
    - 最初の見出しより前に本文がある場合は、無題スライドとして拾う
    - 「引用文献」「参考文献」などの見出しは逐語リスト扱いにして references を
      持たせる（要約させず原文のまま並べる。多すぎる場合は複数枚に分ける）
    - 引用文献スライドは台本での位置にかかわらず末尾へ回す。発表の最後は
      まとめではなく一次ソースの一覧で終える
    """
    slides: list[dict] = []
    current: dict | None = None

    for raw in (md or "").splitlines():
        line = raw.rstrip()
        m = _HEADING_RE.match(line)
        if m:
            if current is not None:
                slides.append(_finalize_slide(current))
            current = _new_slide(m.group(2).strip())
            continue

        stripped = line.strip()
        if not stripped:
            if current is not None:
                current["_raw"].append("")
            continue

        if current is None:
            # 最初の見出し前の本文 → 無題スライドを開始
            current = _new_slide()

        img = _IMAGE_RE.match(stripped)
        if img:
            current["content"].append({"kind": "image", "path": img.group(1)})
            continue

        # スピーカーノート用に原文をそのまま控える（記号を落とす前）
        current["_raw"].append(stripped)

        # 見出し4以降・箇条書き・番号付きの記号を除去して本文化
        stripped = re.sub(r"^#{4,}\s*", "", stripped)
        stripped = re.sub(r"^[-*+]\s+", "", stripped)
        stripped = re.sub(r"^\d+[.)]\s+", "", stripped)
        if stripped:
            current["content"].append({"kind": "text", "text": stripped})

    if current is not None:
        slides.append(_finalize_slide(current))

    slides = _references_last(_split_reference_slides(slides))
    for slide in slides:
        if slide.get("references"):
            # 足りない項目をプレビューで1件ずつ指摘できるよう、行と同じ並びで持たせる
            slide["reference_gaps"] = [
                reference_field_gaps(e) for e in slide["references"]
            ]
    return slides


class CondenseWarning(Exception):
    """要約だけ失敗したことを伝える（構成案は plan に入っていて使える）。"""

    def __init__(self, message: str, plan: list):
        super().__init__(message)
        self.plan = plan


def generate_slide_plan(project_name: str, condense: bool = False) -> list[dict]:
    """保存済み Markdown から構成案を生成して plan.json に保存する。

    condense=True のときはテキストLLMで本文をキーワードへ絞り込む。
    要約に失敗しても構成案そのものは失われないよう、素の構成案を土台にして
    要約結果をマージする（失敗時は本文がそのまま素材として使われる）。
    """
    md = read_markdown(project_name)
    plan = parse_markdown_to_slides(md)

    condense_error = ""
    if condense and plan:
        try:
            plan = condense_svc.condense_slides(plan, _client())
        except Exception as e:
            condense_error = str(e)

    plan_path(project_name).write_text(
        json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    if condense_error:
        raise CondenseWarning(condense_error, plan)
    return plan


def load_plan(project_name: str) -> list[dict]:
    p = plan_path(project_name)
    if not p.exists():
        return []
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return []


def save_plan(project_name: str, plan: list) -> None:
    plan_path(project_name).write_text(
        json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ---------------------------------------------------------------------------
# 医用画像ライブラリ由来の素材
#
# ライブラリの図は「検証済みの図解そのもの」なので、AIに別の図解を描かせると
# 内容の重複した（そして解剖学的に怪しい）イラストが並んでしまう。どの素材が
# ライブラリ由来かを覚えておき、そのスライドでは illustration の指示を出さない。
# ---------------------------------------------------------------------------

def library_materials(slide: dict) -> dict:
    """スライドに割り当てられた素材のうち、ライブラリ由来のもの（k -> 相対パス）。"""
    v = slide.get("library_images")
    return v if isinstance(v, dict) else {}


def set_library_material(project_name: str, index: int, k: int, rel_path: str) -> None:
    plan = load_plan(project_name)
    if not (0 <= index < len(plan)):
        return
    mapping = dict(library_materials(plan[index]))
    mapping[str(k)] = rel_path
    plan[index]["library_images"] = mapping
    save_plan(project_name, plan)


def clear_library_material(project_name: str, index: int, k: int) -> None:
    plan = load_plan(project_name)
    if not (0 <= index < len(plan)):
        return
    mapping = dict(library_materials(plan[index]))
    if mapping.pop(str(k), None) is None:
        return
    if mapping:
        plan[index]["library_images"] = mapping
    else:
        plan[index].pop("library_images", None)
    save_plan(project_name, plan)


# ---------------------------------------------------------------------------
# ② 1スライドの画像生成
# ---------------------------------------------------------------------------

def _document_flow_block(slide: dict, doc_image_rels: list) -> str:
    """ドキュメント上の本文と画像の並びを、画像番号つきで書き出す。

    ドキュメントで画像が挟まっている位置は「その画像がどの話題のものか」を
    表しているので、その前後関係をそのままモデルに渡して文脈を読ませる。
    """
    content = slide.get("content") or []
    if not content or not doc_image_rels:
        return ""

    # 添付した画像だけを番号づけの対象にする（解決できなかったものは飛ばす）
    number_of = {rel: i + 1 for i, rel in enumerate(doc_image_rels)}

    flow = []
    for item in content:
        if item.get("kind") == "image":
            n = number_of.get(item.get("path"))
            if n:
                flow.append(f"［ここに画像{n}が置かれている］")
        else:
            text = (item.get("text") or "").strip()
            if text:
                flow.append(f"本文: {text}")

    if not any(line.startswith("［") for line in flow):
        return ""

    return (
        "\n【ドキュメント上の並び順（各画像がどの文脈に置かれていたか）】\n"
        "元のドキュメントでは、本文と画像が次の順で並んでいました。\n"
        "画像はこの位置に貼られている＝直前後の本文の内容を示すものです。\n"
        "各画像を、対応する話題のキーワードの近くに配置し、"
        "関係のない項目の横に置かないでください。\n"
        + "\n".join(flow)
        + "\n\n"
    )


def _images_block(
    has_reference: bool, doc_image_rels: list, image_notes: list, material_count: int,
    library_positions: list = None,
) -> str:
    """添付画像の役割・扱い方の指示を組み立てる。"""
    doc_count = len(doc_image_rels)
    use_count = doc_count + material_count
    if use_count <= 0:
        return ""

    if has_reference:
        order_note = (
            "添付画像のうち1枚目はデザイン参照用です。"
            f"それ以降の{use_count}枚が、このスライドに必ず盛り込む「使用したい画像」で、"
            "添付された順に「画像1」「画像2」…と呼びます。\n"
        )
    else:
        order_note = (
            f"添付した{use_count}枚が、このスライドに必ず盛り込む「使用したい画像」で、"
            "添付された順に「画像1」「画像2」…と呼びます。\n"
        )

    # 要約時に作った「その画像が何を示しているか」のメモを番号つきで添える
    notes = ""
    for i in range(doc_count):
        note = (image_notes[i] if i < len(image_notes) else "") or ""
        note = str(note).strip()
        if note:
            notes += f"・画像{i + 1}: {note}\n"
    # ライブラリ由来（＝医学的に検証済みの図）と、手動アップロードの素材を区別して伝える。
    # 前者はそのスライドの図解そのものなので、別の図を描き足させない。
    library_positions = library_positions or []
    lib_nums = [doc_count + j + 1 for j in library_positions]
    other_nums = [doc_count + j + 1 for j in range(material_count) if j not in library_positions]

    if lib_nums:
        notes += (
            f"・画像{('・'.join(str(n) for n in lib_nums))}: "
            "医学的に検証済みの解剖図。**このスライドの図解はこれ**なので、"
            "内容を描き変えず、そのまま主役として大きく配置する。\n"
        )
    if other_nums:
        label = "・".join(str(n) for n in other_nums)
        notes += f"・画像{label}: 手動で追加された素材画像。同じく主役として扱う。\n"

    return (
        "\n【使用したい画像・最優先で反映／厳守】\n"
        + order_note
        + (f"\n{notes}" if notes else "")
        + "\n〈画像は絶対に切り取らない／最重要・厳守〉\n"
        "・渡された画像は必ず『全体』をそのまま使うこと。\n"
        "・重要そうな部分を判断して切り出す、拡大する、寄せる（クロップ・ズーム・"
        "トリミング）ことは一切しないこと。これらの画像は既に必要な範囲だけに"
        "整えてあるので、こちらで判断して切り取る必要はまったくない。\n"
        "・画像の上下左右を削らない。四隅・端まで全部が見えるように配置する。\n"
        "・縦横比を変えて引き伸ばす・押し潰すこともしない"
        "（元の縦横比を保ったまま拡大縮小する）。\n"
        "・画像の端を、フレーム・帯・カード・文字などで覆って隠さないこと。\n"
        "・写っているもの・構図・色味・画像内の文字を描き変えて別物にしないこと。"
        "画像内の文字は読み取れればそのまま保ち、読み取れない場合は捏造せず"
        "文字なしで表現する。\n"
        "\n〈大きさと配置〉\n"
        "・すべての画像をスライド内に配置し、主役として大きく目立たせること"
        "（全体で目安としてスライド面積の3〜5割）。\n"
        "・画像が複数ある場合は、全体が見える大きさを保ったまま整列させて並べる"
        "（重ねて隠さない）。\n"
        "・見出しやキーワードは画像の周囲に添える形にし、画像の上に重ねて"
        "見えなくしないこと。\n"
        "・画像を極端に小さくしたり、隅に追いやったりしないこと。\n"
    )


_DESIGN_REFERENCE_BLOCK = (
    "添付した参照画像は「デザイン見本」です。配色・フォントの雰囲気・"
    "レイアウトの方向性・装飾のテイストだけを参考にしてください。\n"
    "【参照画像の文字は絶対に使わない／最重要・厳守】\n"
    "・参照画像に写っている見出し・本文・キャプション・ロゴ・数字などの"
    "文字情報は一切読み取らず、生成するスライドに転記・引用・再現しないこと。\n"
    "・参照画像はあくまで見た目のお手本であり、内容（トピック・文言）は"
    "このスライドとは無関係だとみなすこと。\n"
    "・スライドに描画してよい文字は、下記で指定する見出しとテキスト素材から"
    "取ったものだけ。それ以外の言葉を勝手に足さない。\n"
)


def _build_references_prompt(
    project_prompt: str, slide: dict, slide_no: int, total: int, fix: str = "",
    has_reference: bool = False,
) -> str:
    """引用文献スライドのプロンプト（要約せず、1件1行でそのまま並べる）。

    他のスライドは「キーワードだけを大きく」が大原則だが、引用文献はその逆で、
    書誌情報を1文字も変えずに全件載せることが目的。共通の大原則をそのまま
    渡すと必ずまとめられてしまうため、別のプロンプトを組み立てる。
    """
    title = slide.get("title") or "引用文献"
    entries = [str(e).strip() for e in (slide.get("references") or []) if str(e).strip()]
    entries_block = "\n".join(entries)

    prompt = (
        f"これはプレゼンテーションのスライド（{slide_no}/{total}枚目）で、"
        "発表の最後に置く「引用文献リスト（一次ソース一覧）」です。\n"
        "16:9 のスライド1枚として、文献リストをきれいに組版してください。\n\n"
        "【このスライドだけの大原則／最重要・厳守】\n"
        "・これは学術的な引用文献の一覧です。要約・キーワード化・グループ化・"
        "言い換えは一切禁止。下に並べた行を、書かれているとおりに転記する。\n"
        f"・文献は全部で{len(entries)}件。{len(entries)}行すべてを1枚に載せる。"
        "省略・抜粋・「ほか」「etc.」でのまとめは禁止。\n"
        f"・各行には{REFERENCE_REQUIRED_FIELDS}が入っている。聞き手がこの1行だけで"
        "原著にたどり着けることが目的なので、行末の巻・号・ページ・年まで"
        "1つも落とさずに描く（長いからといって途中で切らない）。\n"
        "・著者名・雑誌名・巻・号・ページ・年・記号（. , ; – ( )）を1文字も変えない。"
        "欧文はアルファベットのまま描き、日本語に訳さない。\n"
        "・存在しない文献・番号・年を絶対に作り足さない。行を入れ替えない。\n"
        "・このスライドに限り小さめの文字でよい（引用情報は例外）。"
        "ただし全行を同じ大きさ・同じ書体で、はっきり読める濃さで描く。\n\n"
    )
    if has_reference:
        prompt += _DESIGN_REFERENCE_BLOCK + "\n"
    prompt += (
        "【レイアウト】\n"
        "・上部に見出しを大きく置き、その下に文献を1件1行で縦に並べる。\n"
        "・行頭の番号は丸バッジなどで揃え、本文の左端も縦にきれいに揃える。\n"
        "・1件は必ず1行に収める（折り返さない）。長い行が入らない場合は、"
        "文字を少し小さくして幅を合わせる。\n"
        "・右下の角にはロゴ用の空きが必要なので、下の方の行はそこに掛からない"
        "長さで組む（背景のまま残す）。\n"
        "・行間を詰めすぎず、上下左右の端に文字を接触させない。\n"
        "・カードや吹き出しで1件ずつ囲まない（リストとして淡々と並べる）。\n"
        "・右側に余白ができた場合は、本・書類・虫めがねなどの小さな装飾イラストを"
        "文字に重ならない位置に添えてよい。人体図や解説用の図解は描かない。\n\n"
        f"【スライド見出し（正確に・誤字なく日本語で描画）】\n{title}\n\n"
        f"【文献リスト（この{len(entries)}行を、この順で、この通りに描画）】\n"
        f"{entries_block}\n\n"
    )
    if project_prompt:
        prompt += (
            f"【デザイン・トーンの指示（配色・装飾のみ適用）】\n{project_prompt}\n"
            "※文字数を減らす・キーワード中心にする・文字を大きくするといった指示が"
            "含まれていても、このスライドでは適用しない（上の大原則が優先）。\n\n"
        )
    if fix.strip():
        prompt += f"【修正指示（最優先で反映）】\n{fix.strip()}\n"

    return prompt[:30000]


def _build_image_prompt(
    project_prompt: str, slide: dict, slide_no: int, total: int, fix: str = "",
    has_reference: bool = False, doc_image_rels: list = None, material_count: int = 0,
    library_positions: list = None,
) -> str:
    if slide.get("references"):
        return _build_references_prompt(
            project_prompt, slide, slide_no, total, fix=fix, has_reference=has_reference,
        )

    title = slide.get("title", "")
    lines = slide.get("lines", []) or []
    keywords = [k for k in (slide.get("keywords") or []) if str(k).strip()]
    illustration = (slide.get("illustration") or "").strip()
    doc_image_rels = doc_image_rels or []
    image_notes = slide.get("image_notes") or []

    lines_block = "\n".join(f"・{l}" for l in lines)
    keywords_block = "\n".join(f"・{k}" for k in keywords)

    prompt = (
        f"これはプレゼンテーションのスライド（{slide_no}/{total}枚目）です。\n"
        "16:9 のスライド1枚として、見やすく洗練されたレイアウトでデザインしてください。\n\n"
        "【最重要・全スライド共通の大原則／厳守】\n"
        "・文章ではなくキーワード／短いフレーズ中心にする。"
        "提供されたテキストが長くても、そのまま全文を書き写さず、"
        "要点となる単語・数字・短い言い回しだけを抜き出して載せること。\n"
        "・文字は大きく、はっきりと。小さな文字で長い文章を詰め込むのは絶対に禁止。\n"
        "・1スライドの情報量は思い切って絞り、余白・図版・アイコンで直感的に理解させる。\n"
        "・高齢者でも一目で読めるフォントサイズを最優先する"
        "（出典・引用などの補足情報のみ小さくてよい）。\n\n"
    )
    if has_reference:
        prompt += _DESIGN_REFERENCE_BLOCK
    prompt += _images_block(
        has_reference, doc_image_rels, image_notes, material_count, library_positions
    )
    prompt += _document_flow_block(slide, doc_image_rels)
    prompt += (
        "【余白ルール・重要】見出し・本文・図版など、すべての要素を画像の内側に"
        "余裕をもって収め、上下左右の端に文字を接触させないでください。\n\n"
        f"【スライド見出し（正確に・誤字なく日本語で描画）】\n{title}\n\n"
    )
    if keywords_block:
        # 要約済み（キーワード確定）。ここに無い言葉を足させない。
        prompt += (
            "【スライドに描くキーワード（これだけを大きな文字で描く／厳守）】\n"
            "↓この項目を、そのままの文言で大きく読みやすく配置する。\n"
            "・ここに書かれていない文章・説明文・補足を勝手に足さないこと。\n"
            "・各項目を長い文章に膨らませないこと。日本語は誤字なく正確に。\n"
            f"{keywords_block}\n\n"
        )
    elif lines_block:
        prompt += (
            "【本文の素材（このテキストはそのまま全文を描画しないこと）】\n"
            "↓ここからキーワード・要点だけを抜き出し、短く大きな文字で表現する。"
            "長い文章のまま小さく書かない。日本語は誤字なく。\n"
            f"{lines_block}\n\n"
        )
    if library_positions:
        # 検証済みの解剖図を渡してあるスライド。要約時に作られた図解指示をそのまま
        # 使うと、渡した図と内容が重複する別のイラストが並んで描かれてしまう。
        prompt += (
            "【図解を新しく描かないこと／厳守】\n"
            "・このスライドの図解は、添付した解剖図がその役割を果たします。\n"
            "・添付図とは別に、イラスト・模式図・断面図・人体図・アイコン化した図解などを"
            "新たに描き起こして並べないこと。内容が重複し、医学的にも不正確になります。\n"
            "・図が足りないと感じても描き足さず、余白として残すか、文字の配置で埋めること。\n"
            "・キーワードの先頭に添える小さなアイコン程度は可。\n\n"
        )
    elif illustration:
        prompt += f"【添えるイラスト・図版】\n{illustration}\n（図版内に余計な説明文は入れない）\n\n"
    if project_prompt:
        prompt += f"【デザイン・トーンの指示】\n{project_prompt}\n\n"
    if fix.strip():
        prompt += f"【修正指示（最優先で反映）】\n{fix.strip()}\n"

    return prompt[:30000]


def _fit_to_16x9(img: Image.Image) -> Image.Image:
    """生成画像(3:2)を 16:9 に「縦圧縮でフィット」させる（クロップしない）。

    クロップだと最上端の見出し・最下端の要素のどちらかが切れるため、
    上下を一切切らずに縦方向へ軽く圧縮して 16:9 に収める。
    """
    img = img.convert("RGB")
    return img.resize((TARGET_W, TARGET_H), Image.LANCZOS)


def _sample_background_color(img: Image.Image) -> tuple[int, int, int]:
    """画像の端付近（上辺・左右の内側）から背景色を推定する。"""
    w, h = img.size

    def avg(x0: float, y0: float, x1: float, y1: float) -> tuple[int, int, int]:
        box = (int(x0 * w), int(y0 * h), int(x1 * w), int(y1 * h))
        px = img.crop(box).resize((1, 1), Image.LANCZOS).getpixel((0, 0))
        return px[:3]

    samples = [
        avg(0.10, 0.02, 0.90, 0.05),   # 上辺の内側
        avg(0.015, 0.10, 0.03, 0.90),  # 左辺の内側
        avg(0.97, 0.10, 0.985, 0.90),  # 右辺の内側
    ]
    return tuple(sum(s[i] for s in samples) // len(samples) for i in range(3))


def _apply_negative_space(img: Image.Image, ns: dict) -> Image.Image:
    """右下のネガティブスペースを背景色で軽く均し、きれいな余白に仕上げる。

    外周フレームは inset 分だけ残し、内側の矩形だけをぼかしマスクで合成する。
    """
    img = img.convert("RGB")
    w, h = img.size
    right = float(ns.get("right", 0.15))
    bottom = float(ns.get("bottom", 0.20))
    inset = float(ns.get("inset", 0.014))
    feather = float(ns.get("feather", 0.018))
    color = tuple(ns["color"]) if ns.get("color") else _sample_background_color(img)

    margin = int(inset * w)
    x0 = int((1 - right) * w)
    y0 = int((1 - bottom) * h)
    x1 = w - margin
    y1 = h - margin
    radius = int(min(w, h) * 0.016)

    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle([x0, y0, x1, y1], radius=radius, fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(max(1, int(feather * w))))

    fill = Image.new("RGB", (w, h), color)
    return Image.composite(fill, img, mask)


def generate_slide_image(
    project_name: str,
    index: int,
    slide: dict,
    total: int,
    project_prompt: str,
    fix: str = "",
) -> Path:
    """1スライドを GPT Image 2.0 で生成し、16:9 にフィットして保存する。

    参照画像（デザイン見本）、ドキュメントから取り込んだ画像、スライドに紐づく
    素材画像があれば、まとめて images.edit に渡す。ドキュメント画像は構成案の
    images（Markdown の画像参照）から、素材画像はプロジェクトフォルダから
    自動で拾うため、単体再生成・一括再生成のどちらでも反映される。
    """
    client = _client()
    ref = reference_path(project_name)
    has_ref = ref.exists()

    # ドキュメント画像と素材画像を、API が一度に受け取れる枚数まで詰める。
    # 参照画像もこの枚数に含まれるため、その分を先に引いておく。
    budget = MAX_ATTACHED_IMAGES - (1 if has_ref else 0)

    doc_image_rels: list[str] = []
    doc_images: list[Path] = []
    for rel in slide.get("images") or []:
        if len(doc_images) >= budget:
            break
        p = resolve_doc_image(project_name, rel)
        if p is not None:
            doc_image_rels.append(rel)
            doc_images.append(p)

    material_ks = list_material_indexes(project_name, index)[: budget - len(doc_images)]
    materials = [material_path(project_name, index, k) for k in material_ks]

    # 添付順のうち、どれがライブラリ由来の解剖図かをプロンプト側に伝える
    lib_map = library_materials(slide)
    library_positions = [j for j, k in enumerate(material_ks) if str(k) in lib_map]

    prompt = _build_image_prompt(
        project_prompt, slide, index + 1, total, fix=fix,
        has_reference=has_ref,
        doc_image_rels=doc_image_rels,
        material_count=len(materials),
        library_positions=library_positions,
    )

    # 添付順: [参照画像(あれば)] + ドキュメント画像 + 素材画像。
    # プロンプト側の「画像1」「画像2」という呼び方とこの順序を必ず一致させる。
    image_paths: list[Path] = ([ref] if has_ref else []) + doc_images + materials

    if image_paths:
        handles = [open(p, "rb") for p in image_paths]
        try:
            result = client.images.edit(
                model=SLIDE_IMAGE_MODEL,
                image=handles,
                prompt=prompt,
                size=GEN_SIZE,
            )
        finally:
            for h in handles:
                h.close()
    else:
        result = client.images.generate(
            model=SLIDE_IMAGE_MODEL,
            prompt=prompt,
            size=GEN_SIZE,
        )

    b64 = result.data[0].b64_json
    img_bytes = base64.b64decode(b64)

    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    img = _fit_to_16x9(img)

    # プロジェクト設定に右下ネガティブスペースがあれば後処理で均す
    ns = read_project_config(project_name).get("negative_space")
    if ns:
        img = _apply_negative_space(img, ns)

    out = slide_png_path(project_name, index)
    img.save(out, "PNG")
    return out


# ---------------------------------------------------------------------------
# ③ PPTX 組立
# ---------------------------------------------------------------------------

def slide_note(slide: dict) -> str:
    """そのスライドに対応する台本原文（スピーカーノートに入れる本文）を返す。

    script を持たない古い plan.json は lines から組み立てる。
    """
    script = (slide.get("script") or "").strip()
    if script:
        return script
    return "\n".join(slide.get("lines") or []).strip()


def build_pptx(project_name: str, num_slides: int, with_notes: bool = True) -> Path:
    """slide_NN.png を 1枚1スライドの 16:9 PPTX に組み立てる。

    with_notes=True のとき、各スライドのノート欄に台本の該当部分を入れる。
    スライドに描かれるのは要約されたキーワードだけなので、発表時に読み上げる
    元の文章がノートに無いと、結局ドキュメントを別に開くことになる。
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    plan = load_plan(project_name) if with_notes else []

    for i in range(num_slides):
        png = slide_png_path(project_name, i)
        if not png.exists():
            continue
        slide = prs.slides.add_slide(blank_layout)
        slide.shapes.add_picture(
            str(png), 0, 0,
            width=prs.slide_width,
            height=prs.slide_height,
        )
        if with_notes and i < len(plan):
            note = slide_note(plan[i])
            if note:
                slide.notes_slide.notes_text_frame.text = note

    out = pptx_path(project_name)
    prs.save(str(out))
    return out
