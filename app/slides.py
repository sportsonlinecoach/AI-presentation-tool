"""
スライド生成サービス層

方針:
1. Markdown の見出し1/2/3ごとに1スライドへ分解（決定論パーサ・LLM不要）
   - h4以降・箇条書き・段落は直近スライドの本文に束ねる
2. 各スライドを GPT Image 2.0 で1枚絵として生成
   - プロジェクトの見本画像があれば images.edit の参照に使う（無ければ images.generate）
   - プロジェクトごとの生成プロンプトをスタイル指示に流用
   - 任意で「修正指示(fix)」をプロンプト末尾に追加して再生成
3. PNG 群を python-pptx で「1枚画像=1スライド（16:9 全面）」の PPTX に組立

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

# 1スライドに追加できる素材画像の上限
MAX_MATERIALS_PER_SLIDE = 4

# プロジェクト内の固定ファイル名
MARKDOWN_FILE = "markdown.md"
PROMPT_FILE = "_prompt.txt"
REFERENCE_FILE = "_reference.png"
PLAN_FILE = "plan.json"
CONFIG_FILE = "_config.json"


# ---------------------------------------------------------------------------
# プロンプトテンプレート（新規作成時に選択・編集できる雛形）
# ---------------------------------------------------------------------------

PROMPT_TEMPLATES: dict[str, str] = {
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


def template_reference_path(template_name: str) -> Path | None:
    """テンプレートに同梱されたデフォルト見本画像のパスを返す（無ければ None）。"""
    fname = TEMPLATE_REFERENCES.get(template_name or "")
    if not fname:
        return None
    p = TEMPLATE_ASSETS_DIR / fname
    return p if p.exists() else None


def templates_with_reference() -> list[str]:
    """デフォルト見本画像を持つテンプレート名の一覧。"""
    return [name for name in TEMPLATE_REFERENCES if template_reference_path(name)]


def template_negative_space(template_name: str) -> dict | None:
    """テンプレートに定義された右下ネガティブスペース設定を返す（無ければ None）。"""
    ns = TEMPLATE_NEGATIVE_SPACE.get(template_name or "")
    return dict(ns) if ns else None


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


def parse_markdown_to_slides(md: str) -> list[dict]:
    """Markdown を見出し1/2/3ごとに1スライドへ分解する。

    - `#` `##` `###` が現れるたびに新スライド境界
    - h4以降・箇条書き・番号付き・段落は直近スライドの本文(lines)に束ねる
    - 最初の見出しより前に本文がある場合は、無題スライドとして拾う
    """
    slides: list[dict] = []
    current: dict | None = None

    for raw in (md or "").splitlines():
        line = raw.rstrip()
        m = _HEADING_RE.match(line)
        if m:
            if current is not None:
                slides.append(current)
            current = {"title": m.group(2).strip(), "lines": [], "illustration": ""}
            continue

        stripped = line.strip()
        if not stripped:
            continue

        if current is None:
            # 最初の見出し前の本文 → 無題スライドを開始
            current = {"title": "", "lines": [], "illustration": ""}

        # 見出し4以降・箇条書き・番号付きの記号を除去して本文化
        stripped = re.sub(r"^#{4,}\s*", "", stripped)
        stripped = re.sub(r"^[-*+]\s+", "", stripped)
        stripped = re.sub(r"^\d+[.)]\s+", "", stripped)
        if stripped:
            current["lines"].append(stripped)

    if current is not None:
        slides.append(current)

    return slides


def generate_slide_plan(project_name: str) -> list[dict]:
    """保存済み Markdown から構成案を生成して plan.json に保存する。"""
    md = read_markdown(project_name)
    plan = parse_markdown_to_slides(md)
    plan_path(project_name).write_text(
        json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return plan


def load_plan(project_name: str) -> list[dict]:
    p = plan_path(project_name)
    if not p.exists():
        return []
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return []


# ---------------------------------------------------------------------------
# ② 1スライドの画像生成
# ---------------------------------------------------------------------------

def _build_image_prompt(
    project_prompt: str, slide: dict, slide_no: int, total: int, fix: str = "",
    has_reference: bool = False, material_count: int = 0,
) -> str:
    title = slide.get("title", "")
    lines = slide.get("lines", []) or []
    illustration = (slide.get("illustration") or "").strip()

    lines_block = "\n".join(f"・{l}" for l in lines)

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
        prompt += (
            "添付した参照画像は「デザイン見本」です。配色・フォントの雰囲気・"
            "レイアウトの方向性・装飾のテイストだけを参考にしてください。\n"
            "【参照画像の文字は絶対に使わない／最重要・厳守】\n"
            "・参照画像に写っている見出し・本文・キャプション・ロゴ・数字などの"
            "文字情報は一切読み取らず、生成するスライドに転記・引用・再現しないこと。\n"
            "・参照画像はあくまで見た目のお手本であり、内容（トピック・文言）は"
            "このスライドとは無関係だとみなすこと。\n"
            "・スライドに描画してよい文字は、下記で指定する『スライド見出し』と"
            "『本文の素材』から抜き出したものだけ。それ以外の言葉を勝手に足さない。\n"
        )
    if material_count > 0:
        if has_reference:
            order_note = (
                "添付画像のうち最初の1枚はデザイン参照用です。"
                "それ以降に添付された画像が、このスライドに必ず盛り込む「素材画像」です。\n"
            )
        else:
            order_note = "添付した画像が、このスライドに必ず盛り込む「素材画像」です。\n"
        prompt += (
            "\n【素材画像・最優先で反映／厳守】\n"
            + order_note
            + f"・素材画像は全部で{material_count}枚あります。すべてをスライド内に配置してください。\n"
            "・素材画像をこのスライドの主役として、大きく中央付近に目立つように配置し、"
            "その素材画像を中心にレイアウト全体を構成してください。\n"
            "・素材画像の内容（写真・図・人物など）はできるだけ元の見た目・雰囲気を保ったまま使い、"
            "見出しや本文はその周囲に添える形にしてください。\n"
            "・素材画像を極端に小さくしたり、隅に追いやったり、隠したりしないこと。\n"
        )
    prompt += (
        "【余白ルール・重要】見出し・本文・図版など、すべての要素を画像の内側に"
        "余裕をもって収め、上下左右の端に文字を接触させないでください。\n\n"
        f"【スライド見出し（正確に・誤字なく日本語で描画）】\n{title}\n\n"
    )
    if lines_block:
        prompt += (
            "【本文の素材（このテキストはそのまま全文を描画しないこと）】\n"
            "↓ここからキーワード・要点だけを抜き出し、短く大きな文字で表現する。"
            "長い文章のまま小さく書かない。日本語は誤字なく。\n"
            f"{lines_block}\n\n"
        )
    if illustration:
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

    参照画像（デザイン見本）とスライドに紐づく素材画像があれば、まとめて
    images.edit に渡す。素材画像はプロジェクトフォルダから自動で拾うため、
    単体再生成・一括再生成のどちらでも反映される。
    """
    client = _client()
    ref = reference_path(project_name)
    has_ref = ref.exists()
    materials = list_material_paths(project_name, index)
    prompt = _build_image_prompt(
        project_prompt, slide, index + 1, total, fix=fix,
        has_reference=has_ref, material_count=len(materials),
    )

    # 添付画像の順序: [参照画像(あれば)] + 素材画像。プロンプト側の役割説明と一致させる。
    image_paths: list[Path] = ([ref] if has_ref else []) + materials

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

def build_pptx(project_name: str, num_slides: int) -> Path:
    """slide_NN.png を 1枚1スライドの 16:9 PPTX に組み立てる。"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

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

    out = pptx_path(project_name)
    prs.save(str(out))
    return out
