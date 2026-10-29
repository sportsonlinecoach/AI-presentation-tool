"""
ドキュメント取り込み層（Google ドキュメント / .docx → ブロック列）

方針:
1. Google ドキュメントの URL から docx 形式でエクスポートして取得する
   - `https://docs.google.com/document/d/<ID>/export?format=docx` は
     「リンクを知っている全員が閲覧可」の共有設定なら認証なしで取得できる。
   - OAuth を使わないぶん初期設定ゼロで済み、docx には画像が同梱されるため
     期限切れする画像URLを追いかけなくてよい（Docs API の contentUri は期限切れする）。
   - 非公開ドキュメントの場合は Google がログインHTMLを返すので、それを検出して
     「共有設定を変える」か「.docx を直接アップロードする」よう案内する。
2. docx を「文書順のブロック列」に分解する
   - 見出し1〜3・本文・箇条書き・画像を、元ドキュメントに出てくる順で並べる。
   - この順序が保たれていることが重要（画像を直前の見出しのスライドに割り当てるため）。

このモジュールは Flask やプロジェクト保存の都合を知らない純粋な変換層。
保存やスライド化は slides.py 側が担当する。
"""
from __future__ import annotations

import io
import re
import urllib.error
import urllib.request
from dataclasses import dataclass, field

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from PIL import Image

# ---------------------------------------------------------------------------
# 取得（Google ドキュメント URL → docx バイト列）
# ---------------------------------------------------------------------------

# https://docs.google.com/document/d/<ID>/edit などから ID を取り出す
_DOC_URL_RE = re.compile(r"/document/(?:u/\d+/)?d/([A-Za-z0-9_-]{15,})")
# ?id=<ID> 形式（旧URL）
_DOC_ID_PARAM_RE = re.compile(r"[?&]id=([A-Za-z0-9_-]{15,})")
# ID だけを貼られた場合
_BARE_ID_RE = re.compile(r"^[A-Za-z0-9_-]{25,}$")

# zip（=docx）のマジックナンバー
_ZIP_MAGIC = b"PK\x03\x04"

FETCH_TIMEOUT = 60


class DocumentImportError(Exception):
    """取り込みに失敗したときの、ユーザーにそのまま見せられるエラー。"""


def extract_doc_id(url_or_id: str) -> str:
    """Google ドキュメントの URL（または ID）からドキュメント ID を取り出す。"""
    s = (url_or_id or "").strip()
    if not s:
        raise DocumentImportError("ドキュメントの URL を入力してください。")

    for pattern in (_DOC_URL_RE, _DOC_ID_PARAM_RE):
        m = pattern.search(s)
        if m:
            return m.group(1)

    if _BARE_ID_RE.match(s):
        return s

    if "docs.google.com" in s and "/document/" not in s:
        raise DocumentImportError(
            "Google ドキュメント以外の URL のようです。"
            "スプレッドシートやスライドではなく、ドキュメントの URL を貼ってください。"
        )
    raise DocumentImportError(
        "Google ドキュメントの URL を認識できませんでした。\n"
        "例: https://docs.google.com/document/d/xxxxxxxx/edit"
    )


def fetch_docx_from_google(url_or_id: str) -> bytes:
    """Google ドキュメントを docx としてダウンロードする。

    共有設定が「制限付き（自分だけ）」だと Google はログイン用の HTML を返す。
    その場合は docx として読めないので、対処法を添えて例外にする。
    """
    doc_id = extract_doc_id(url_or_id)
    export_url = f"https://docs.google.com/document/d/{doc_id}/export?format=docx"

    req = urllib.request.Request(
        export_url,
        headers={"User-Agent": "Mozilla/5.0 (Macintosh) AI-presentation-tool"},
    )
    try:
        with urllib.request.urlopen(req, timeout=FETCH_TIMEOUT) as res:
            content = res.read()
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            raise DocumentImportError(_SHARING_HELP)
        if e.code == 404:
            raise DocumentImportError(
                "ドキュメントが見つかりませんでした（404）。URL が正しいか確認してください。"
            )
        raise DocumentImportError(f"ドキュメントを取得できませんでした（HTTP {e.code}）。")
    except urllib.error.URLError as e:
        raise DocumentImportError(f"ドキュメントを取得できませんでした（通信エラー）: {e.reason}")

    if not (content or b"").startswith(_ZIP_MAGIC):
        # 非公開ドキュメントではログイン画面の HTML が返ってくる
        raise DocumentImportError(_SHARING_HELP)

    return content


_SHARING_HELP = (
    "このドキュメントは非公開のため読み取れませんでした。\n"
    "次のどちらかで取り込めます:\n"
    "① Google ドキュメントで「共有」→「リンクを知っている全員」を"
    "「閲覧者」に設定してから、もう一度取り込む\n"
    "② Google ドキュメントで「ファイル」→「ダウンロード」→"
    "「Microsoft Word (.docx)」で保存し、その .docx ファイルを直接アップロードする"
)


# ---------------------------------------------------------------------------
# ブロック列（docx を「文書順の意味のかたまり」に分解した中間表現）
# ---------------------------------------------------------------------------

@dataclass
class DocBlock:
    """ドキュメント内の1要素。

    kind:
      "heading" : 見出し（level 1〜3）
      "para"    : 本文段落
      "bullet"  : 箇条書き・番号付きリストの1項目
      "image"   : 画像（image_bytes に PNG/JPEG などの生バイト）
    """
    kind: str
    text: str = ""
    level: int = 0
    image_bytes: bytes = b""
    image_ext: str = "png"


@dataclass
class ParsedDocument:
    title: str = ""
    blocks: list = field(default_factory=list)

    @property
    def image_count(self) -> int:
        return sum(1 for b in self.blocks if b.kind == "image")

    @property
    def heading_count(self) -> int:
        return sum(1 for b in self.blocks if b.kind == "heading")


# 見出しスタイル名の判定（英語UI/日本語UI/スタイルID の揺れを吸収する）
_HEADING_STYLE_RE = re.compile(r"^(?:heading|見出し)\s*([1-9])$", re.IGNORECASE)

# 小さすぎる画像は装飾・区切り線・箇条書きアイコンなどのノイズとみなして捨てる
MIN_IMAGE_PIXELS = 64

# 画像参照のタグ名。VML（v:）は python-docx の名前空間表に無いため直接組み立てる。
_TAG_BLIP = qn("a:blip")
_ATTR_EMBED = qn("r:embed")
_ATTR_LINK = qn("r:link")
_NS_VML = "urn:schemas-microsoft-com:vml"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_TAG_IMAGEDATA = f"{{{_NS_VML}}}imagedata"
_ATTR_REL_ID = f"{{{_NS_REL}}}id"


def _style_names(paragraph: Paragraph) -> list:
    """段落に付いているスタイル名の候補を集める。

    styles.xml に定義が無いスタイルを参照している docx もあるため、
    まず XML の pStyle 値を直接読む（例外が起きない）。そのうえで
    python-docx 経由の表示名も、取得できれば候補に加える。
    """
    names = []

    p_pr = paragraph._p.pPr
    if p_pr is not None and p_pr.pStyle is not None:
        val = p_pr.pStyle.val
        if val:
            names.append(val)

    try:
        style = paragraph.style
    except KeyError:
        style = None
    if style is not None:
        if style.name:
            names.append(style.name)
        style_id = getattr(style, "style_id", None)
        if style_id:
            names.append(style_id)

    # "Heading1" / "見出し1" のように数字が直に続く形式も拾えるよう空白を挿す
    return names + [re.sub(r"(\D)(\d)", r"\1 \2", n) for n in names]


def _heading_level(paragraph: Paragraph) -> int:
    """段落の見出しレベル（1〜3）を返す。見出しでなければ 0。

    見出し4以降は 4 以上を返し、呼び出し側で本文に寄せる。
    """
    for name in _style_names(paragraph):
        stripped = name.strip()
        m = _HEADING_STYLE_RE.match(stripped)
        if m:
            return int(m.group(1))
        if stripped.lower() == "title" or stripped in ("タイトル", "表題"):
            return 1
        if stripped.lower() == "subtitle" or stripped in ("副題", "サブタイトル"):
            return 2
    return 0


def _is_list_item(paragraph: Paragraph) -> bool:
    """箇条書き／番号付きリストの項目かどうか。"""
    p_pr = paragraph._p.pPr
    if p_pr is not None and p_pr.numPr is not None:
        return True
    return any(
        name.strip().lower().startswith("list") or name.startswith("リスト")
        for name in _style_names(paragraph)
    )


def _image_ext(content_type: str, partname: str) -> str:
    ct = (content_type or "").lower()
    for key, ext in (("png", "png"), ("jpeg", "jpg"), ("jpg", "jpg"),
                     ("gif", "gif"), ("webp", "webp"), ("bmp", "bmp")):
        if key in ct:
            return ext
    suffix = partname.rsplit(".", 1)[-1].lower() if "." in partname else ""
    return suffix or "png"


def _extract_images(element, doc) -> list:
    """段落・表セル要素に含まれる画像を、出現順に取り出す。

    通常の埋め込み画像（DrawingML の a:blip）と、古い形式の
    VML 画像（v:imagedata）の両方を拾う。インライン配置・回り込み配置の
    どちらでも a:blip は同じように現れるため区別しない。
    """
    blocks = []
    related = doc.part.related_parts
    seen_rids = set()

    for node in element.iter():
        rid = None
        if node.tag == _TAG_BLIP:
            rid = node.get(_ATTR_EMBED) or node.get(_ATTR_LINK)
        elif node.tag == _TAG_IMAGEDATA:
            rid = node.get(_ATTR_REL_ID)
        if not rid or rid in seen_rids:
            continue
        seen_rids.add(rid)

        part = related.get(rid)
        if part is None:
            continue
        blob = getattr(part, "blob", b"")
        if not blob:
            continue

        # 画像として開けるか＆十分な大きさかを確認する
        try:
            with Image.open(io.BytesIO(blob)) as im:
                w, h = im.size
        except Exception:
            continue
        if w < MIN_IMAGE_PIXELS or h < MIN_IMAGE_PIXELS:
            continue

        blocks.append(DocBlock(
            kind="image",
            image_bytes=blob,
            image_ext=_image_ext(getattr(part, "content_type", ""), str(part.partname)),
        ))
    return blocks


def _table_lines(table_element, doc) -> list:
    """表を「1行=1テキスト」の箇条書きに落とす（セルは / 区切り）。"""
    from docx.table import Table

    table = Table(table_element, doc)
    lines = []
    for row in table.rows:
        cells = [c.text.strip().replace("\n", " ") for c in row.cells]
        # 横方向に結合されたセルは同じテキストが重複するので畳む
        deduped = []
        for c in cells:
            if not deduped or deduped[-1] != c:
                deduped.append(c)
        text = " / ".join(c for c in deduped if c)
        if text:
            lines.append(text)
    return lines


def parse_docx(data: bytes) -> ParsedDocument:
    """docx バイト列を、文書順のブロック列に分解する。"""
    try:
        doc = Document(io.BytesIO(data))
    except Exception as e:
        raise DocumentImportError(
            f"ドキュメント（.docx）として読み込めませんでした: {e}"
        )

    parsed = ParsedDocument()
    body = doc.element.body

    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            paragraph = Paragraph(child, doc)
            text = (paragraph.text or "").strip()
            level = _heading_level(paragraph)

            if text:
                if 1 <= level <= 3:
                    parsed.blocks.append(DocBlock(kind="heading", text=text, level=level))
                    if not parsed.title:
                        parsed.title = text
                elif _is_list_item(paragraph):
                    parsed.blocks.append(DocBlock(kind="bullet", text=text))
                else:
                    # 見出し4以降も本文として束ねる（スライドを増やしすぎない）
                    parsed.blocks.append(DocBlock(kind="para", text=text))

            # 同じ段落に画像がある場合はテキストの直後に置く
            parsed.blocks.extend(_extract_images(child, doc))

        elif child.tag == qn("w:tbl"):
            for line in _table_lines(child, doc):
                parsed.blocks.append(DocBlock(kind="bullet", text=line))
            parsed.blocks.extend(_extract_images(child, doc))

    if not parsed.blocks:
        raise DocumentImportError(
            "ドキュメントから内容を読み取れませんでした（中身が空のようです）。"
        )
    return parsed


def parse_google_doc(url_or_id: str) -> ParsedDocument:
    """Google ドキュメントの URL から取得してブロック列に分解する。"""
    return parse_docx(fetch_docx_from_google(url_or_id))
