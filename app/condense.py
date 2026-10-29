"""
本文要約層（長い本文 → スライドに載せるキーワード＋図解指示）

ドキュメントの本文をそのままスライドに流し込むと文字量が多すぎて読めないため、
テキストLLMで「スライドに載せる分量」まで絞り込む。

各スライドについて以下を作る:
  - title        : 見出し（長すぎる場合だけ短く言い換える）
  - keywords     : スライドに大きく描く短いキーワード・数字（3〜6個）
  - illustration : 内容を直感的に伝える図解・イラストの指示
  - image_notes  : スライド内の画像1枚ごとに「何を示している画像か」のメモ

ドキュメント上で画像が挟まっている位置は「その画像がどの話題のものか」を
表す情報なので、本文と画像の並び順をそのまま渡して文脈ごと読ませる。
image_notes は画像生成時にそのまま「画像1はこういう画像」という説明として使う。

ドキュメント全体を1回のリクエストで渡すことで、スライド間でキーワードが
重複したり話の流れが分断されるのを防ぐ。

引用文献スライド（slides.py が references を持たせたもの）はここを素通りする。
文献は番号・著者・雑誌名・巻・号・ページ・年がそろって初めて引用として成立するため、
キーワードにまとめてはいけない。
"""
from __future__ import annotations

import json
import os

# 1スライドあたり・全体でLLMに渡す本文の上限（長大なドキュメントの保険）
MAX_CHARS_PER_SLIDE = 3000
MAX_TOTAL_CHARS = 120000

# 生成するキーワードの目安
MIN_KEYWORDS = 3
MAX_KEYWORDS = 6


def text_model() -> str:
    return os.environ.get("SLIDE_TEXT_MODEL", "gpt-6-astra")


def reasoning_effort() -> str:
    # gpt-6-astra は none / minimal を受け付けない。旧モデル向けの .env が
    # 残っていても落ちないよう、公式の移行手順どおり low に寄せる。
    effort = os.environ.get("SLIDE_TEXT_EFFORT", "low").strip().lower()
    return "low" if effort in ("", "none", "minimal") else effort


_SYSTEM_PROMPT = f"""あなたはプレゼンテーション構成の専門家です。
ドキュメントの各セクションを「1枚のスライドに載せる分量」まで絞り込みます。

出力ルール:
1. keywords は各スライド {MIN_KEYWORDS}〜{MAX_KEYWORDS} 個。
   ・1個は原則20文字以内の単語・短いフレーズ・数字にする。文章にしない。
   ・「〜です」「〜ます」などの述語で終わる文にしない（体言止め・名詞句にする）。
   ・本文の要点を writer の言葉で言い換えてよいが、事実・数値・固有名詞は変えない。
   ・数字・割合・年数・期間などの具体的な数値は、含まれていれば優先的に残す
     （数字は聞き手の記憶に残りやすいため）。
   ・箇条書きが元々ある場合は、それを短くしたものを優先して使う。
   ・スライドをまたいで同じキーワードを繰り返さない。
2. title は元の見出しをそのまま使う。ただし25文字を超える場合のみ、
   意味を変えずに短く言い換える。
3. illustration は、そのスライドの内容を絵で直感的に理解させるための指示。
   ・何をどう配置するか具体的に書く（例:「左に正常な腱板、右に炎症で赤くなった
     腱板を並べた対比イラスト」「3段の階段で改善ステップを示す図」）。
   ・図解の中に長い説明文を入れる指示はしない。
   ・そのスライドに「既存の画像」がある場合は、その画像が主役なので、
     illustration はそれを邪魔しない控えめな装飾・補助的な図にとどめる。
     何も足さないほうが良ければ空文字列にする。
4. image_notes は、入力に［画像1］［画像2］として現れた画像の枚数ぶん、
   同じ順番で作る（画像が無いスライドは空配列）。
   ・画像はドキュメントの「その位置」に貼られているので、直前後の本文から
     その画像が何を示しているかを読み取り、1行で書く
     （例:「腱板の炎症部位を示す解剖図。『腱板炎』の説明に対応」）。
   ・あわせて、どのキーワードの近くに置くと文脈が通るかを一言添える。
   ・画像の中身を勝手に決めつけず、本文から言えることだけを書く。
   ・画像を切り取る・一部を強調するといった指示は書かない。
5. 医療・健康の内容では、断定的な効果保証や誇張表現を作らない。
   元の本文に書かれていないことを足さない。
6. すべて日本語で出力する。
"""

_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["slides"],
    "properties": {
        "slides": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["index", "title", "keywords", "illustration", "image_notes"],
                "properties": {
                    "index": {"type": "integer", "description": "対象スライドの番号"},
                    "title": {"type": "string"},
                    "keywords": {"type": "array", "items": {"type": "string"}},
                    "illustration": {"type": "string"},
                    "image_notes": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "画像1枚ごとの説明。画像の出現順に並べる",
                    },
                },
            },
        }
    },
}


def _slide_body(slide: dict) -> str:
    """本文と画像を、ドキュメントに出てきた順に並べた本文ブロックを作る。

    画像は［画像N］という行として本文の間に挟む。この前後関係が、
    その画像がどの話題のものかを示す手がかりになる。
    """
    content = slide.get("content")
    if not content:
        # content を持たない古い plan.json との互換
        return "\n".join(slide.get("lines") or [])

    lines = []
    image_no = 0
    for item in content:
        if item.get("kind") == "image":
            image_no += 1
            lines.append(f"［画像{image_no}］")
        else:
            text = (item.get("text") or "").strip()
            if text:
                lines.append(text)
    return "\n".join(lines)


def is_verbatim(slide: dict) -> bool:
    """要約してはいけないスライドか（引用文献リストなど）。

    文献は番号・著者・雑誌名・巻・号・ページ・年がそろって初めて引用として成立する。
    キーワードにまとめた時点で引用ではなくなるので、要約層を素通りさせる。
    """
    return bool(slide.get("references"))


def _build_input(slides: list) -> str:
    """スライド構成案をLLMへの入力テキストに整形する。"""
    parts = []
    for i, slide in enumerate(slides):
        if is_verbatim(slide):
            continue
        body = _slide_body(slide)[:MAX_CHARS_PER_SLIDE]
        image_count = len(slide.get("images") or [])
        block = [f"=== スライド {i} ==="]
        block.append(f"見出し: {slide.get('title') or '(無題)'}")
        block.append(
            f"このスライドの画像: {image_count}枚"
            f"（本文中の［画像1］〜［画像{image_count}］の位置に貼られている）"
            if image_count else "このスライドの画像: なし"
        )
        block.append("本文（ドキュメントに出てくる順）:")
        block.append(body if body.strip() else "(本文なし。見出しだけのスライド)")
        parts.append("\n".join(block))

    text = "\n\n".join(parts)
    return text[:MAX_TOTAL_CHARS]


def condense_slides(slides: list, client) -> list:
    """各スライドに keywords と illustration を付与した新しいリストを返す。

    LLM が一部のスライドを返さなかった場合、そのスライドは元のまま
    （keywords 無し）になる。呼び出し側は keywords が無ければ本文をそのまま
    素材として使うため、部分的な失敗でも生成は続行できる。
    """
    if not slides or all(is_verbatim(s) for s in slides):
        return list(slides)

    user_input = (
        "以下はドキュメントを見出しごとにスライドへ分解したものです。\n"
        "すべてのスライドについて、index を対応させたまま出力してください。\n\n"
        + _build_input(slides)
    )

    response = client.responses.create(
        model=text_model(),
        input=[
            {"role": "developer", "content": _SYSTEM_PROMPT},
            {"role": "user", "content": user_input},
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "slide_keywords",
                "strict": True,
                "schema": _SCHEMA,
            }
        },
        reasoning={"effort": reasoning_effort()},
    )

    data = json.loads(response.output_text or "{}")

    by_index = {}
    for item in data.get("slides") or []:
        try:
            by_index[int(item.get("index"))] = item
        except (TypeError, ValueError):
            continue

    result = []
    for i, slide in enumerate(slides):
        merged = dict(slide)
        item = None if is_verbatim(slide) else by_index.get(i)
        if item:
            keywords = [
                k.strip() for k in (item.get("keywords") or [])
                if isinstance(k, str) and k.strip()
            ]
            if keywords:
                merged["keywords"] = keywords[:MAX_KEYWORDS]
            title = (item.get("title") or "").strip()
            if title:
                merged["title"] = title
            illustration = (item.get("illustration") or "").strip()
            if illustration:
                merged["illustration"] = illustration
            # 画像メモは画像と1対1で対応させる。多い分は捨て、足りない分は空にして
            # 「画像N のメモ」と実際の N 番目の画像がずれないようにする。
            image_count = len(slide.get("images") or [])
            if image_count:
                notes = [
                    str(n).strip() for n in (item.get("image_notes") or [])
                    if isinstance(n, str)
                ][:image_count]
                notes += [""] * (image_count - len(notes))
                merged["image_notes"] = notes
        result.append(merged)
    return result
