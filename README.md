# AI presentation tool

Markdown を入力すると、見出しごとに1枚のスライド画像を **GPT Image 2.0** で生成し、
PPTX として書き出す3ペインのデスクトップツール（Flask）。

`premovie-4pines` のスライド生成エンジンを汎用化して構築。

## 3ペイン構成

| ペイン | 役割 |
|--------|------|
| 1 プロジェクト管理 | 一覧／`＋`新規作成（Markdown・デザインプロンプト・見本画像） |
| 2 台本編集 | Markdown を編集 →「スライド生成」 |
| 3 プレビュー | 各スライド表示／修正入力＋個別再生成／一括再生成／PPTX書き出し |

## スライド分割ルール

- Markdown の **見出し1/2/3（`#` `##` `###`）ごとに1スライド**
- 見出し4以降・箇条書き・段落は直近スライドの本文にまとまる

## セットアップ

```bash
cd ~/src/AI-presentation-tool
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp .env.example .env   # OPENAI_API_KEY などを設定
```

`.env` の主な項目:

| 変数 | 説明 |
|------|------|
| `OPENAI_API_KEY` | OpenAI APIキー |
| `SLIDE_IMAGE_MODEL` | 画像生成モデル（既定 `gpt-image-2`） |
| `OUTPUT_ROOT` | 出力先ルート（既定 iCloud Drive の `Downloads`） |

## 起動

```bash
./start.sh
# または
.venv/bin/python run.py
```

→ http://localhost:5070

出力は `OUTPUT_ROOT/<プロジェクト名>/` に `slide_NN.png` と `<プロジェクト名>.pptx` が保存される。

## 重要な運用上の注意（macOS）

出力先が iCloud Drive のため、**サーバーは単一プロセスで起動する必要がある**
（`run.py` は `debug=False, use_reloader=False`）。
LaunchAgent/launchd 経由で起動すると macOS の TCC 権限継承が切れて
iCloud にアクセスできなくなる。Automator アプリやターミナル（フルディスクアクセス付与済み）
の子プロセスとして起動すること。詳細は `docs/MIGRATION.md` 参照。
