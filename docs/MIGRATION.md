# AI-presentation-tool 移植指示書（チェックリスト）

`premovie-4pines` の資産を新リポジトリ **`AI-presentation-tool`**（3ペインのスライド自動生成ツール）へ
移植するための作業指示書。Agent に渡してそのまま実装できる粒度で記述する。

> このファイルは `premovie-4pines/docs/` で作成。新リポジトリ作成後はそちらの `docs/` へ移動してよい。

---

## 0. 完成イメージ（要件）

3ペイン構成:

| ペイン | 役割 |
|--------|------|
| 1 プロジェクト管理 | プロジェクト一覧／`＋`新規作成（Markdown入力・見本画像添付・プロンプト選択） |
| 2 台本編集 | Markdown を表示・直接編集／上部に「スライド生成」ボタン |
| 3 プレビュー | スライドを1枚ずつ表示／各下に修正入力＋「再生成」／「一括再生成」／「PPTX書き出し」 |

中核ルール:
- **Markdown の見出し1/2/3ごとに1スライド**（h4以降・箇条書きは分割しない）
- 画像生成は **GPT Image 2.0 API**（`images.edit` に見本画像を渡す）
- 出力先は **`~/Library/Mobile Documents/com~apple~CloudDocs/Downloads/<プロジェクト名>/`**
  - そこに `slide_NN.png` 群と `<プロジェクト名>.pptx` を保存

---

## 1. 移植するファイル一覧

### そのままコピー（変更ほぼ不要）

- [ ] `app/__init__.py` … Flask app factory。**DB 初期化呼び出し（`db_svc.init_db()`）は削除**
- [ ] `run.py` … `debug=False, use_reloader=False` の単一プロセス起動を**必ず踏襲**（iCloud/TCC対策）
- [ ] `start.sh` … Dock/Automator 起動用。二重起動防止＋ブラウザ起動
- [ ] `.gitignore`

### コピー後に改修

- [ ] `app/slides.py` → スライドエンジン（最重要・後述の改修）
- [ ] `app/routes.py` → スライド系APIのみ抽出して改修
- [ ] `app/templates/index.html` → 4ペイン→3ペインに削減
- [ ] `requirements.txt` → 下記だけに絞る

```
flask>=3.0.0
openai>=1.30.0
python-dotenv>=1.0.0
python-pptx>=0.6.23
pillow>=10.0.0
```

### 移植しない（不要）

- [ ] `app/citations.py`（引用論文）
- [ ] `app/db.py`（Neon/Postgres）
- [ ] `app/google_docs.py`（Google Docs）
- [ ] `app/thumbnails.py`（サムネ・冒頭画像）
- [ ] `app/archive.py`（アーカイブ）
- [ ] `data/style-guide.md`、`scripts/verify_*.py`

---

## 2. `app/slides.py` の改修指示

### 2-1. パス設定の差し替え
- [ ] `PROJECTS_DIR` → 廃止。代わりに出力ルートを
      `OUTPUT_ROOT = ~/Library/Mobile Documents/com~apple~CloudDocs/Downloads`
- [ ] `slides_dir(project)` → `OUTPUT_ROOT / project`（pptx・png をここに保存）
- [ ] `AGENT_ROOT / PROMPTS_DIR / COMMON_DIR / CHANNEL_SETTINGS_PATH` → 削除

### 2-2. チャンネル概念をプロジェクト概念へ置換
- [ ] `_parse_channel_settings` / `get_channel_config` / `_read_channel_prompt` → **削除**
- [ ] 代わりに「プロジェクトごとの見本画像＋生成プロンプト」を扱う:
  - 見本画像: `OUTPUT_ROOT/<project>/_reference.png`（新規作成時にアップロード。無ければデフォルト画像）
  - 生成プロンプト: プロジェクト作成時に選んだテンプレ文字列を `OUTPUT_ROOT/<project>/_prompt.txt` に保存

### 2-3. `generate_slide_plan` を「Markdown見出しパーサ」に置換（LLM不要）
- [ ] 既存の LLM 呼び出し版は削除し、決定論パーサに:
  - 入力: Markdown 全文
  - `#` `##` `###`（h1/h2/h3）が出るたびに新スライド境界
  - 各スライド = `{title: 見出しテキスト, lines: [配下のh4/箇条書き/段落を整形], illustration: ""}`
  - h4以降・箇条書き・段落は**直近スライドの本文**に束ねる（スライドを分けない）
- [ ] （任意・後段オプション）各スライドの `illustration` を AI に肉付けさせるモードを別関数で用意

### 2-4. 画像生成まわり（流用、軽微改修）
- [ ] `_build_image_prompt` … 医療/右1/3固定/出典の文言を**プレゼン汎用**に書き換え。
      ベース引数 `channel_prompt` → `project_prompt` にリネーム
- [ ] `generate_slide_image` … 引数 `base_image_path` → プロジェクトの `_reference.png`。
      **修正テキスト(fix)** を受け取り、プロンプト末尾に「【修正指示】…」を足せるよう拡張
- [ ] `_fit_to_16x9` … **そのまま流用**（縦圧縮フィット）。citation 分岐は削除可
- [ ] `_draw_citation_footer` / `_load_citation_font` / `CITATION_*` … 出典不要なら削除
- [ ] `_parse_json_array` … 任意機能で使うなら残す
- [ ] `build_pptx` … **そのまま流用**（1枚画像=1スライド・16:9）。保存先を `OUTPUT_ROOT/<project>` に

---

## 3. `app/routes.py` の改修指示

### 流用するエンドポイント（スライド系）
- [ ] `GET /api/slides_status` … チャンネル依存を外し、プロジェクトの plan/生成済みPNG/pptx有無を返す
- [ ] `POST /api/slide_plan` … 台本Markdown→見出しパーサで plan.json 生成
- [ ] `POST /api/generate_slide` … index指定で1枚生成（**fixテキスト引数を追加**）
- [ ] `GET /api/slide_image` … PNG配信（そのまま）
- [ ] `POST /api/build_pptx` … PPTX組立（そのまま）
- [ ] `GET /api/download_pptx` … ダウンロード（そのまま）

### 新規・改修が必要なエンドポイント
- [ ] `GET /api/projects` … `OUTPUT_ROOT` 配下のプロジェクトフォルダ一覧
- [ ] `POST /api/create_project` … 名前未指定なら「プロジェクトN」自動採番。
      Markdown本文・選択プロンプト・（任意）見本画像を保存
- [ ] `GET /api/project/<name>` … Markdown本文・plan・生成状況を返す
- [ ] `POST /api/save_markdown` … ペイン2の編集内容を保存
- [ ] `POST /api/upload_reference` … 見本画像を `_reference.png` で保存（既存 `upload_model_thumbnail` が雛形）
- [ ] `POST /api/regenerate_all` … 各スライドの修正テキストをまとめて受け、全再生成
- [ ] プロジェクト名のサニタイズ（`_valid_folder_name` を流用）

### 削除するエンドポイント
- [ ] 台本生成(E)/引用/Google認証/Docs/サムネ/冒頭画像/スタイル学習/アーカイブ 系すべて

---

## 4. `index.html` の改修指示
- [ ] 4ペイン→**3ペイン**レイアウトに削減
- [ ] ペイン1: プロジェクト一覧＋`＋`（新規作成モーダル: Markdown textarea・プロンプト選択・画像添付）
- [ ] ペイン2: Markdown textarea（編集可）＋「スライド生成」ボタン
- [ ] ペイン3: スライドカードを縦並び。各カードに `<img>`＋修正テキスト入力＋「再生成」。
      上部に「一括再生成」「PPTX書き出し」
- [ ] fetch ヘルパ・モーダル・エスケープ処理（`escHtml`）は既存から流用

---

## 5. 必ず引き継ぐ「知見」（ハマりどころ）

- [ ] **iCloud × macOS TCC**: Flask は単一プロセス（`debug=False, use_reloader=False`）。
      LaunchAgent/launchd 経由起動は iCloud アクセスが切れるので避け、
      FDA を持つ親（Automator/ターミナル）の子として起動する。**出力が iCloud Downloads なので直撃する**
- [ ] **gpt-image-2 の癖**: 小さな英数字は化けやすい→重要英数字は PIL で焼き込む（必要時）。
      `images.edit` に見本画像を渡して構図を固定
- [ ] **16:9 整形**: クロップは端が切れる→縦圧縮フィット（`_fit_to_16x9`）
- [ ] **JSON頑健パース**: `_parse_json_array`（```フェンス除去）
- [ ] **.env**: `OPENAI_API_KEY` / `SLIDE_IMAGE_MODEL=gpt-image-2` /（任意）`OUTPUT_ROOT`

---

## 6. 新リポジトリ構成（完成形の目安）

```
AI-presentation-tool/
├── app/
│   ├── __init__.py        # app factory（DB初期化なし）
│   ├── routes.py          # スライド系API＋プロジェクトCRUD
│   ├── slides.py          # 見出しパーサ＋画像生成＋PPTX
│   └── templates/
│       └── index.html     # 3ペインUI
├── run.py                 # 単一プロセス起動
├── start.sh
├── requirements.txt
├── .env.example
├── .gitignore
└── docs/
    └── AI-presentation-tool-migration.md  # 本書
```

---

## 7. 実装順（推奨）
1. スキャフォールド（`__init__.py`/`run.py`/`requirements.txt`/`.env.example`/`start.sh`）
2. `slides.py`: 見出しパーサ → `build_pptx` → `generate_slide_image` の順で単体動作
3. `routes.py`: projects CRUD → slide_plan → generate_slide → build_pptx
4. `index.html`: 3ペインを配線
5. iCloud Downloads 保存＆単一プロセス起動を実機確認
6. 修正テキスト→個別/一括再生成を仕上げ
