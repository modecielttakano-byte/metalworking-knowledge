# 加工方法ナレッジ検索

社内の **問題・対策ログ / 工具・トラブルシュート** と、外部の **機械加工 Q&A** を、同じ検索箱から横断します。  
日本語キーワード向けのローカル Web アプリです（認証なし）。

ソースは信頼度を分けて表示します。

| バッジ | 意味 |
|--------|------|
| **社内ログ**（社内実績） | 工程シートの現象・実施対策・結果 |
| **社内TS・工具**（社内実績） | ドリル/リーマ TS、エンドミル選定、加工条件メモ |
| **外部Q&A**（掲示板意見） | mori.nc-net / OKWave 系の質問と回答。未検証 |

この GitHub リポジトリは **公開** です。会社の Excel や氏名は **絶対にコミットしないでください**。

---

## 必要なもの

- Python 3.10 以降

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

---

## Excel の配置（相対パス）

リポジトリ直下に `data/` を作り、次の **固定ファイル名** で置きます。

| 相対パス | 中身 |
|----------|------|
| `data/internal.xlsx` | 社内ナレッジブック（53 シート想定。`SWG` / `ARM` / `mil` / `ドリルトラブルシューティング` など） |
| `data/qa.xlsx` | 外部 Q&A。シート名 **`Q&Aデータ`**（列: URL, カテゴリ, 質問タイトル, 質問本文, 回答, ベストアンサー） |
| `data/knowledge.db` | 取り込み後に自動生成（Git 管理外） |

```bash
mkdir -p data
cp /path/to/社内ブック.xlsx data/internal.xlsx
cp /path/to/機械加工Q&A.xlsx data/qa.xlsx
```

`data/` 以下の xlsx と db は `.gitignore` 済みです。

取り込み時に **作成者・担当者セル** と `〇〇さん` / `By Name` は検索テキストから外します。工程本文まで完全に匿名化するものではありません。

---

## 取り込み

```bash
python ingest.py
```

オプション:

```bash
python ingest.py --internal data/internal.xlsx --qa data/qa.xlsx --db data/knowledge.db
```

デモ用の合成ブック（実データではない）だけを試す場合:

```bash
python make_sample.py          # 初回のみ。sample_data/*.xlsx を生成
python ingest.py --sample
```

標準出力に件数が出ます。社内ログは空行・継続行をまとめて 1 件にしています。

---

## 検索 UI の起動

```bash
python app.py
```

ブラウザで [http://127.0.0.1:8000](http://127.0.0.1:8000) を開きます。

- 検索例: `バリ` / `リーマ` / `面取り` / `送り`
- フィルタ: ソース種別、工程、社内ログの結果あり、Q&A の回答あり
- 並び: 社内の対策・結果があるものを優先。外部は回答ありを優先

---

## 取り込むシート（社内側）

完全な ETL ではなく、検索できる PoC です。主に次を索引します。

- 問題・対策ログ: `SWG`, `FPC`, `COIL`, `ARM`, `ARR`, `BH`, `PVT`, `OL1`, `OL2`, `SLIT`, `TAP`, `PH`, `SH`, `DT`, `OD`, `SWS`, `Drill`, `Chamfer`, `APP`, `JIG`, `Coolant`, `Loader`
- TS・工具: `ドリルトラブルシューティング`, `リーマトラブルシューティング`, `mil`, `リーマ加工条件計算`, `Root & Tip Step`, `穴径大`, `穴挽き目` など
- Picture / Pic は **キャプション文字列のみ**（画像ファイルは保存しません）

CMM の広域表や、ほぼ図だけのシートは対象外です。知識は Excel にある文言だけを使い、加工ノウハウを捏造しません。

---

## 開発メモ

- 索引: SQLite FTS5（日本語は文字 n-gram）
- API: `GET /api/search`, `GET /api/facets`, `GET /api/docs/{id}`, `GET /api/health`
- アプリ本体: `ingest.py` / `app.py` / `static/`
