# ローカルデータ（Git 管理外）

このディレクトリに **自分の Excel** を置き、インデックスを作ります。
会社の xlsx や個人情報を含むファイルは **コミットしないでください**。

## 配置するファイル（相対パス）

リポジトリ直下からの相対パス:

| 相対パス | 内容 |
|----------|------|
| `data/internal.xlsx` | 社内ナレッジブック（Problem & Improvement / 工具 / TS など 53 シート想定） |
| `data/qa.xlsx` | 外部 Q&A（シート名 `Q&Aデータ`） |
| `data/knowledge.db` | 取り込み後に生成される SQLite（自動作成） |

```bash
# リポジトリ直下で
cp /path/to/社内ブック.xlsx data/internal.xlsx
cp /path/to/機械加工Q&A.xlsx data/qa.xlsx
python ingest.py
```

デモ用の小さな合成ブックは `sample_data/` にあります（実データではありません）。
