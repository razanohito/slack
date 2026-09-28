# ビー玉TODO ⇄ 滞留ボード 連携手順

claude.ai の Artifact はページごとに別のデータベースを持つため、ページ同士は直接読み合えません。
Claude が `ArtifactData` ツールで両方を読み書きして橋渡しします。

| ページ | URL | 使うコレクション |
|---|---|---|
| ビー玉TODO | https://claude.ai/artifact/CU7Tub9jbdzEdk1s5S2pgb | `inbox`（所有者のみ読み書き可） |
| 滞留ボード | https://claude.ai/artifact/7gNqWhM7ggJAGprsXrfVbc | `decisions` |

## 1. 滞留ボード → ビー玉TODO（案件を届ける）

対象：滞留ボードの `ITEMS` のうち、`decisions/<id>` が未決・`do`・`defer` のもの（`drop` は届けない）。

`inbox/tb-<id>` に `set`（既にあれば `text` などは上書きせず、無ければ作成）：

```json
{
  "text": "一文で書いた、やること",
  "cat": "work", "pri": 3,
  "due": "YYYY-MM-DD（defer なら until、無ければ案件の期限）",
  "done": false, "doneAt": null, "archived": false, "hidden": false,
  "created": 1790584282960,
  "source": "滞留ボード", "srcId": "<滞留ボードの id>",
  "link": "元メール等のURL", "linkLabel": "Gmailで開く",
  "note": "判断に必要な補足（短く）"
}
```

## 2. ビー玉TODO → 滞留ボード（結果を戻す）

`inbox` を読み、`done: true` になった案件は滞留ボードの `decisions/<srcId>` に
`{"decision": "do", "until": null, "at": "<ISO時刻>"}` を書く。
`hidden: true`（一覧から外した）は何もしない（判断は滞留ボードで行う）。

## ページ側の挙動

- `inbox` の案件は自分のタスクと並んで表示され、黒い「滞留ボード」バッジと元リンクが付く
- 完了・編集・一覧から外す操作は `inbox` のドキュメントを `update` する
- 完了するとビー玉・XP に反映される
