# STEAM教育・子ども向けプログラミング教室 競合比較

社内の戦略検討用に、ステモンと競合13社（参考としてナリカ）を比較した資料です。調査日は 2026-09-28 です。

| ファイル | 内容 |
|---|---|
| `STEAM教育_競合比較_サマリー.pptx` | 要点をまとめたスライド17枚（結論、料金、レッスン内容、先生の質、教材、口コミ、課題、ポジショニング、SWOT、提言） |
| `STEAM教育_競合比較_元データ.xlsx` | 全項目の元データ10シート（料金の総額は数式で自動計算、要確認リスト、出典URL） |
| `build/` | 資料を生成するスクリプト（`data.py` がデータの正本） |

## 注意
- 数値はWeb検索結果の要約から集めたものです（公式ページの本文は直接見ていません）。
- 社外で使う前や最終判断の前に、Excelの「要確認リスト」シートの項目を確認してください。

## 更新方法
- **Excelだけ直す場合**：「料金比較」シートの青字セルを書き換えると、総額・グラフ・サマリーが自動で再計算されます。
- **両方を作り直す場合**：`build/data.py` を編集してから、次のコマンドを実行します。

```bash
cd build
pip install openpyxl && npm install
python3 build_xlsx.py
python3 -c "import json;from data import *;[c.update(y1=c['entry']+(c['m1']+c['misc'])*12+c['mat1'],y3=c['entry']+(c['m1']+c['m2']+c['m3']+3*c['misc'])*12+c['mat1']+c['mat2']+c['mat3']) for c in COMPANIES];json.dump(dict(companies=COMPANIES,narika=NARIKA,others=OTHERS,swot=SWOT,actions=ACTIONS,date=SURVEY_DATE),open('data.json','w'),ensure_ascii=False)"
node build_pptx.js
```
