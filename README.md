# インバウンドマップ

国籍別の延べ宿泊者数・国際線・客室稼働率を見る地図サイトです。GitHub Pages では `index.html` を配信しています。

## ニュースまとめの更新

「ニュースまとめ」の本文は `news.json` です。地図本体の `index.html`（約 3MB）は日次更新では触りません。マスターの Excel から JSON だけ作り直し、そのファイルをコミットしてください。

Excel（`旅行業界ニュース` と `海外OTAニュース` の2シート）が、載せる項目と文言の正です。Google スプレッドシートの CSV は、出典・ソース列の URL を全部取り込むために使います（Excel のセルにはハイパーリンクが1つしか入りません）。日付と見出しが一致した行だけ CSV の URL を使い、一致しない場合は Excel のハイパーリンクに戻します。

各項目には元のカテゴリに加えて `theme` が入ります。画面の絞り込みチップは「すべて」と、海外OTA（Expedia・Hotelbeds・サイトコントローラーなどの卸・配信も含む） / 国内OTA・旅行会社 / 新規開業ホテル / 国際線フライト / 行政・インバウンド施策 / その他（該当があるときだけ）です。振り分け規則は `scripts/build_news.py` の `assign_theme` にあり、見出しの先頭の話題を優先します。📌要点はチップにせず、その日の要約として日付の先頭に出ます。

必要なもの: Python 3 と [openpyxl](https://openpyxl.readthedocs.io/)（`pip install -r scripts/requirements.txt`）。

```bash
python3 scripts/build_news.py /path/to/旅行業界・海外OTAニュース.xlsx \
  --industry-csv /path/to/gsheet_industry.csv \
  --ota-csv /path/to/gsheet_ota.csv \
  -o news.json
```

CSV が無い日は `--industry-csv` と `--ota-csv` を省けます。日付に年が無い行は `--year`（既定 2026）を使い、新しい順に読んで月が戻ったら前年とみなします。

```bash
python3 scripts/build_news.py /path/to/旅行業界・海外OTAニュース.xlsx -o news.json
```

生成した `news.json` をリポジトリ直下に置いて push すると、サイトの「ニュースまとめ」ボタンが次回表示時にその内容を読み込みます。
