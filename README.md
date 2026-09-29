# jp-market-dashboard

日本株市場の「セクターシフト」と「資金流出入（Cashout）」を監視する、単一 HTML ファイルのダッシュボード。

- 公開URL: https://tarzanjp.github.io/jp-market-dashboard/
- `index.html` 一つに完結（外部 CSS/JS なし）。ブラウザで直接開いても動作します。
- データの出所はブロックごとに異なる。**ページ本文の表記が正**（この README が古くなりやすいため）。

| ブロック | 出所 | 状態 |
|---|---|---|
| 01 Cashout 判定インジケーター | Yahoo!ファイナンス 売買代金ランキング（`data/jp-market.json`, `quality.primeTurnover: live`） | **実データ** |
| 02 東証33業種 資金流転マトリクス | — | **プリセット／シミュレーション値** |
| 03 投資部門別データ（週次） | JPX 投資部門別売買状況 `.xls` を標準ライブラリで解析（`data/jp-investor.json`, `quality.investorType: live`） | **実データ** |
| 04 セクター主導銘柄の監視 | — | **プリセット／シミュレーション値** |

02 と 04 は実際の公開データに置き換えるまでプリセットのままであり、ページ上でもそう明記している。
