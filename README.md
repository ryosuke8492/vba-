# MUFG Rate Fetcher (VBA)

三菱UFJ銀行・三菱UFJリサーチ&コンサルティングの公開データから
USD/EURのT.T.S.（電信売相場）を取得するExcel VBAマクロです。

## ファイル
- `Get1DayAgoRateToExcel.bas` : 1営業日前（土日祝はさらに遡る）のUSD/EUR TTSを取得し、
  アクティブシートの A1:C2 に書き込みます。

## 使い方
1. Excelで `Alt+F11` を押しVBEを開く
2. 「ファイル」→「ファイルのインポート」で `Get1DayAgoRateToExcel.bas` を読み込む
3. `Get1DayAgoRateToExcel` マクロを実行

## データソース
https://www.murc-kawasesouba.jp/fx/past_3month.php
