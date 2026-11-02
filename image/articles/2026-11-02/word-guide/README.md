# Word記事の追加画像（2026-11-02）

Windows版Word / Excelの実画面です。画面を生成AIで再現したものは含みません。
本文に掲載する箇所を切り出しています。Excelは150%表示で撮影しました。

|画像|内容|
|---|---|
|excel-postal-date.png|郵便番号と日付の元列・TEXT列|
|excel-money-percent.png|金額12000→12,000円、割合0.08→8%|
|excel-print-columns.png|上の印刷用列の撮影時の全体画面|
|replace-paragraphs.png|検索と置換の詳細設定。入力前の画面|
|compare-dialog.png|元の文書・変更された文書を選ぶ画面。指定前|
|compare-options.png|比較の詳細条件と新規文書への出力指定|
|textbox-menu.png|横書き・縦書きテキストボックスの描画メニュー|
|paste-menu.png|ホームの貼り付けメニュー|
|word-pasted-table.png|Excelの住所録をWordの表へ通常貼り付けした結果|
|paste-link-dialog.png|リンク貼り付け・Excelオブジェクトの設定。更新動作の証明ではない|
|word-qr-code.png|DISPLAYBARCODEで固定値TEST-1001を指定したコード|
|word-qr-result.png|同じコードを更新したWord上のQR表示|

QRの表示例は固定値の生成です。MERGEBARCODEによる各レコードの一括出力、紙への印刷・実機読み取りとは区別しています。
比較画面の画像は設定説明用で、2文書を比較した結果の画面ではありません。
撮影用の氏名・住所・番号は架空です。切り出し前の追加画像は非公開の `_dev/word-guide-raw` に保存しています。

## 画像形式・寸法の修正（2026-11-02）

compare-dialog、compare-options、excel-print-columns、paste-link-dialog、paste-menu、replace-paragraphs、textbox-menu は、当初PNG拡張子のJPEGだったため、実際のPNG形式へ統一しました。貼り付け記事のメニュー・リンク設定画面に入っていた異常な画像寸法も、実ファイルの寸法へ修正しました。形式の統一は元画像の文字を高精細化する処理ではありません。

ラベル・封筒の撮り直しと原寸ビューアーの修正は [word-image-fixes/README.md](../word-image-fixes/README.md) を参照してください。
