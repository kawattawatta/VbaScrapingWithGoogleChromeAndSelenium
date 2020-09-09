# VbaScrapingWithGoogleChromeAndSelenium

VBAでスクレイピングするソースです。

まだ問題ありますが、ある程度使えるレベルまでいったので公開します。

【問題点】

・10位を取得するまでループが止まらない

・検索結果のレイアウトによりXpathを変更要

【ファイルの説明】
vbaScrapingGoogleChrome.bas
→Googleの検索順位TOP10を集計結果シートに入力します。

vbaScrapingGoogleChrome_Functions.bas
→Googleの検索結果の順位を取得するための基本形(ひな形)です。

【動作環境】
Office：EXCEL2007

OS：windows8.1 64bit x64ベースプロセッサ

CPU：Intel Celeron CPU1000M 1.80GHz,1.80GHz

メモリ：4GB

ブラウザ：Google Chrome

ライブラリ：SeleniumBasic

細かい注意点は下記を参照(QiitaHP)。

参考URL様
https://qiita.com/400800mkouyou/items/735704557e52bd5c08dc
