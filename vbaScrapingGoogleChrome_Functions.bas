'
'　VBAでGoogle Chromeからスクレイピングするプログラムの基本形です。
'  処理内容...検索→Xpathの有無を確認、あればURL取得→自分の所望する処理を行う。
'
'  まだまだ親切な形なコードになっていないため、お好きに改良してください。
'
Sub VBAScrapingGoogleChrome_Base()
  '変数定義
  Dim driver As New Selenium.ChromeDriver
  Dim element_SupeUnit As Selenium.WebElements
  Dim elements As Selenium.WebElements

  Dim wsSearch As Worksheet
  Set wsSearch = Worksheets("Google")  

  Dim strURL As String
  
  '初期化処理
  driver.Start
  driver.Get "https://www.google.com"
  driver.Wait (5000) '待機(msec)

  '検索窓にキーワード入力
  wsSearch.Activate 'EXCELにGoogleシートを認識させる(1004エラー対策)
  driver.FindElementByName("q").SendKeys(wsSearch.Cells(3, 2)).Value  '検索ワードはセルから読み取り(注：必要に応じて変える)
  driver.Wait (1500) '念のため待機キーボード入力完了待ち
  SendKeys "{ENTER}" '検索ボタンエンター(Clickは動作しない時があるため)
  driver.Wait (2000) '表示待ち

    
  '処理部(Xpathを探す→あればURLを取得)：この部分に必要に応じてループを加えていく。
      
  tmpStrXPath = "dummy" 'ここにXpathを入れる(dummyのままじゃ動きませんので！)
    Call GetURLWithXPath(tmpStrXPath, driver, elements, strURL) 'Xpathの場所が見つかればカウント(elements, strURLに値が返る)
    If elements.Count = 1 Then
      '↓↓↓↓ここにURLを使ったコードを書いていく↓↓↓↓

      '↑↑↑↑ここにURLを使ったコードを書いていく↑↑↑↑
    End If
  End If

  MsgBox "処理完了(デバッグ機能などで変数の中身をウォッチしてみてください)"

  '終了処理
  driver.Close
  driver.Quit
  Set driver = Nothing
End Sub

'''''''''''''''''''''''
'機能：Xpathの場所のURLを取得する。
'      事前にXpathのチェックを行った後、URLを取得する。
'引数：tmpStrXPath...探したいXpath(値渡し)
'     driver...ChromeDriver変数(値渡し)
'     elements...WebElements変数(参照渡し)
'     strURL...Xpathの示すURL(参照渡し)
'返り値：strURL...Xpathの示すURL
' 
Function GetURLWithXPath(ByVal tmpStrXPath As String, ByVal driver As Selenium.ChromeDriver, _
                            ByRef elements As Selenium.WebElements, ByRef strURL As String)
  
  Set elements = driver.FindElementsByXPath(tmpStrXPath)
  If elements.Count = 1 Then
    strURL = driver.FindElementByXPath(tmpStrXPath).Attribute("href")
  End If

End Function

'''''''''''''''''''''''
'機能：Xpathの場所の有無を確認する(★本コードでは使っていません★)
'引数：tmpStrXPath...探したいXpath(値渡し)
'     driver...ChromeDriver変数(値渡し)
'     elements...WebElements情報(参照渡し)
'返り値：JudgeSuggestWithXPath...bool値
'       True：Xpathの要素あり False：Xpathの要素なし
'
Function JudgeSuggestWithXPath(ByVal tmpStrXPath As String, ByVal driver As Selenium.ChromeDriver, _
                            ByVal elements As Selenium.WebElements) As Boolean
  
  Set elements = driver.FindElementsByXPath(tmpStrXPath)
  If elements.Count = 1 Then
     JudgeSuggestWithXPath = True
  End If

End Function
