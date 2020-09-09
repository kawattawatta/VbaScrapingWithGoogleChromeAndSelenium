'☆注意☆
'　VBAでGoogle Chromeからプログラムの基本形です。
'  順位を1つとります。
Sub FunctionsFotScrapingGoogleChrome()
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
  driver.Wait (5000) '待機(2秒)

  '検索窓にキーワード入力
  driver.FindElementByName("q").SendKeys(wsSearch.Cells(3, 2)).Value  '検索ワードはセルから読み取り(★必要に応じて変えてください)
  driver.Wait (1500) '待機(1.5秒)キーボード入力完了待ち
  SendKeys "{ENTER}" '検索ボタンエンター(Clickは動作しない時があるため)
  driver.Wait (2000) '表示待ち(2秒)

    
  '処理部(Xpathを探す→あればURLを取得)：この部分に必要に応じてループを加えていく。
  tmpStrXPath = "dummy" 'ここにXpathを入れる
    Call GetURLWithXPath(tmpStrXPath, driver, elements, strURL) 'Xpathの場所が見つかればカウント(elements, strURLに値が返る)
    If elements.Count = 1 Then 'XpathのURLあれば取得 
      Call InsertTopPageURL(strURL, wsResult, ranking)
    End If
  End If

  MsgBox "処理完了(デバッグ機能などで変数の中身をウォッチしてみてください)"

  '終了処理
  driver.Close
  driver.Quit
  Set driver = Nothing
End Sub

Function JudgeSuggestWithXPath(ByVal tmpStrXPath As String, ByVal driver As Selenium.ChromeDriver, _
                            ByVal elements As Selenium.WebElements) As Boolean
  
  Set elements = driver.FindElementsByXPath(tmpStrXPath)
  If elements.Count = 1 Then
     JudgeSuggestWithXPath = True
  End If

End Function


Function GetURLWithXPath(ByVal tmpStrXPath As String, ByVal driver As Selenium.ChromeDriver, _
                            ByRef elements As Selenium.WebElements, ByRef strURL As String)
  
  Set elements = driver.FindElementsByXPath(tmpStrXPath)
  If elements.Count = 1 Then
    strURL = driver.FindElementByXPath(tmpStrXPath).Attribute("href")
  End If

End Function


Function InsertTopPageURL(ByVal tmpStr As String, ByVal wsResult As Worksheet, ranking)
    'トップページのURLを入れる
    strAddress = Split(tmpStr, "/") '/で文字を分解
    strAddress = strAddress(0) & "//" & strAddress(2) 'http://～～～.comを作成
    'EXCELに値を入れる
    wsResult.Activate
    wsResult.Range(Cells(ranking + 2, 2), Cells(ranking + 2, 2)).Value = strAddress 'セルにURLを入力
    ranking = ranking + 1
End Function
