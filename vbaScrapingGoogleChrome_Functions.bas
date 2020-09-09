'☆注意☆
'　このコードは関数をまとめたものです。
'
Sub FunctionsFotScrapingGoogleChrome()
  '変数定義
  Dim driver As New Selenium.ChromeDriver
  Dim element_SupeUnit As Selenium.WebElements
  Dim elements As Selenium.WebElements
  
  Dim strURL As String
  
  'Googleアクセス
  driver.Start
  driver.Get "https://www.google.com"
  driver.Wait (5000) '待機(2秒)
  'キーワード入力
  wsSearch.Activate 'EXCELにGoogleシートを認識させる(1004エラー対策)

  driver.FindElementByName("q").SendKeys(wsSearch.Cells(3, 2)).Value  '検索ワードはセルから読み取り
  driver.Wait (1500) '待機(1.5秒)キーボード入力完了待ち

  '検索ボタンエンター(検索ボタンのXPathクリックでは文字長さの状態によってクリックできない場合があった)
  SendKeys "{ENTER}"
  driver.Wait (2000) 'サジェスト表示分を加味して待機(2秒)

  tmpStrXPath = "dummy" 'ここにXpathを入れます
    Call GetURLWithXPath(tmpStrXPath, driver, elements, strURL) 'Xpathの場所が見つかればカウント(elements, strURLに値が返る)
    If elements.Count = 1 Then 'XpathのURLあれば取得 
      Call InsertTopPageURL(strURL, wsResult, ranking)
    End If
  End If

  MsgBox "順位収集完了しました"

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
