'◆注意　本プログラムで発生した万一のトラブル、問題や損害が発生しましても、一切の責任を負いません◆
'スクレイピングは良識の範囲で行い、かつ接続先の規約の有無の確認、およびスクレイピングの実施可否を確認してください
'特にdriver.Waitの値は短い時間を入力しないでください
'(DOS攻撃と勘違いされます)
'
'Google検索の全てのレイアウトパターンに対応していません。
'上手くXpathが読み込めないと無限ループから抜け出せずフリーズします。
'そのためXPathは必要に応じて変更してください。
'また、1ページに10位までが表示されないと、こちらも無限ループから抜け出せません。
'まだまだ改良要ですが、お好きにご使用ください。
'
'キーワードの入力セルはCells(3, 2)としていますが、自由に変更してください。
'「集計結果」「Google」シートがありますが、シート名は何でもOKです。
Sub 検索結果順位タイトル取得()
  '変数定義
  Dim driver As New Selenium.ChromeDriver
  Dim element_SupeUnit As Selenium.WebElements
  
  Dim wsResult As Worksheet
  Set wsResult = Worksheets("集計結果")
  Dim wsSearch As Worksheet
  Set wsSearch = Worksheets("Google")
  
  Dim elements As Selenium.WebElements
  
  Dim strURL As String
  Dim ranking As Long
  
  
  If wsSearch.Cells(3, 2).Value = "" Then
    MsgBox ("黄色の塗りつぶし部分に" & vbCrLf & "検索ワードを入れてください")
  Else
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
    
    '処理開始
    flagSnippet = False '強調スぺニットをカウントしたフラグ
    ranking = 1 '順位
    
    'ドメインまでリンクを取得(トップページ)
    '10位を取得するまで続ける
    Do Until ranking > 10
      tmpStrNum = Replace(Str(ranking), " ", "") '数字をStr関数で文字にするとスペースができるため、スペースを削除
      
      If flagSnippet = False Then
        '一番目に強調スぺニットがあるか確認
        
        tmpStrXPath = "//*[@id=""rso""]/div[" & tmpStrNum & "]/div[1]/div/div[1]/div/div[2]/div/div[1]/a"
        Call GetURLWithXPath(tmpStrXPath, driver, elements, strURL)
        If elements.Count = 1 Then
          'トップページのURLを入れる
          Call InsertTopPageURL(strURL, wsResult, ranking)
        End If
        '通常の1位を取得
        tmpStrXPath = "//*[@id=""rso""]/div[" & tmpStrNum & "]/div/div[1]/a"
        Call GetURLWithXPath(tmpStrXPath, driver, elements, strURL)
        If elements.Count = 1 Then
          'トップページのURLを入れる
          Call InsertTopPageURL(strURL, wsResult, ranking)
        End If
        
        flagSnippet = True '強調スぺニットフラグをON
      Else '強調スぺニットではないとき
        tmpStrXPath = "//*[@id=""rso""]/div[" & tmpStrNum & "]/div/div[1]/a"
        Call GetURLWithXPath(tmpStrXPath, driver, elements, strURL)
        If elements.Count = 1 Then
          '通常の順位を取得
          Call InsertTopPageURL(strURL, wsResult, ranking)
        End If
      End If
    Loop
    
    MsgBox "順位収集完了しました"
  
    '終了処理
    driver.Close
    driver.Quit
    Set driver = Nothing
  End If
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
'機能：XpathのトップURLを取得する
'引数：tmpStr...ページURL(値渡し)
'     wsResult...結果シート(値渡し)
'     ranking...検索順位(参照渡し)
'返り値：ranking...検索順位(参照渡し)
'
'
Function InsertTopPageURL(ByVal tmpStr As String, ByVal wsResult As Worksheet, ByRef ranking As Long)
    'トップページのURLを入れる
    strAddress = Split(tmpStr, "/") '/で文字を分解
    strAddress = strAddress(0) & "//" & strAddress(2) 'http://～～～.comを作成
    'EXCELに値を入れる
    wsResult.Activate
    wsResult.Range(Cells(ranking + 2, 2), Cells(ranking + 2, 2)).Value = strAddress 'セルにURLを入力
    ranking = ranking + 1
End Function

'''''''''''''''''''''''
'★本コードでは使っていません★
'機能：Xpathの場所の有無を確認する
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


