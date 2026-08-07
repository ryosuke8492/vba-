Attribute VB_Name = "Module1"
Sub Get1DayAgoRateToExcel()
    On Error GoTo ErrHandler

    Dim http As Object, stream As Object, RE As Object
    Dim html As String, plainText As String, idStr As String, url As String
    Dim targetDate As Date
    Dim usdTTS As String, eurTTS As String
    Dim i As Integer

    ' --- RegExpが使えるか事前チェック ---
    On Error Resume Next
    Set RE = CreateObject("VBScript.RegExp")
    On Error GoTo ErrHandler
    If RE Is Nothing Then
        MsgBox "正規表現エンジン(VBScript.RegExp)が利用できません。" & vbCrLf & _
               "セキュリティポリシーでVBScriptが無効化されている可能性があります。" & vbCrLf & _
               "IT部門に確認してください。", vbCritical
        Exit Sub
    End If

    targetDate = Date - 1

    For i = 1 To 5
        idStr = Format(targetDate, "yy") & Format(targetDate, "mm") & Format(targetDate, "dd")
        url = "https://www.murc-kawasesouba.jp/fx/past/index.php?id=" & idStr

        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.setTimeouts 5000, 5000, 5000, 5000   ' 接続/送信/受信のタイムアウトを明示
        http.Open "GET", url, False
        http.setRequestHeader "User-Agent", "Mozilla/5.0"
        http.send

        If http.Status <> 200 Then
            MsgBox "通信に失敗しました。Status: " & http.Status & vbCrLf & _
                   "社内ネットワークの制限（プロキシ/ファイアウォール）の可能性があります。", vbExclamation
            Exit Sub
        End If

        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Write http.responseBody
        stream.Position = 0
        stream.Type = 2
        stream.Charset = "Shift-JIS"
        html = stream.ReadText
        stream.Close

        RE.Global = True
        RE.Pattern = "<[^>]+>"
        plainText = RE.Replace(html, " ")

        RE.Global = False
        RE.Pattern = "USD[^\d]*([\d.]+)"
        If RE.Test(plainText) Then
            usdTTS = RE.Execute(plainText)(0).SubMatches(0)
            RE.Pattern = "EUR[^\d]*([\d.]+)"
            eurTTS = RE.Execute(plainText)(0).SubMatches(0)
            Exit For
        Else
            targetDate = targetDate - 1
            usdTTS = "": eurTTS = ""
        End If
    Next i

    If usdTTS = "" Then
        MsgBox "レートが取得できませんでした（対象期間外の可能性があります）"
        Exit Sub
    End If

    With ActiveSheet
        .Range("A1").Value = "日付"
        .Range("B1").Value = "USD/TTS"
        .Range("C1").Value = "EUR/TTS"
        .Range("A2").Value = targetDate
        .Range("B2").Value = CDbl(usdTTS)
        .Range("C2").Value = CDbl(eurTTS)
    End With

    MsgBox Format(targetDate, "yyyy/mm/dd") & "のレートを入力しました"
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "内容: " & Err.Description & vbCrLf & vbCrLf & _
           "この画面のスクリーンショットを送ってもらえれば原因を特定できます。", vbCritical
End Sub
