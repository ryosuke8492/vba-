Attribute VB_Name = "Module1"
Sub Get1DayAgoRateToExcel()
    Dim http As Object, stream As Object, RE As Object, m As Object
    Dim html As String, plainText As String, idStr As String, url As String
    Dim targetDate As Date
    Dim usdTTS As String, eurTTS As String
    Dim i As Integer

    targetDate = Date - 1   ' 1日前

    ' 土日祝でデータが無い場合に備え、最大5日分さかのぼる
    For i = 1 To 5
        idStr = Format(targetDate, "yy") & Format(targetDate, "mm") & Format(targetDate, "dd")
        url = "https://www.murc-kawasesouba.jp/fx/past/index.php?id=" & idStr

        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.Open "GET", url, False
        http.setRequestHeader "User-Agent", "Mozilla/5.0"
        http.send

        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Write http.responseBody
        stream.Position = 0
        stream.Type = 2
        stream.Charset = "Shift-JIS"
        html = stream.ReadText
        stream.Close

        Set RE = CreateObject("VBScript.RegExp")
        RE.Global = True
        RE.Pattern = "<[^>]+>"
        plainText = RE.Replace(html, " ")

        RE.Global = False
        RE.Pattern = "USD[^\d]*([\d.]+)"
        If RE.Test(plainText) Then
            usdTTS = RE.Execute(plainText)(0).SubMatches(0)

            RE.Pattern = "EUR[^\d]*([\d.]+)"
            eurTTS = RE.Execute(plainText)(0).SubMatches(0)

            Exit For   ' 取得できたら抜ける
        Else
            targetDate = targetDate - 1   ' データなし→もう1日前へ
            usdTTS = "": eurTTS = ""
        End If
    Next i

    If usdTTS = "" Then
        MsgBox "レートが取得できませんでした"
        Exit Sub
    End If

    ' シートに書き込み（セル位置は好きに変更してください）
    With ActiveSheet
        .Range("A1").Value = "日付"
        .Range("B1").Value = "USD/TTS"
        .Range("C1").Value = "EUR/TTS"
        .Range("A2").Value = targetDate
        .Range("B2").Value = CDbl(usdTTS)
        .Range("C2").Value = CDbl(eurTTS)
    End With

    MsgBox Format(targetDate, "yyyy/mm/dd") & "のレートを入力しました"
End Sub
