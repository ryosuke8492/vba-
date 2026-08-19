Sub 転記処理()

    Dim wsData As Worksheet
    Dim wsFormat As Worksheet
    Dim wsNew As Worksheet
    Dim targetHeaders() As Variant
    Dim dataHeaderRow As Long, formatHeaderRow As Long
    Dim dataStartRow As Long, formatStartRow As Long
    Dim lastRowData As Long
    Dim i As Long, j As Long
    Dim colData As Long, colFormat As Long
    Dim headerName As String

    '=== 設定項目 ===
    Set wsData = ThisWorkbook.Sheets("データ")
    Set wsFormat = ThisWorkbook.Sheets("フォーマット")

    dataHeaderRow = 1      ' データシートのヘッダー行
    formatHeaderRow = 1    ' フォーマットシートのヘッダー行
    dataStartRow = 2       ' データシートの取得開始行(ヘッダー以降)
    formatStartRow = 2     ' フォーマットシートの貼り付け開始行

    ' 転記したいヘッダーをとりあえず3つ指定
    targetHeaders = Array("項目A", "項目B", "項目C")

    '=== フォーマットシートをコピーして新シート作成 ===
    wsFormat.Copy After:=wsFormat
    Set wsNew = ActiveSheet
    wsNew.Name = "フォーマット_" & Format(Now, "yyyymmdd_hhnnss")

    '=== データの最終行を取得 ===
    lastRowData = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    '=== ヘッダーごとにマッピングして転記 ===
    For i = LBound(targetHeaders) To UBound(targetHeaders)
        headerName = targetHeaders(i)

        colData = FindHeaderColumn(wsData, dataHeaderRow, headerName)
        colFormat = FindHeaderColumn(wsNew, formatHeaderRow, headerName)

        If colData = 0 Then
            MsgBox "データシートにヘッダー『" & headerName & "』が見つかりません", vbExclamation
            GoTo NextHeader
        End If
        If colFormat = 0 Then
            MsgBox "フォーマットシートにヘッダー『" & headerName & "』が見つかりません", vbExclamation
            GoTo NextHeader
        End If

        ' データを1行ずつ転記(まとめてコピーしたい場合は下のコメント参照)
        For j = dataStartRow To lastRowData
            wsNew.Cells(formatStartRow + (j - dataStartRow), colFormat).Value = _
                wsData.Cells(j, colData).Value
        Next j

NextHeader:
    Next i

    MsgBox "転記が完了しました:" & wsNew.Name, vbInformation

End Sub

'=== ヘッダー名から列番号を取得する関数 ===
Function FindHeaderColumn(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If Trim(ws.Cells(headerRow, c).Value) = Trim(headerName) Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c

    FindHeaderColumn = 0 ' 見つからない場合
End Function
