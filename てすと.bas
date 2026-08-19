Sub 転記処理()

    Dim wsData As Worksheet
    Dim wsFormat As Worksheet
    Dim wsNew As Worksheet
    Dim targetHeaders() As Variant
    Dim dataHeaderRow As Long, formatHeaderRow As Long
    Dim dataStartRow As Long, formatStartRow As Long
    Dim lastRowData As Long, lastColData As Long
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

    ' ここまでの列はヘッダー名で一致させて転記(とりあえず3つ)
    targetHeaders = Array("項目A", "項目B", "項目C")

    ' データ側:ヘッダーマッチング対象は何列目まで
    Dim dataHeaderMatchLastCol As Long
    dataHeaderMatchLastCol = 3   ' 例:データのA〜C列まではヘッダー一致

    ' データ側:直接コピー開始列(ヘッダーマッチング対象の次から)
    Dim dataDirectCopyStartCol As Long
    dataDirectCopyStartCol = dataHeaderMatchLastCol + 1

    ' フォーマット側:直接コピーを貼り付け始める列
    Dim formatDirectCopyStartCol As Long
    formatDirectCopyStartCol = 5   ' 例:フォーマットのE列以降にそのまま貼る

    '=== フォーマットシートをコピーして新シート作成 ===
    wsFormat.Copy After:=wsFormat
    Set wsNew = ActiveSheet
    wsNew.Name = "フォーマット_" & Format(Now, "yyyymmdd_hhnnss")

    '=== データの最終行・最終列を取得 ===
    lastRowData = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastColData = wsData.Cells(dataHeaderRow, wsData.Columns.Count).End(xlToLeft).Column

    '=== 行ごとに処理 ===
    For j = dataStartRow To lastRowData

        '--- 前半:ヘッダー名一致で転記 ---
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

            wsNew.Cells(formatStartRow + (j - dataStartRow), colFormat).Value = _
                wsData.Cells(j, colData).Value

NextHeader:
        Next i

        '--- 後半:指定列以降はそのまま位置コピー ---
        Dim offset As Long
        offset = 0
        For colData = dataDirectCopyStartCol To lastColData
            wsNew.Cells(formatStartRow + (j - dataStartRow), formatDirectCopyStartCol + offset).Value = _
                wsData.Cells(j, colData).Value
            offset = offset + 1
        Next colData

    Next j

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
