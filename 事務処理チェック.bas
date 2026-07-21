Option Explicit

'==============================================================
' 事務処理日チェック VBA サンプル
' ボタンから Main を呼び出す
'==============================================================

' ---------------------------------------------------------------
' ユーティリティ関数：営業日加算（土日祝除外）
'   baseDate  : 起算日
'   n         : 加算営業日数
'   holidays  : 祝日リスト（省略可）
' ---------------------------------------------------------------
Public Function AddBusinessDays(baseDate As Date, n As Long, _
                                Optional holidays As Variant) As Date
    Dim d As Date
    Dim added As Long
    Dim hasHolidays As Boolean

    hasHolidays = Not IsMissing(holidays)
    d = baseDate
    added = 0

    Do While added < n
        d = d + 1
        If Weekday(d, vbMonday) <= 5 Then          ' 月〜金
            If hasHolidays Then
                If Not IsHoliday(d, holidays) Then  ' 祝日でない
                    added = added + 1
                End If
            Else
                added = added + 1
            End If
        End If
    Loop

    AddBusinessDays = d
End Function

' ---------------------------------------------------------------
' ユーティリティ関数：カレンダー日加算（単純 +N 日）
'   baseDate : 起算日
'   n        : 加算日数
' ---------------------------------------------------------------
Public Function AddCalendarDays(baseDate As Date, n As Long) As Date
    AddCalendarDays = baseDate + n
End Function

' ---------------------------------------------------------------
' ユーティリティ関数：祝日判定
' ---------------------------------------------------------------
Private Function IsHoliday(d As Date, holidays As Variant) As Boolean
    Dim i As Long
    For i = LBound(holidays) To UBound(holidays)
        If d = CDate(holidays(i)) Then
            IsHoliday = True
            Exit Function
        End If
    Next i
    IsHoliday = False
End Function

' ---------------------------------------------------------------
' 判定結果を書き込むヘルパー
'   cell      : 対象セル
'   result    : 1（期限内）/ 0（期限外）/ -1（空白）
' ---------------------------------------------------------------
Private Sub WriteResult(cell As Range, result As Long)
    Const COLOR_OK   As Long = &HC6EFCE  ' 緑系（期限内）
    Const COLOR_NG   As Long = &HFFC7CE  ' 赤系（期限外）
    Const COLOR_NONE As Long = &HFFFFFF  ' 白（空白）

    Select Case result
        Case 1
            cell.Value = 1
            cell.Interior.Color = COLOR_OK
        Case 0
            cell.Value = 0
            cell.Interior.Color = COLOR_NG
        Case Else  ' -1 = 空白
            cell.Value = ""
            cell.Interior.Color = COLOR_NONE
    End Select
End Sub

'==============================================================
' シートレイアウト想定
'   1行目 : ヘッダー
'   A列   : 作業日
'   B列   : 書類A完了日  条件：作業日 +1 営業日以内
'   C列   : 書類B完了日  条件：作業日 +2 営業日以内
'   D列   : 書類C完了日  条件：作業日 +3 営業日以内
'   E列   : 書類D完了日  条件：作業日 +5 営業日以内
'   F列   : 書類E完了日  条件：作業日 +7 カレンダー日以内
'   G列   : 書類F完了日  条件：作業日 +14 カレンダー日以内
'   H列   : 書類G完了日  条件：作業日 +30 カレンダー日以内
'   I列   : 書類H完了日  条件：作業日 +60 カレンダー日以内
'   J〜Q列: 各書類の判定結果列（1/0/空白）
'==============================================================

' ---------------------------------------------------------------
' メイン処理：ボタンから呼び出す
' ---------------------------------------------------------------
Public Sub Main()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim workDate As Date
    Dim completionDate As Date
    Dim deadline As Date
    Dim resultCol As Long
    Dim completionCell As Range
    Dim resultCell As Range

    ' 祝日リスト（必要に応じて追記・Sheet化も可）
    Dim holidays As Variant
    holidays = Array("2025/1/1", "2025/1/13", "2025/2/11", _
                     "2025/2/23", "2025/3/20", "2025/4/29", _
                     "2025/5/3", "2025/5/4", "2025/5/5", _
                     "2025/7/21", "2025/8/11", "2025/9/15", _
                     "2025/9/23", "2025/10/13", "2025/11/3", _
                     "2025/11/23")

    Set ws = ThisWorkbook.Sheets("Sheet1")  ' シート名を合わせる

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 処理前にステータス表示
    Application.ScreenUpdating = False
    Application.StatusBar = "処理中..."

    For i = 2 To lastRow  ' 2行目からデータ

        ' 作業日が空なら skip
        If IsEmpty(ws.Cells(i, 1)) Or Not IsDate(ws.Cells(i, 1)) Then
            GoTo NextRow
        End If
        workDate = CDate(ws.Cells(i, 1).Value)

        '------------------------------------------------------------
        ' 各書類の判定ブロック
        ' (書類列, 結果列, 期限計算方法)
        '------------------------------------------------------------

        ' 書類A：+1 営業日
        Call CheckItem(ws.Cells(i, 2), ws.Cells(i, 10), _
                       AddBusinessDays(workDate, 1, holidays))

        ' 書類B：+2 営業日
        Call CheckItem(ws.Cells(i, 3), ws.Cells(i, 11), _
                       AddBusinessDays(workDate, 2, holidays))

        ' 書類C：+3 営業日
        Call CheckItem(ws.Cells(i, 4), ws.Cells(i, 12), _
                       AddBusinessDays(workDate, 3, holidays))

        ' 書類D：+5 営業日
        Call CheckItem(ws.Cells(i, 5), ws.Cells(i, 13), _
                       AddBusinessDays(workDate, 5, holidays))

        ' 書類E：+7 カレンダー日
        Call CheckItem(ws.Cells(i, 6), ws.Cells(i, 14), _
                       AddCalendarDays(workDate, 7))

        ' 書類F：+14 カレンダー日
        Call CheckItem(ws.Cells(i, 7), ws.Cells(i, 15), _
                       AddCalendarDays(workDate, 14))

        ' 書類G：+30 カレンダー日
        Call CheckItem(ws.Cells(i, 8), ws.Cells(i, 16), _
                       AddCalendarDays(workDate, 30))

        ' 書類H：+60 カレンダー日
        Call CheckItem(ws.Cells(i, 9), ws.Cells(i, 17), _
                       AddCalendarDays(workDate, 60))

NextRow:
    Next i

    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "チェック完了しました。", vbInformation

End Sub

' ---------------------------------------------------------------
' 1項目の判定＆書き込み
'   srcCell  : 完了日が入力されているセル
'   dstCell  : 結果（1/0/空白）を書くセル
'   deadline : 期限日（AddBusinessDays or AddCalendarDays で計算済み）
' ---------------------------------------------------------------
Private Sub CheckItem(srcCell As Range, dstCell As Range, deadline As Date)
    If IsEmpty(srcCell) Or srcCell.Value = "" Then
        ' 未入力 → 空白のまま
        Call WriteResult(dstCell, -1)
    ElseIf Not IsDate(srcCell.Value) Then
        ' 日付として認識できない → 空白
        Call WriteResult(dstCell, -1)
    ElseIf CDate(srcCell.Value) <= deadline Then
        ' 期限以内 → 1
        Call WriteResult(dstCell, 1)
    Else
        ' 期限超過 → 0
        Call WriteResult(dstCell, 0)
    End If
End Sub

'==============================================================
' シート初期設定（ヘッダー＋書式設定）
' 初回セットアップ用：別ボタンに割り当てる or 手動実行
'==============================================================
Public Sub SetupSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' ヘッダー
    Dim headers As Variant
    headers = Array("作業日", _
                    "書類A完了日", "書類B完了日", "書類C完了日", "書類D完了日", _
                    "書類E完了日", "書類F完了日", "書類G完了日", "書類H完了日", _
                    "A結果(+1営業)", "B結果(+2営業)", "C結果(+3営業)", "D結果(+5営業)", _
                    "E結果(+7日)", "F結果(+14日)", "G結果(+30日)", "H結果(+60日)")

    Dim col As Long
    For col = 1 To UBound(headers) + 1
        ws.Cells(1, col).Value = headers(col - 1)
        ws.Cells(1, col).Font.Bold = True
    Next col

    ' 結果列（J〜Q = 10〜17列）にパーセンテージ書式
    ws.Range(ws.Cells(2, 10), ws.Cells(1000, 17)).NumberFormat = "0%"

    ' 列幅自動調整
    ws.Columns("A:Q").AutoFit

    MsgBox "シート初期設定完了。", vbInformation
End Sub
