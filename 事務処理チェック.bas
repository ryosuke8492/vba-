Option Explicit

'==============================================================
' 事務処理日チェック VBA サンプル
' ボタンから Main を呼び出す
'==============================================================

' ---------------------------------------------------------------
' 日付文字列から末尾（終了）日付を抽出する
'
' 対応フォーマット：
'   "7/22"         → 7/22
'   " 7/22 "       → 7/22  （前後スペース除去）
'   "7/22 - 23"    → 7/23  （終端が日だけ → 開始日の月を補完）
'   "7/22-23"      → 7/23
'   "7/22～7/23"   → 7/23
'   "7/22~7/23"    → 7/23
'
' 戻り値：
'   解析成功 → Date 型の終了日（年は実行時の年）
'   解析失敗 → 0（= #1899/12/30#）を返す ※呼び出し元で IsError 等確認推奨
' ---------------------------------------------------------------
Public Function ExtractEndDate(ByVal input As String) As Date
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim cleaned   As String  ' トリム済み入力文字列
    Dim separator As String  ' 検出したセパレータ（～ / ~ / -）
    Dim sepPos    As Long    ' セパレータの位置
    Dim startPart As String  ' セパレータ左側（開始日）
    Dim endPart   As String  ' セパレータ右側（終了日 or 日のみ）
    Dim slashPos  As Long    ' startPart 内の "/" 位置
    Dim monthStr  As String  ' 開始日から取り出した月文字列
    ' ────────────────────────────────────────────────────

    On Error GoTo ErrHandler

    cleaned = Trim(input)

    ' ── セパレータ検出（優先順位：～ > ~ > -）──────────
    If InStr(cleaned, "～") > 0 Then
        separator = "～"
    ElseIf InStr(cleaned, "~") > 0 Then
        separator = "~"
    ElseIf InStr(cleaned, "-") > 0 Then
        separator = "-"
    Else
        ' セパレータなし → そのまま日付として解析
        ExtractEndDate = CDate(cleaned)
        Exit Function
    End If

    ' ── 開始・終了に分割 ─────────────────────────────
    sepPos    = InStr(cleaned, separator)
    startPart = Trim(Left(cleaned, sepPos - 1))
    endPart   = Trim(Mid(cleaned, sepPos + Len(separator)))

    ' ── 終端が数字のみ（日だけ）なら月を補完 ──────────
    ' 例："7/22 - 23" → endPart="23" → "7/23"
    If IsNumeric(endPart) Then
        slashPos = InStr(startPart, "/")
        If slashPos > 0 Then
            monthStr = Left(startPart, slashPos - 1)
            endPart  = monthStr & "/" & endPart
        End If
    End If

    ExtractEndDate = CDate(endPart)
    Exit Function

ErrHandler:
    ExtractEndDate = 0  ' 解析失敗時は 0 を返す
End Function

' ---------------------------------------------------------------
' 日付文字列から末尾（終了）日付を抽出する（余計な文字列除去版）
'
' 正規表現（Regular Expression。文字列の「パターン」を表すミニ言語。
' 例： \d は数字1文字、{1,2} は直前を1〜2回、| は「または」の意味）
' を使って文字列中の数値パターン（n/n や n）だけを拾い出すため、
' "open" "colse" のような日付以外の文字列が混ざっていても
' 末尾の日付だけを取り出せる。
'
' 対応フォーマット例：
'   "open7/1 - colse 7/3" → 7/3
'   "7/22 - 23"           → 7/23  （終端が日だけ → 開始日の月を補完）
'   "7/22～7/23"          → 7/23
'   "7/22"                → 7/22
'
' 戻り値：
'   解析成功 → Date 型の終了日（年は実行時の年）
'   解析失敗 → 0（= #1899/12/30#）を返す
' ---------------------------------------------------------------
Public Function ExtractEndDateEx(ByVal input As String) As Date
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim re       As Object  ' 正規表現エンジン本体（VBScript.RegExp オブジェクト）
                             ' ※参照設定不要でCreateObjectだけで使える（遅延バインディング）
    Dim mc       As Object  ' MatchCollection（マッチ結果の集合）の略。
                             ' re.Execute(input) を実行すると、パターンに一致した
                             ' 箇所が全部まとめて入ったコレクションが返ってくる。
                             ' mc.Count       … 見つかった数
                             ' mc(0), mc(1)…  … 1つ1つの一致（Matchオブジェクト）
                             ' mc(i).Value    … その一致箇所の実際の文字列
    Dim lastPart As String  ' 最後に見つかった数値部分（終了日候補）
    Dim prevPart As String  ' 1つ前に見つかった数値部分（開始日候補）
    Dim slashPos As Long    ' prevPart 内の "/" 位置
    Dim monthStr As String  ' 開始日から取り出した月文字列
    ' ────────────────────────────────────────────────────

    On Error GoTo ErrHandler

    Set re = CreateObject("VBScript.RegExp")  ' 正規表現エンジンを生成
    re.Global = True                          ' True＝文字列全体を検索して全一致を集める
                                               '（Falseだと最初の1件しか見つからない）
    ' パターンの意味：
    '   \d{1,2}/\d{1,2}  … "7/22" のような "数字/数字"（月/日）の形
    '   |                … 「または」
    '   \d{1,2}          … "23" のような数字だけの形（月が省略された終了日用）
    re.Pattern = "\d{1,2}/\d{1,2}|\d{1,2}"

    Set mc = re.Execute(input)  ' input 内でパターンに一致する箇所を全部検索してmcに格納

    If mc.Count = 0 Then        ' 一致が1つもない＝日付らしき文字列が見つからなかった
        ExtractEndDateEx = 0
        Exit Function
    End If

    ' mc.Count - 1 が「最後に見つかった一致」のインデックス（0始まりのため）
    lastPart = mc(mc.Count - 1).Value

    ' ── 終端が数字のみ（日だけ）なら、直前のマッチから月を補完 ──
    If InStr(lastPart, "/") = 0 And mc.Count >= 2 Then
        prevPart = mc(mc.Count - 2).Value
        slashPos = InStr(prevPart, "/")
        If slashPos > 0 Then
            monthStr = Left(prevPart, slashPos - 1)
            lastPart = monthStr & "/" & lastPart
        End If
    End If

    ExtractEndDateEx = CDate(lastPart)
    Exit Function

ErrHandler:
    ExtractEndDateEx = 0  ' 解析失敗時は 0 を返す
End Function

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
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim ws       As Worksheet  ' 処理対象シート
    Dim lastRow  As Long       ' データ最終行
    Dim i        As Long       ' 行ループカウンタ
    Dim workDate As Date       ' 作業日（A列）
    Dim holidays As Variant    ' 祝日リスト
    ' ────────────────────────────────────────────────────

    ' 祝日リスト（必要に応じて追記・Sheet化も可）
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

' ---------------------------------------------------------------
' 1項目の判定＆書き込み（許容日数版）
'   CheckItem が「targetDate 以前ならOK」なのに対し、こちらは
'   「targetDate との差が toleranceDays 以内ならOK」という判定。
'   toleranceDays を省略（＝0）すると、targetDate と完全に一致した
'   日付のみOKになる（早すぎても遅すぎてもNG）。
'
'   srcCell       : 完了日が入力されているセル
'   dstCell       : 結果（1/0/空白）を書くセル
'   targetDate    : 基準日（この日 ±toleranceDays 以内ならOK）
'   toleranceDays : 許容日数（省略時 0 ＝ 基準日ちょうどのみOK）
' ---------------------------------------------------------------
Private Sub CheckItemTolerance(srcCell As Range, dstCell As Range, _
                                targetDate As Date, _
                                Optional ByVal toleranceDays As Long = 0)
    Dim diffDays As Long  ' 完了日と基準日の差（絶対値・日数）

    If IsEmpty(srcCell) Or srcCell.Value = "" Then
        ' 未入力 → 空白のまま
        Call WriteResult(dstCell, -1)
    ElseIf Not IsDate(srcCell.Value) Then
        ' 日付として認識できない → 空白
        Call WriteResult(dstCell, -1)
    Else
        diffDays = Abs(CDate(srcCell.Value) - targetDate)
        If diffDays <= toleranceDays Then
            ' 基準日 ±toleranceDays 以内 → 1
            Call WriteResult(dstCell, 1)
        Else
            ' 許容範囲外（早すぎ／遅すぎ） → 0
            Call WriteResult(dstCell, 0)
        End If
    End If
End Sub

'==============================================================
' シート初期設定（ヘッダー＋書式設定）
' 初回セットアップ用：別ボタンに割り当てる or 手動実行
'==============================================================
Public Sub SetupSheet()
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim ws      As Worksheet  ' 処理対象シート
    Dim headers As Variant    ' ヘッダー文字列配列
    Dim col     As Long       ' 列ループカウンタ
    ' ────────────────────────────────────────────────────

    Set ws = ThisWorkbook.Sheets("Sheet1")

    headers = Array("作業日", _
                    "書類A完了日", "書類B完了日", "書類C完了日", "書類D完了日", _
                    "書類E完了日", "書類F完了日", "書類G完了日", "書類H完了日", _
                    "A結果(+1営業)", "B結果(+2営業)", "C結果(+3営業)", "D結果(+5営業)", _
                    "E結果(+7日)", "F結果(+14日)", "G結果(+30日)", "H結果(+60日)")

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
