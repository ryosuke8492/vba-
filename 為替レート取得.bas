Option Explicit

'==============================================================
' 三菱UFJ銀行 為替レート取得 VBA サンプル
' 米ドル(USD)・ユーロ(EUR)の T.T.S.（電信売相場）を取得しシートへ書き込む
'
' データ取得元：三菱UFJ銀行 公開CSV（Shift-JIS）
'   https://www.bk.mufg.jp/gdocs/kinri/list_j/kinri/spot_rate.csv
' CSV列構成：
'   通貨コード, 通貨名, T.T.S., ACC., CASH S., T.T.B., A/S, D/P・D/A, CASH B.
'   例）"001","USD (ドル)","164.12","164.54","165.92","162.12","161.70","161.40","160.12"
'
' 注意：参照設定は不要（CreateObject による遅延バインディングのみ使用）
'==============================================================

Private Const MUFG_CSV_URL As String = "https://www.bk.mufg.jp/gdocs/kinri/list_j/kinri/spot_rate.csv"
Private Const OUTPUT_SHEET_NAME As String = "為替レート"

' ---------------------------------------------------------------
' メイン処理：ボタンから呼び出す
'   三菱UFJ銀行のCSVからUSD/EURのTTSを取得し、シートに書き込む
' ---------------------------------------------------------------
Public Sub GetMufgFxRates()
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim csvText As String   ' CSV全文（Shift-JIS→変換済み）
    Dim usdTts  As Variant  ' USD の T.T.S.（取得失敗時 Null）
    Dim eurTts  As Variant  ' EUR の T.T.S.（取得失敗時 Null）
    ' ────────────────────────────────────────────────────

    On Error GoTo ErrHandler

    Application.StatusBar = "為替レート取得中..."

    csvText = DownloadCsvAsShiftJis(MUFG_CSV_URL)

    usdTts = ExtractTtsRate(csvText, "USD")
    eurTts = ExtractTtsRate(csvText, "EUR")

    If IsNull(usdTts) Or IsNull(eurTts) Then
        MsgBox "USDまたはEURのレートがCSVから見つかりませんでした。" & vbCrLf & _
               "サイトの構成が変更された可能性があります。", vbExclamation
        GoTo CleanExit
    End If

    Call WriteRatesToSheet(CDbl(usdTts), CDbl(eurTts))

    MsgBox "為替レートを取得しました。" & vbCrLf & _
           "USD TTS：" & usdTts & vbCrLf & _
           "EUR TTS：" & eurTts, vbInformation

CleanExit:
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "為替レートの取得に失敗しました。" & vbCrLf & Err.Description, vbCritical
End Sub

' ---------------------------------------------------------------
' 指定URLのCSVをバイナリ取得し、Shift-JISとしてテキスト変換する
'   url : 取得先CSVのURL
'   戻り値：変換後のCSV全文
' ---------------------------------------------------------------
Private Function DownloadCsvAsShiftJis(ByVal url As String) As String
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim http   As Object  ' MSXML2.XMLHTTP（HTTP取得用）
    Dim stream As Object  ' ADODB.Stream（文字コード変換用）
    ' ────────────────────────────────────────────────────

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.Send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1, "DownloadCsvAsShiftJis", _
                  "HTTPステータス " & http.Status & "：CSVを取得できませんでした。"
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1              ' adTypeBinary
    stream.Open
    stream.Write http.responseBody
    stream.Position = 0
    stream.Type = 2              ' adTypeText
    stream.Charset = "Shift_JIS" ' 三菱UFJ銀行CSVの文字コード
    DownloadCsvAsShiftJis = stream.ReadText
    stream.Close
End Function

' ---------------------------------------------------------------
' CSV全文から指定通貨コードの T.T.S. を抽出する
'   csvText      : CSV全文
'   currencyCode : "USD" や "EUR" など
'   戻り値：TTSレート（Double）／該当行なしの場合 Null
' ---------------------------------------------------------------
Private Function ExtractTtsRate(ByVal csvText As String, ByVal currencyCode As String) As Variant
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim lines()  As String  ' CSVを行単位に分割した配列
    Dim fields() As String  ' 1行をカンマ区切りで分割した配列
    Dim i        As Long    ' 行ループカウンタ
    Dim thisLine As String  ' 改行コード除去後の1行
    Dim ttsText  As String  ' T.T.S.列の文字列（クォート除去後）
    ' ────────────────────────────────────────────────────

    ExtractTtsRate = Null

    lines = Split(csvText, vbLf)

    For i = LBound(lines) To UBound(lines)
        thisLine = Replace(lines(i), vbCr, "")

        ' 通貨名列は "USD (ドル)" の形式のため、コード＋" (" で一致判定する
        If InStr(thisLine, currencyCode & " (") > 0 Then
            fields = Split(thisLine, ",")
            If UBound(fields) >= 2 Then
                ttsText = Trim(Replace(fields(2), """", ""))
                If IsNumeric(ttsText) Then
                    ExtractTtsRate = CDbl(ttsText)
                End If
            End If
            Exit For
        End If
    Next i
End Function

' ---------------------------------------------------------------
' 取得したTTSレートを「為替レート」シートに書き込む
' シートが存在しない場合は新規作成する
'   usdTts : USD の T.T.S.
'   eurTts : EUR の T.T.S.
' ---------------------------------------------------------------
Private Sub WriteRatesToSheet(ByVal usdTts As Double, ByVal eurTts As Double)
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim ws As Worksheet  ' 出力先シート
    ' ────────────────────────────────────────────────────

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(OUTPUT_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = OUTPUT_SHEET_NAME

        ws.Range("A1").Value = "通貨"
        ws.Range("B1").Value = "TTS（電信売相場）"
        ws.Range("C1").Value = "取得日時"
        ws.Range("A1:C1").Font.Bold = True
    End If

    ws.Range("A2").Value = "USD"
    ws.Range("B2").Value = usdTts
    ws.Range("C2").Value = Now

    ws.Range("A3").Value = "EUR"
    ws.Range("B3").Value = eurTts
    ws.Range("C3").Value = Now

    ws.Columns("A:C").AutoFit
End Sub
