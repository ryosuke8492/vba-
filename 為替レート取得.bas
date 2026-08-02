Option Explicit

'==============================================================
' 為替レート取得 VBA サンプル
' 米ドル(USD)・ユーロ(EUR)の T.T.S.（電信売相場）を取得しシートへ書き込む
'
' データ取得元（優先順）：
'   ① 三菱UFJ銀行 公開CSV（Shift-JIS、本日時点のレート）
'      https://www.bk.mufg.jp/gdocs/kinri/list_j/kinri/spot_rate.csv
'      CSV列構成：
'        通貨コード, 通貨名, T.T.S., ACC., CASH S., T.T.B., A/S, D/P・D/A, CASH B.
'        例）"001","USD (ドル)","164.12","164.54","165.92","162.12","161.70","161.40","160.12"
'   ② ①が失敗した場合のフォールバック：MURC（三菱UFJリサーチ&コンサルティング）
'      過去レートページ https://www.murc-kawasesouba.jp/fx/past/index.php?id=YYMMDD
'      をHTMLスクレイピング（正規表現でタグ除去→USD/EURの数値を抽出）。
'      土日祝でデータが無い日は最大 MURC_MAX_LOOKBACK_DAYS 日分さかのぼる。
'
' 注意：参照設定は不要（CreateObject による遅延バインディングのみ使用）
'==============================================================

Private Const MUFG_CSV_URL As String = "https://www.bk.mufg.jp/gdocs/kinri/list_j/kinri/spot_rate.csv"
Private Const MURC_PAST_RATE_URL As String = "https://www.murc-kawasesouba.jp/fx/past/index.php?id="
Private Const MURC_MAX_LOOKBACK_DAYS As Long = 5  ' フォールバック時、最大何日前まで遡るか
Private Const OUTPUT_SHEET_NAME As String = "為替レート"
Private Const HTTP_TIMEOUT_MS As Long = 15000  ' 通信タイムアウト（ミリ秒）。回線が重いときに無限待機させない
Private Const HTTP_MAX_RETRY As Long = 3       ' 通信失敗時の最大試行回数（Teams通話中などの瞬断対策）

' Sleep API（リトライ前の待機用）。Win32/Win64どちらでも動くようVBA7判定を入れる
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ---------------------------------------------------------------
' メイン処理：ボタンから呼び出す
'   ① MUFG CSV（本日レート）を試す
'   ② 失敗したら MURC 過去レートのHTMLスクレイピングにフォールバック
'   のいずれかでUSD/EURのTTSを取得し、シートに書き込む
' ---------------------------------------------------------------
Public Sub GetMufgFxRates()
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim usdTts     As Double  ' USD の T.T.S.
    Dim eurTts     As Double  ' EUR の T.T.S.
    Dim rateDate   As Date    ' レートの対象日（MUFG成功時は本日、MURC成功時は遡った日）
    Dim sourceName As String  ' 取得元（シート記録・メッセージ表示用）
    Dim gotRate    As Boolean ' いずれかの方法で取得に成功したか
    ' ────────────────────────────────────────────────────

    On Error GoTo ErrHandler

    Application.StatusBar = "為替レート取得中...（MUFG CSV）"
    gotRate = TryGetRatesFromMufgCsv(usdTts, eurTts)

    If gotRate Then
        rateDate = Date
        sourceName = "MUFG CSV（本日）"
    Else
        ' ① が失敗 → ② MURC 過去レートのスクレイピングにフォールバック
        Application.StatusBar = "MUFG CSV取得失敗。過去レート(MURC)にフォールバック中..."
        gotRate = TryGetRatesFromMurc(usdTts, eurTts, rateDate)
        sourceName = "MURC過去レート（" & Format(rateDate, "yyyy/mm/dd") & "）"
    End If

    If Not gotRate Then
        MsgBox "USD/EURのレートを取得できませんでした。" & vbCrLf & _
               "MUFG CSV・MURC過去レートのどちらからも取得に失敗しています。" & vbCrLf & _
               "ネットワーク環境（プロキシ/ファイアウォール等）をご確認ください。", vbExclamation
        GoTo CleanExit
    End If

    Call WriteRatesToSheet(usdTts, eurTts, rateDate, sourceName)

    MsgBox "為替レートを取得しました。" & vbCrLf & _
           "取得元：" & sourceName & vbCrLf & _
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
' ① MUFG CSV からのレート取得を試みる（失敗してもエラーを発生させない）
'   usdTts, eurTts : 取得成功時に値がセットされる（ByRef）
'   戻り値：成功なら True、失敗なら False
' ---------------------------------------------------------------
Private Function TryGetRatesFromMufgCsv(ByRef usdTts As Double, ByRef eurTts As Double) As Boolean
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim csvText As String   ' CSV全文（Shift-JIS→変換済み）
    Dim usdVal  As Variant  ' USD の T.T.S.（取得失敗時 Null）
    Dim eurVal  As Variant  ' EUR の T.T.S.（取得失敗時 Null）
    ' ────────────────────────────────────────────────────

    TryGetRatesFromMufgCsv = False

    On Error GoTo Failed  ' DownloadCsvAsShiftJisは最大HTTP_MAX_RETRY回試行後に失敗するとErr.Raiseする

    csvText = DownloadCsvAsShiftJis(MUFG_CSV_URL)

    usdVal = ExtractTtsRate(csvText, "USD")
    eurVal = ExtractTtsRate(csvText, "EUR")

    If IsNull(usdVal) Or IsNull(eurVal) Then GoTo Failed  ' サイト構成変更等で見つからない

    usdTts = CDbl(usdVal)
    eurTts = CDbl(eurVal)
    TryGetRatesFromMufgCsv = True
    Exit Function

Failed:
    TryGetRatesFromMufgCsv = False
End Function

' ---------------------------------------------------------------
' ② MURC（三菱UFJリサーチ&コンサルティング）過去レートページからの
'   レート取得を試みる（HTMLスクレイピング版・フォールバック用）
'   土日祝でデータが無い日は MURC_MAX_LOOKBACK_DAYS 日分さかのぼる。
'
'   usdTts, eurTts : 取得成功時に値がセットされる（ByRef）
'   rateDate       : 取得成功時、実際に値が取れた対象日がセットされる（ByRef）
'   戻り値：成功なら True、失敗なら False
' ---------------------------------------------------------------
Private Function TryGetRatesFromMurc(ByRef usdTts As Double, ByRef eurTts As Double, _
                                      ByRef rateDate As Date) As Boolean
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim http       As Object  ' MSXML2.ServerXMLHTTP（HTTP取得用）
    Dim stream     As Object  ' ADODB.Stream（文字コード変換用）
    Dim re         As Object  ' 正規表現オブジェクト（VBScript.RegExp）
    Dim html       As String  ' 取得したページのHTML全文
    Dim plainText  As String  ' HTMLタグ除去後のプレーンテキスト
    Dim idStr      As String  ' URLパラメータ用の日付文字列（YYMMDD）
    Dim url        As String  ' リクエスト先URL
    Dim usdText    As String  ' 正規表現で拾ったUSDレート文字列
    Dim eurText    As String  ' 正規表現で拾ったEURレート文字列
    Dim i          As Long    ' 遡り日数カウンタ
    ' ────────────────────────────────────────────────────

    TryGetRatesFromMurc = False
    rateDate = Date - 1  ' 当日分はまだ確定していないため1日前から探す

    On Error GoTo Failed

    Set re = CreateObject("VBScript.RegExp")

    For i = 1 To MURC_MAX_LOOKBACK_DAYS
        idStr = Format(rateDate, "yy") & Format(rateDate, "mm") & Format(rateDate, "dd")
        url = MURC_PAST_RATE_URL & idStr

        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.setTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS
        http.Open "GET", url, False
        http.setRequestHeader "User-Agent", "Mozilla/5.0"
        http.Send

        If http.Status = 200 Then
            Set stream = CreateObject("ADODB.Stream")
            stream.Type = 1               ' adTypeBinary
            stream.Open
            stream.Write http.responseBody
            stream.Position = 0
            stream.Type = 2               ' adTypeText
            stream.Charset = "Shift-JIS"
            html = stream.ReadText
            stream.Close
            Set stream = Nothing

            ' HTMLタグを空白に置換してプレーンテキスト化してから数値を検索する
            re.Global = True
            re.Pattern = "<[^>]+>"
            plainText = re.Replace(html, " ")

            re.Global = False
            re.Pattern = "USD[^\d]*([\d.]+)"
            If re.Test(plainText) Then
                usdText = re.Execute(plainText)(0).SubMatches(0)

                re.Pattern = "EUR[^\d]*([\d.]+)"
                If re.Test(plainText) Then
                    eurText = re.Execute(plainText)(0).SubMatches(0)

                    usdTts = CDbl(usdText)
                    eurTts = CDbl(eurText)
                    Set http = Nothing
                    TryGetRatesFromMurc = True
                    Exit Function
                End If
            End If
        End If

        Set http = Nothing
        rateDate = rateDate - 1  ' データなし（土日祝等）→ もう1日前へ
    Next i

Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    Set http = Nothing
    On Error GoTo 0
    TryGetRatesFromMurc = False
End Function

' ---------------------------------------------------------------
' 指定URLのCSVをバイナリ取得し、Shift-JISとしてテキスト変換する
'   url : 取得先CSVのURL
'   戻り値：変換後のCSV全文
'
' 備考：
'   MSXML2.XMLHTTPはタイムアウトを明示指定できず、回線が不安定な時
'   （Teams通話中の帯域圧迫や瞬断など）に応答が返らず固まりやすい。
'   さらにデバッグ中に強制停止すると通信中のオブジェクトが解放され
'   ないまま残り、次回実行時にも失敗しやすくなる。
'   → WinHttp.WinHttpRequest.5.1（独立したHTTPスタック）に変更し、
'     ①明示的タイムアウト ②失敗時の自動リトライ ③オブジェクトの
'     確実な解放（Set ... = Nothing）を行うことで両方に対処する。
' ---------------------------------------------------------------
Private Function DownloadCsvAsShiftJis(ByVal url As String) As String
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim http      As Object  ' WinHttp.WinHttpRequest（HTTP取得用）
    Dim stream    As Object  ' ADODB.Stream（文字コード変換用）
    Dim attempt   As Long    ' リトライ回数カウンタ
    Dim succeeded As Boolean ' 取得成功フラグ
    Dim lastErr   As String  ' 直近の失敗時エラーメッセージ
    ' ────────────────────────────────────────────────────

    succeeded = False

    For attempt = 1 To HTTP_MAX_RETRY
        On Error GoTo AttemptFailed

        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        ' 名前解決/接続/送信/受信の各タイムアウト（ミリ秒）を明示指定
        http.SetTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS
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

        succeeded = True
        Exit For

AttemptFailed:
        lastErr = Err.Description
        On Error Resume Next
        If Not stream Is Nothing Then stream.Close
        Set stream = Nothing
        Set http = Nothing
        On Error GoTo 0

        If attempt < HTTP_MAX_RETRY Then
            Sleep 1000 * attempt  ' 少し間隔をあけてリトライ（瞬断からの回復待ち）
        End If
    Next attempt

    Set stream = Nothing
    Set http = Nothing

    If Not succeeded Then
        Err.Raise vbObjectError + 2, "DownloadCsvAsShiftJis", _
                  HTTP_MAX_RETRY & "回試行しましたが通信に失敗しました：" & lastErr
    End If
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
'   usdTts     : USD の T.T.S.
'   eurTts     : EUR の T.T.S.
'   rateDate   : レートの対象日（MUFG＝本日／MURC＝遡った日）
'   sourceName : 取得元の表示名（"MUFG CSV（本日）" 等）
' ---------------------------------------------------------------
Private Sub WriteRatesToSheet(ByVal usdTts As Double, ByVal eurTts As Double, _
                               ByVal rateDate As Date, ByVal sourceName As String)
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim ws As Worksheet  ' 出力先シート
    ' ────────────────────────────────────────────────────

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(OUTPUT_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = OUTPUT_SHEET_NAME
    End If

    ' ヘッダーは毎回書き直す（旧バージョンで作られた3列シートでも列が揃うように）
    ws.Range("A1").Value = "通貨"
    ws.Range("B1").Value = "TTS（電信売相場）"
    ws.Range("C1").Value = "対象日"
    ws.Range("D1").Value = "取得元"
    ws.Range("E1").Value = "取得日時"
    ws.Range("A1:E1").Font.Bold = True

    ws.Range("A2").Value = "USD"
    ws.Range("B2").Value = usdTts
    ws.Range("C2").Value = rateDate
    ws.Range("D2").Value = sourceName
    ws.Range("E2").Value = Now

    ws.Range("A3").Value = "EUR"
    ws.Range("B3").Value = eurTts
    ws.Range("C3").Value = rateDate
    ws.Range("D3").Value = sourceName
    ws.Range("E3").Value = Now

    ws.Columns("A:E").AutoFit
End Sub
