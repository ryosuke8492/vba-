Option Explicit

'==============================================================
' モジュール名：（ここにモジュール名）
' 概要　　　　：（このモジュールが何をするか1〜2行で）
'
' 構成（C言語ソース風）：
'   1. 列挙型・定数定義（enum / #define 相当）
'   2. モジュールレベル変数（グローバル変数相当。配列はここで宣言だけし、
'      実際の値は Initialize() でセットする）
'   3. 関数一覧（プロトタイプ宣言相当。全体像を把握するための目次）
'   4. Initialize（グローバル初期化。C の初期化処理相当）
'   5. Main（エントリポイント。C の main() 相当。ボタンから呼び出す）
'   6. 以降、個別の関数・サブルーチン（XML docコメント形式で説明を付与）
'
' コーディングルール：
'   ・引数は ByVal / ByRef を必ず明示する
'     （VBAは省略時ByRefになるため、値渡しのつもりが参照渡しになる事故を防ぐ）
'   ・関連する定数の羅列は Const ではなく Enum でグループ化する
'==============================================================

'--------------------------------------------------------------
' 列挙型定義（C言語の enum 相当）
'   Enum メンバー名はプロジェクト全体でグローバルな名前空間を共有するため、
'   他モジュールと衝突しないよう接頭辞（ここでは Tpl_）を付けて回避する。
'--------------------------------------------------------------
Public Enum ETplResult
    Tpl_ResultOK = 0  ' 正常終了
    Tpl_ResultNG = 1  ' 異常終了（値なし等）
End Enum

'--------------------------------------------------------------
' 定数定義
'--------------------------------------------------------------
Private Const CONST_SAMPLE_A As Long = 1        ' （説明）
Private Const CONST_SAMPLE_B As String = "サンプル"  ' （説明）

'--------------------------------------------------------------
' モジュールレベル変数
'   配列やオブジェクトはここで宣言のみ行い、値は Initialize() でセットする。
'   宣言時点で初期値を書かないことで「初期化のタイミング」を1箇所に集約する。
'--------------------------------------------------------------
Private mSampleArray()  As Variant  ' サンプル配列（Initialize()で初期化）
Private mIsInitialized  As Boolean  ' 初期化済みフラグ（多重初期化防止）

'--------------------------------------------------------------
' 関数一覧（この一覧を見れば全体像がわかるようにする）
'   Initialize        : モジュール変数・配列を初期化する
'   Main              : エントリポイント。ボタンから呼び出す
'   SampleFunction    : 値の妥当性を判定する（ByVal引数の例）
'   FillSampleValues  : 呼び出し元の配列に値を詰める（ByRef引数の例）
'--------------------------------------------------------------

''' <summary>
''' モジュールレベル変数・配列を初期化する。
''' 多重初期化を防ぐため、mIsInitialized で二重実行をガードする。
''' </summary>
Private Sub Initialize()
    If mIsInitialized Then Exit Sub

    mSampleArray = Array("値1", "値2", "値3")

    mIsInitialized = True
End Sub

''' <summary>
''' エントリポイント。ボタン等から呼び出す想定。
''' </summary>
Public Sub Main()
    ' ─── 変数宣言 ───────────────────────────────────────
    Dim status As ETplResult  ' SampleFunction の判定結果
    Dim values() As Variant   ' FillSampleValues で受け取る配列
    ' ────────────────────────────────────────────────────

    Call Initialize

    ' ─── ここから処理 ─────────────────────────────
    status = SampleFunction(CStr(mSampleArray(0)))
    Call FillSampleValues(values)
    ' ────────────────────────────────────────────
End Sub

''' <summary>
''' 引数の妥当性を判定する（ByVal = 値渡し。呼び出し元の変数は変更されない）
''' </summary>
''' <param name="value">ByVal: 判定対象の文字列（呼び出し元にはコピーが渡る）</param>
''' <returns>Tpl_ResultOK / Tpl_ResultNG</returns>
Private Function SampleFunction(ByVal value As String) As ETplResult
    If Len(value) > 0 Then
        SampleFunction = Tpl_ResultOK
    Else
        SampleFunction = Tpl_ResultNG
    End If
End Function

''' <summary>
''' 呼び出し元の配列に値を詰める（ByRef = 参照渡し。呼び出し元の変数が書き換わる）
''' </summary>
''' <param name="outValues">ByRef: 結果を書き込む配列（呼び出し元でDimしたものを渡す）</param>
Private Sub FillSampleValues(ByRef outValues() As Variant)
    outValues = mSampleArray
End Sub
