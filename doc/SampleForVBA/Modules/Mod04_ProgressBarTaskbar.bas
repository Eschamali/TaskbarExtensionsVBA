Attribute VB_Name = "Mod04_ProgressBarTaskbar"
'***************************************************************************************************
'                   Windows 7 以降 のタスクバーに、プログレスバーを 反映させます
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'               ■■■ VBA用にカスタマイズした専用DLL 内部関数宣言セクション ■■■
'***************************************************************************************************
' 機能     ：C++で書かれたDLLに、 ITaskbarList3 インターフェイスのプログレスバー 関連の処理を埋め込ませ、機能を拡張します
'***************************************************************************************************
Private Declare PtrSafe Sub SetTaskbarProgress Lib "TaskbarExtensions" (ByVal hwnd As LongPtr, ByVal current As Long, ByVal maximum As Long, ByVal Status As Long)                    'タスク バー ボタンでホストされている進行状況バー、インジケータを表示または更新



'***************************************************************************************************
'                                   ■■■ 列挙型定義 ■■■
'***************************************************************************************************
' タスク バー ボタンに表示される進行状況インジケーターの種類と状態を設定します。
' パラメーターは右記に準拠→https://learn.microsoft.com/ja-jp/windows/win32/api/shobjidl_core/nf-shobjidl_core-itaskbarlist3-setprogressstate
'***************************************************************************************************
'全アプリ共通
Public Enum SetProgressState
    TBPF_NOPROGRESS     '進行状況バーを非表示にする。コマンドが完了したらこの状態を使用して進行状況の状態をクリアします。
    TBPF_INDETERMINATE  '"不確定" 状態に設定します。 これは、進行状況の値を持たないが、まだ実行中のコマンドに役立ちます。
    TBPF_NORMAL         '"既定" 状態で設定します。進行中
    TBPF_ERROR = 4      '"エラー" 状態で設定します。黄色ゲージになります
    TBPF_PAUSED = 8     '"警告" 状態で設定します。一時停止・赤色ゲージになります
End Enum

'Windows Terminal限定
Public Enum SetProgressStateForTerminal
    TER_NOPROGRESS     '進行状況バーを非表示にする。コマンドが完了したらこの状態を使用して進行状況の状態をクリアします。
    TER_NORMAL         '"既定" 状態で設定します。進行中
    TER_ERROR          '"エラー" 状態で設定します。黄色ゲージになります
    TER_INDETERMINATE  '"不確定" 状態に設定します。 これは、進行状況の値を持たないが、まだ実行中のコマンドに役立ちます。
    TER_PAUSED         '"警告" 状態で設定します。一時停止・赤色ゲージになります
End Enum



'***************************************************************************************************
'                               ■■■ メインプロシージャ ■■■
'***************************************************************************************************
'* 機能    ：Windows Taskbar Progress Barを設定/更新します
'---------------------------------------------------------------------------------------------------
'* 引数　　：currentProgress    現在の進捗値
'            maxProgress        100%とする値
'            Status             プログレスバーの種類
'            hwnd               ウィンドウハンドル(通常は、Application.hwnd)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ウィンドウハンドルが取れるアプリであれば何でもOKです
'***************************************************************************************************
Sub UpdateTaskbarProgress(currentProgress As Long, Optional maxProgress As Long = 100, Optional Status As SetProgressState = 2, Optional hwnd As LongPtr)
    'hwnd未指定なら、Excelを指定
    If hwnd = 0 Then hwnd = Application.hwnd

    ' DLL関数の呼び出し
    SetTaskbarProgress hwnd, currentProgress, maxProgress, Status
End Sub

'***************************************************************************************************
'* 機能    ：Windows Terminal 専用の、Windows Taskbar Progress Barを設定します
'---------------------------------------------------------------------------------------------------
'* 返り値　：Windows Terminalに進捗状況を送信する制御文字列。
'* 引数　　：progress      0～100　の進捗値を設定
'            state         進捗状況のステータスを設定。0、1、2、3、4 のいずれかです。上部のユーザー定義型「SetProgressStateForTerminal」を参考に
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Windows Terminal内のbatファイルで扱えるConEmu "進行状況バー" シーケンス ("OSC 9;4" とも呼ばれます)の設定文字列を返します。
'* 注意事項：batファイルから、進捗状況を反映するときに使うこと。Windows Terminal v1.6 以降。
'* URL     ：https://learn.microsoft.com/ja-jp/windows/terminal/tutorials/progress-bar-sequences
'***************************************************************************************************
Public Function SetProgressValueForWindowsTerminal(progress As Long, Optional state As SetProgressStateForTerminal = 1) As String
    '0-100の範囲に収める
    Dim ProgressValue As Integer
    If progress < 0 Then
        ProgressValue = 0
    ElseIf progress > 100 Then
        ProgressValue = 100
    Else
        ProgressValue = progress
    End If

    '進行状況バー シーケンスの形式を返します。コマンド プロンプトの場合、制御文字を直接、埋め込むことで実現します
    '→https://learn.microsoft.com/en-us/windows/terminal/tutorials/progress-bar-sequences
    ForWindowsTerminal = "<NUL SET /p =" & Chr(27) & "]9;4;" & state & ";" & ProgressValue & Chr(7)
End Function
