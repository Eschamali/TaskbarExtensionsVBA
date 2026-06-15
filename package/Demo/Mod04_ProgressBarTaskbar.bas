Attribute VB_Name = "Mod04_ProgressBarTaskbar"
'***************************************************************************************************
'                   Windows 7 以降 のタスクバーに、プログレスバーを反映させます
'                              (VBA単体版 - ITaskbarList3.cls を使用)
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ メンバ変数 ■■■
'***************************************************************************************************
Private m_taskbar As ITaskbarList3



'***************************************************************************************************
'                               ■■■ 内部ヘルパー ■■■
'***************************************************************************************************
Private Function Taskbar() As ITaskbarList3
    If m_taskbar Is Nothing Then Set m_taskbar = New ITaskbarList3
    Set Taskbar = m_taskbar
End Function



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
Sub UpdateTaskbarProgress(currentProgress As Long, Optional maxProgress As Long = 100, Optional Status As SetProgressState = TBPF_NORMAL, Optional hwnd As LongPtr)
    If hwnd = 0 Then hwnd = Application.hwnd
    Taskbar.UpdateProgress currentProgress, maxProgress, Status, hwnd
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
Public Function SetProgressValueForWindowsTerminal(progress As Long, Optional state As SetProgressStateForTerminal = TER_NORMAL) As String
    SetProgressValueForWindowsTerminal = Taskbar.ProgressSequenceForTerminal(progress, state)
End Function
