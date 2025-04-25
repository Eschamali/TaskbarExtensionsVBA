Attribute VB_Name = "Mod04_ProgressBarTaskbar"
'***************************************************************************************************
'           Windows のタスクバーで出せるプログレスバーを ITaskbarList3 を使って反映させます
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'               ■■■ VBA用にカスタマイズした専用DLL 内部関数宣言セクション ■■■
'***************************************************************************************************
' 機能     ：C++で書かれたDLLに、 ITaskbarList3 インターフェイスのプログレスバー 関連の処理を埋め込ませ、機能を拡張します
'***************************************************************************************************
Private Declare PtrSafe Sub SetTaskbarProgress Lib "TaskbarProgress.dll" (ByVal hwnd As LongPtr, ByVal current As Long, ByVal maximum As Long, ByVal Status As Long)                    'タスク バー ボタンでホストされている進行状況バー、インジケータを表示または更新
Private Declare PtrSafe Sub SetTaskbarOverlayIcon Lib "TaskbarProgress.dll" (ByVal hwnd As LongPtr, ByVal dllPath As LongPtr, ByVal IconIndex As Long, ByVal Description As LongPtr)    'タスク バー ボタンにオーバーレイを適用。Windows.UI.Notifications.BadgeUpdateManager　のような振る舞いが可能です。



'***************************************************************************************************
'                                   ■■■ 列挙型定義 ■■■
'***************************************************************************************************
' タスク バー ボタンに表示される進行状況インジケーターの種類と状態を設定します。
' パラメーターは右記に準拠→https://learn.microsoft.com/ja-jp/windows/win32/api/shobjidl_core/nf-shobjidl_core-itaskbarlist3-setprogressstate
'***************************************************************************************************
Public Enum SetProgressState
    TBPF_NOPROGRESS     '進行状況バーを非表示にする。コマンドが完了したらこの状態を使用して進行状況の状態をクリアします。
    TBPF_INDETERMINATE  '"不確定" 状態に設定します。 これは、進行状況の値を持たないが、まだ実行中のコマンドに役立ちます。
    TBPF_NORMAL         '"既定" 状態で設定します。進行中
    TBPF_ERROR = 4      '"エラー" 状態で設定します。黄色ゲージになります
    TBPF_PAUSED = 8     '"警告" 状態で設定します。一時停止・赤色ゲージになります
End Enum



'***************************************************************************************************
'                               ■■■ 公開プロシージャ ■■■
'***************************************************************************************************
'* 機能    ：Windows Taskbar Progress Barを設定/更新します
'---------------------------------------------------------------------------------------------------
'* 引数　　：currentProgress    現在の進捗値
'            maxProgress        100%とする値
'            Status             プログレスバーの種類
'            hwnd               ウィンドウハンドル(通常は、Application.hwnd)
'***************************************************************************************************
Sub UpdateTaskbarProgress(currentProgress As Long, maxProgress As Long, Optional Status As SetProgressState = 2, Optional hwnd As LongPtr)
    'hwnd未指定なら、Excelを指定
    If hwnd = 0 Then hwnd = Application.hwnd

    ' DLL関数の呼び出し
    SetTaskbarProgress hwnd, currentProgress, maxProgress, Status
End Sub

'***************************************************************************************************
'* 機能    ：Windows Taskbar にステータスアイコンを設定/更新します
'---------------------------------------------------------------------------------------------------
'* 引数　　：iconPath           アイコンデータのあるフルパス
'            iconIndex          複数アイコンがある場合の、Index値。-1以下で、リセットします。
'            description        アクセシビリティ向け説明文
'            hwnd               ステータスアイコンを設定させるウィンドウハンドル
'***************************************************************************************************
Sub UpdateTaskbarOverlayIcon(IconPath As String, Optional IconIndex As Long = 0, Optional Description As String, Optional hwnd As LongPtr)
    'hwnd未指定なら、Excelを指定
    If hwnd = 0 Then hwnd = Application.hwnd

    ' DLL関数を呼び出し、タスクバーにオーバーレイアイコンを設定
    SetTaskbarOverlayIcon hwnd, StrPtr(IconPath), IconIndex, StrPtr(Description)

    '任意のDLLのパス例
    'iconPath = "C:\Users\XXX\Saved Games\Game.exe"
    'iconPath = "C:\Windows\System32\shell32.dll"
    'iconPath = "C:\Users\XXX\Downloads\YYY.ico"
End Sub
