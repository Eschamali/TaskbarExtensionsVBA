Attribute VB_Name = "Mod01_BadgeUpdateManager"
'***************************************************************************************************
'               VBAから、Windows PowerShell を経由して、バッジ通知を表示させます
'公式ドキュメント：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/badges
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'               ■■■ VBA用にカスタマイズした専用DLL 内部関数宣言セクション ■■■
'***************************************************************************************************
' 機能     ：C++で書かれたDLLに、 BadgeNotification関連の処理を埋め込ませ、Shell経由より高速に処理できます
'***************************************************************************************************
Private Declare PtrSafe Sub SetTaskbarOverlayBadge Lib "TaskbarProgress.dll" (ByVal badgeValue As Long, ByVal appId As LongPtr)
Private Declare PtrSafe Sub SetTaskbarOverlayBadgeForWin32 Lib "TaskbarProgress.dll" (ByVal badgeValue As Long, ByVal hwnd As LongPtr)



'***************************************************************************************************
'                             ■■■ 動作に必要な定数定義 ■■■
'***************************************************************************************************
'* 機能：PowerShell経由で実行する際の決まった文字列
'***************************************************************************************************
Private Const ActionPS As String = "powershell -Command """                                                                                                                                 'PowerShell起動コマンド
Private Const ReadXml As String = "$XmlDocument = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]::New();$XmlDocument.loadXml($xml)"     'xmlコンテンツを制御するオブジェクトを定義し、xml内容を読み込むコマンド文字列

'***************************************************************************************************
'* 機能：[Windows.UI.Notifications]に関する宣言
'***************************************************************************************************
'　xmlコンテンツから、BadgeNotificationの構造を決めます
'　→https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.badgenotification
Private Const CreateObject__Windows_UI_Notifications__BadgeNotification As String = "$badgeNotification = [Windows.UI.Notifications.BadgeNotification,Windows.UI.Notifications, ContentType = WindowsRuntime]::New($XmlDocument)"

'　Badge通知を実行するコマンド文字列
'　→https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.badgeupdatemanager.createbadgeupdaterforapplication
Private Const Run__Windows_UI_Notifications__CreateBadgeUpdaterForApplication As String = "[Windows.UI.Notifications.BadgeUpdateManager,Windows.UI.Notifications, ContentType = WindowsRuntime]::createBadgeUpdaterForApplication($AppId).update($badgeNotification)"
'***************************************************************************************************



'***************************************************************************************************
'                          ■■■ Badgeを構成するxmlコンテンツ生成 ■■■
'***************************************************************************************************
'* 機能     ：コマンドプロンプト(shell関数など)で認識できるように、xmlコンテンツを整形し、それをセットするコマンド文字列を生成します
'---------------------------------------------------------------------------------------------------
'* 返り値　 ：通知バッチスキーマのxmlContentsが返る
'* 引数　 　：BadgeID    バッジID(便宜上、数値で扱います。)
'---------------------------------------------------------------------------------------------------
'* 機能説明 ：要素が1つのみとシンプルのため、引数から直書きでxmlを作成します
'* 注意事項 ：「Microsoft XML. v6.0」の参照設定が必要です
'* URL      ：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/badgeschema/schema-root
'***************************************************************************************************
Private Function SetFormatBadgesNotification_Xml(BadgeID As Long) As String
    'Badge要素のValue属性値
    Dim badgeValue As String
    
    '引数に応じて、Value属性値を決定
    '※システムが提供するバッジ イメージだけを使うことができます。
    Select Case BadgeID
        Case Is >= 0
            badgeValue = BadgeID        '※99 を超す値の場合は 99+ と表示します。値 0 の場合はバッジを消去します。
        Case -1
            badgeValue = "activity"
        Case -2
            badgeValue = "alert"
        Case -3
            badgeValue = "alarm"
        Case -4
            badgeValue = "available"
        Case -5
            badgeValue = "away"
        Case -6
            badgeValue = "busy"
        Case -7
            badgeValue = "newMessage"
        Case -8
            badgeValue = "paused"
        Case -9
            badgeValue = "playing"
        Case -10
            badgeValue = "unavailable"
        Case -11
            badgeValue = "error"
        Case -12
            badgeValue = "attention"
        Case Else
            badgeValue = "none"
    End Select
    
    'スキーマを直書き
    Dim XmlContents As String: XmlContents = "<badge value=""" & badgeValue & """/>"

    'コマンドプロンプトで実行するため、以下の文字列をエスケープしてます
    '　「"」→「\"」
    SetFormatBadgesNotification_Xml = "$xml = '" & Replace(XmlContents, Chr(34), "\""") & Chr(39)
End Function



'***************************************************************************************************
'                          ■■■ Badge表示させるコマンド文字列を返すメソッド ■■■
'***************************************************************************************************
'* 機能     ：引数に応じたバッジ通知を表示させるコマンド文字列を返します。
'---------------------------------------------------------------------------------------------------
'* 返り値　 ：通知バッチを更新するコマンド文字列
'* 引数　 　：BadgeID       バッジID(便宜上、数値で扱います。)
'           ：AppID         デフォルト(UWP版Excel)のAppIDから変更する場合に設定。DeskTopアプリでは効かないので注意。
'                           PowerShellで、「Get-StartApps -Name "XXX"」と実行することで調べることが可能です。
'---------------------------------------------------------------------------------------------------
'* 機能説明 ：関数を呼び出すだけの簡易仕様です
'* 注意事項 ：コマンド文字列が返るだけなので実際に実行する際は、shell関数等で実行してください。
'***************************************************************************************************
Function BadgeUpdaterCmd(BadgeID As Long, Optional appId As String = "Microsoft.Office.Excel_8wekyb3d8bbwe!microsoft.excel") As String

    '1.引数に応じた、バッジのスキーマを生成し、それを読み込む。
    '2.読み込んだxmlコンテンツから、BadgeNotificationの構造を設定
    '3.AppIDを設定
    '4.Badge表示を実行
    BadgeUpdaterCmd = ActionPS & WorksheetFunction.TextJoin(";", False, _
        SetFormatBadgesNotification_Xml(BadgeID), ReadXml, _
        CreateObject__Windows_UI_Notifications__BadgeNotification, _
        "$AppId = '" & appId & Chr(39), _
        Run__Windows_UI_Notifications__CreateBadgeUpdaterForApplication) & Chr(34)

End Function



'***************************************************************************************************
'                          ■■■ DLL内部処理で、Badge表示させる ■■■
'***************************************************************************************************
'* 機能     ：引数に応じたバッジ通知を表示させます。
'---------------------------------------------------------------------------------------------------
'* 引数　 　：BadgeID       バッジID(便宜上、数値で扱います。)
'           ：AppID         デフォルト(UWP版Excel)のAppIDから変更する場合に設定。DeskTopアプリでは効かないので注意。
'                           PowerShellで、「Get-StartApps -Name "XXX"」と実行することで調べることが可能です。
'---------------------------------------------------------------------------------------------------
'* 機能説明 ：関数を呼び出すだけの簡易仕様です。shell経由よりも高速です
'***************************************************************************************************
Sub BadgeUpdaterDLL(BadgeID As Long, Optional appId As String = "Microsoft.Office.Excel_8wekyb3d8bbwe!microsoft.excel")
    'DLL内の関数を実行
    SetTaskbarOverlayBadge BadgeID, StrPtr(appId)
End Sub

'***************************************************************************************************
'* 機能     ：引数に応じたバッジ通知を表示させます。
'---------------------------------------------------------------------------------------------------
'* 引数　 　：BadgeID       バッジID(便宜上、数値で扱います。)
'           ：hwnd          ウィンドウハンドル
'---------------------------------------------------------------------------------------------------
'* 機能説明 ：DeskTopアプリでも通知バッチが使えるようにしたものです。
'***************************************************************************************************
Sub BadgeUpdaterForWin32(BadgeID As Long, Optional hwnd As LongPtr)
    '未指定ならExcelApplicationを指定
    If hwnd = 0 Then hwnd = Application.hwnd

    'DLL内の関数を実行
    SetTaskbarOverlayBadgeForWin32 BadgeID, hwnd
End Sub
