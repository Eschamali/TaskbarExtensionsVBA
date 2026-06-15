Attribute VB_Name = "Mod01_BadgeUpdateManager"
'***************************************************************************************************
'                     オーバーレイアイコン・バッジ通知に関するモジュールです
'                              (VBA単体版 ? ITaskbarList3.cls を使用)
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ メンバ変数 ■■■
'***************************************************************************************************
Private m_taskbar As ITaskbarList3



'***************************************************************************************************
'                             ■■■ 動作に必要な定数定義 ■■■
'***************************************************************************************************
Private Const ActionPS As String = "powershell -Command """
Private Const ReadXml As String = "$XmlDocument = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]::New();$XmlDocument.loadXml($xml)"

Private Const CreateObject__Windows_UI_Notifications__BadgeNotification As String = "$badgeNotification = [Windows.UI.Notifications.BadgeNotification,Windows.UI.Notifications, ContentType = WindowsRuntime]::New($XmlDocument)"
Private Const Run__Windows_UI_Notifications__CreateBadgeUpdaterForApplication As String = "[Windows.UI.Notifications.BadgeUpdateManager,Windows.UI.Notifications, ContentType = WindowsRuntime]::createBadgeUpdaterForApplication($AppId).update($badgeNotification)"



'***************************************************************************************************
'                          ■■■ Badgeを構成するパラメーター ■■■
'***************************************************************************************************
Public Enum BadgeValueID
    bv_none = -13
    bv_attention
    bv_error
    bv_unavailable
    bv_playing
    bv_paused
    bv_newMessage
    bv_busy
    bv_away
    bv_available
    bv_alarm
    bv_alert
    bv_activity
End Enum



'***************************************************************************************************
'                               ■■■ 内部ヘルパー ■■■
'***************************************************************************************************
Private Function Taskbar() As ITaskbarList3
    If m_taskbar Is Nothing Then Set m_taskbar = New ITaskbarList3
    Set Taskbar = m_taskbar
End Function

Private Function SetFormatBadgesNotification_Xml(BadgeID As Long) As String
    Dim badgeValue As String

    Select Case BadgeID
        Case Is >= 0
            badgeValue = BadgeID
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

    Dim XmlContents As String: XmlContents = "<badge value=""" & badgeValue & """/>"
    SetFormatBadgesNotification_Xml = "$xml = '" & Replace(XmlContents, Chr(34), "\""") & Chr(39)
End Function



'***************************************************************************************************
'                      ■■■ Badge表示させるコマンド文字列を返すメソッド ■■■
'***************************************************************************************************
Function BadgeUpdaterCmd(BadgeID As BadgeValueID, Optional AppID As String = "Microsoft.Office.Excel_8wekyb3d8bbwe!microsoft.excel") As String
    BadgeUpdaterCmd = ActionPS & WorksheetFunction.TextJoin(";", False, _
        SetFormatBadgesNotification_Xml(BadgeID), ReadXml, _
        CreateObject__Windows_UI_Notifications__BadgeNotification, _
        "$AppId = '" & AppID & Chr(39), _
        Run__Windows_UI_Notifications__CreateBadgeUpdaterForApplication) & Chr(34)
End Function



'***************************************************************************************************
'                          ■■■ UWP向けバッジ通知 (PowerShell経由) ■■■
'***************************************************************************************************
Sub BadgeUpdaterDLL(BadgeID As BadgeValueID, Optional AppID As String = "Microsoft.Office.Excel_8wekyb3d8bbwe!microsoft.excel")
    Shell BadgeUpdaterCmd(BadgeID, AppID), vbHide
End Sub



'***************************************************************************************************
'                          ■■■ オーバーレイアイコン ■■■
'***************************************************************************************************
Sub UpdateTaskbarOverlayIcon(IconPath As String, Optional IconIndex As Long = 0, Optional Description As String = vbNullString, Optional hwnd As LongPtr)
    If hwnd = 0 Then hwnd = Application.hwnd
    Taskbar.UpdateOverlayIcon IconPath, IconIndex, Description, hwnd
End Sub
