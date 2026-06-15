Attribute VB_Name = "Demo_ThumbnailToolbar"
'***************************************************************************************************
'                       サムネイルツールバーのデモ用マクロ集です
'
'  事前に以下をインポートしてから実行してください。
'    - package\ITaskbarList3.cls
'    - package\ITaskbarSubclassHandler.cls
'***************************************************************************************************
Option Explicit



Sub Demo_ThumbnailToolbars()
    Static taskbar As New ITaskbarList3

    With taskbar
        .InitThumbBar Application.hwnd
        .ConfigureThumbButton 1, "Run01FromThumbnailToolbars", , , , "クリックしてマクロ発動"
        .UpdateThumbButton 1

        MsgBox "登録完了しました。タスクバーの Excel にカーソルをあわせて、ご確認ください。", vbInformation, "サムネイルツールバー"
    End With
End Sub


Sub Run01FromThumbnailToolbars()
    MsgBox "1つ目のボタンを押しました", vbInformation, "Pushed"
End Sub
