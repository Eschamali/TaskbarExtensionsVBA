Attribute VB_Name = "Demo01_JumplistControl"
'***************************************************************************************************
'                           Windows 7以降にあるジャンプリストを制御します
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'               ■■■ VBA用にカスタマイズした専用DLL 内部関数宣言セクション ■■■
'***************************************************************************************************
'* 機能    ：C++で書かれたDLLに、 ITaskbarList3 インターフェイスのサムネイルツールバー 関連の処理を埋め込ませ、機能を拡張します
'***************************************************************************************************
Private Declare PtrSafe Sub AddJumpListTask Lib "TaskbarExtensions" (ByRef Setting As JumpListData)
Private Declare PtrSafe Sub CommitJumpList Lib "TaskbarExtensions" (ByVal TargetApplicationModelUserID As LongPtr)



'***************************************************************************************************
'                           ■■■ メンバ変数/ユーザー定義型宣言 ■■■
'***************************************************************************************************
'------------------------------------定数-------------------------------------
Private Const AppUserModelID_Excel  As String = "Microsoft.Office.EXCEL.EXE.15" 'DeskTopアプリ：Excel

'---------------------------クラスユーザー定義型------------------------------
'VBA向け構造体作成
Private SettingsData As ジャンプリスト設定値
Public Type ジャンプリスト設定値
    カテゴリ名              As String
    表示名                  As String
    実行パス                As String
    コマンド引数            As String
    アイコンパス            As String
    アイコンIndex           As Long
    説明文                  As String   'カーソルを当てると、ツールチップが出ます
End Type

'DLL向け構造体作成
Private Type JumpListData
    categoryName            As LongPtr
    taskName                As LongPtr
    FilePath                As LongPtr
    cmdArguments            As LongPtr
    IconPath                As LongPtr
    Description             As LongPtr
    IconIndex               As Long
End Type



'***************************************************************************************************
'                           ■■■ VBA→DLLへ受け渡せるように変換 ■■■
'***************************************************************************************************
Private Function ConversionArgument(Target As ジャンプリスト設定値) As JumpListData
    With ConversionArgument
        .categoryName = StrPtr(Target.カテゴリ名)
        .taskName = StrPtr(Target.表示名)
        .FilePath = StrPtr(Target.実行パス)
        .cmdArguments = StrPtr(Target.コマンド引数)
        .IconPath = StrPtr(Target.アイコンパス)
        .Description = StrPtr(Target.説明文)
        .IconIndex = Target.アイコンIndex
    End With
End Function



'***************************************************************************************************
'                           ■■■ ジャンプリスト制御関数 ■■■
'***************************************************************************************************
'* 機能    ：カスタマイズしたショートカットコマンド付きジャンプリストの設定情報を登録します
'---------------------------------------------------------------------------------------------------
'* 引数　  ：※省略
'---------------------------------------------------------------------------------------------------
'* 詳細説明：DLLファイル内の、グローバル変数に保持していきます。
'***************************************************************************************************
Sub Registration(ByVal 表示名 As String, ByVal 実行パス As String, Optional ByVal コマンド引数 As String, Optional ByVal カテゴリ名 As String, Optional ByVal 説明文 As String, Optional ByVal アイコンパス As String, Optional ByVal アイコンIndex As Long)
    'アイコン未指定時は、Excelのアイコンセットをパスにする
    If アイコンパス = "" Then アイコンパス = Application.Path & "\XLICONS.EXE"

    '構造体に設定
    With SettingsData
        .カテゴリ名 = カテゴリ名
        .表示名 = 表示名
        .実行パス = 実行パス
        .コマンド引数 = コマンド引数
        .アイコンパス = アイコンパス
        .アイコンIndex = アイコンIndex
        .説明文 = 説明文
    End With
    
    'DLL内関数を実行
    AddJumpListTask ConversionArgument(SettingsData)
End Sub

'***************************************************************************************************
'* 機能    ：DLL内関数：AddJumpListTask で登録した設定情報を基に、ジャンプリストを反映します
'---------------------------------------------------------------------------------------------------
'* 引数　  ：TargetApplicationModelUserID   アプリ固有にあるID
'            PowerShellで、「Get-StartApps -Name "XXX"」と実行することで調べることが可能です。
'---------------------------------------------------------------------------------------------------
'* 注意事項：・使用後は、設定情報をクリアします。
'            ・設定情報がない状態でこれを呼び出すと、ジャンプリストをクリアします
'***************************************************************************************************
Sub Import(Optional TargetApplicationModelUserID As String = AppUserModelID_Excel)
    CommitJumpList StrPtr(TargetApplicationModelUserID)
End Sub
