VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsThumbButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'           Windows 7以降のサムネイルツールバー(縮小表示ツール バー)を制御します
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'               ■■■ VBA用にカスタマイズした専用DLL 内部関数宣言セクション ■■■
'***************************************************************************************************
'* 機能    ：C++で書かれたDLLに、 ITaskbarList3 インターフェイスのサムネイルツールバー 関連の処理を埋め込ませ、機能を拡張します
'***************************************************************************************************
Private Declare PtrSafe Sub InitializeThumbnailButton Lib "TaskbarProgress.dll" (ByVal NumberSecured As Long, ByVal hwnd As LongPtr)
Private Declare PtrSafe Sub UpdateThumbnailButton Lib "TaskbarProgress.dll" (ByRef TargetUpdata As THUMBBUTTONDATA, ByVal hwnd As LongPtr)



'***************************************************************************************************
'                                   ■■■ 列挙型定義 ■■■
'***************************************************************************************************
'* 機能    ：THUMBBUTTON によって、ボタンの特定の状態と動作を制御するために使用されます。
'---------------------------------------------------------------------------------------------------
'* URL     ：https://learn.microsoft.com/ja-jp/windows/win32/api/shobjidl_core/ne-shobjidl_core-thumbbuttonflags
'***************************************************************************************************
Public Enum THUMBBUTTONFLAGS
  THBF_ENABLED                  'ボタン活性状態
  THBF_DISABLED                 'ボタン非活性状態
  THBF_DISMISSONCLICK           'ボタンをクリックされると、タスク バー ボタンのポップアップがすぐに閉じます。連続押下を防止できます
  THBF_NOBACKGROUND = 4         'ボタンの境界線をなくします
  THBF_HIDDEN = 8               'ボタン自体を非表示します。UPDATEで、再度表示可能です
  THBF_NONINTERACTIVE = 16      '対話型にしません。
End Enum



'***************************************************************************************************
'                           ■■■ メンバ変数/ユーザー定義型宣言 ■■■
'***************************************************************************************************
'* 機能    ：THUMBBUTTON 関連でよく使うやつ
'***************************************************************************************************
'---------------------------クラス定数------------------------------
Private Const MAX_BUTTONS                       As Integer = 7  '配置可能なボタンの上限数
Private Const ButtonID_Correction               As Long = 1000  'ボタンIDの採番開始番号

'---------------------------クラス変数------------------------------
Private SettingTHUMBBUTTON(1 To MAX_BUTTONS)    As THUMBBUTTON  'ボタン情報格納用

'---------------------------クラスエラー番号用定数------------------------------
Private Const ErrorValue_範囲外THUMBBUTTON      As Long = 1
Private Const ErrorNAme_範囲外THUMBBUTTON       As String = "THUMBBUTTON 範囲外"

'---------------------------クラスユーザー定義型------------------------------
'VBA向け構造体作成
Private Type THUMBBUTTON
    ProcedureName   As String   '実行するプロシージャ名
    ButtonType      As Long     'ボタン制御コード
    IconPath        As String   'アイコンのパス
    IconIndex       As Long     'アイコンインデックス
    Description     As String   'ツールチップ
    ButtonIndex     As Long     'ボタン位置
End Type

'DLL向け構造体作成
Private Type THUMBBUTTONDATA
    ProcedureName   As LongPtr  '実行するプロシージャ名
    IconPath        As LongPtr  'ボタン制御コード
    Description     As LongPtr  'アイコンのパス
    ButtonType      As Long     'アイコンインデックス
    IconIndex       As Long     'ツールチップ
    ButtonIndex     As Long     'ボタン位置
End Type



'***************************************************************************************************
'                                ■■■ THUMBBUTTON 構造体を追加 ■■■
'***************************************************************************************************
'* 機能    ：ウィンドウのサムネイル表現に埋め込まれたツール バーで使用されるボタンを定義します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：ボタンを定義一覧情報が詰まった2次元配列
'* 引数　　：ButtonIndex        追加位置
'            ProcedureName      標準プロシージャにある引数なし関数名
'            ButtonType         ボタンの挙動
'            iconPath           アイコンデータのあるフルパス
'            iconIndex          複数アイコンがある場合の、Index値。
'            description        ボタンにカーソルを当てた際のツールチップ
'---------------------------------------------------------------------------------------------------
'* URL      ：https://learn.microsoft.com/ja-jp/windows/win32/api/shobjidl_core/ns-shobjidl_core-thumbbutton
'***************************************************************************************************
Public Property Let SetProperty_THUMBBUTTON(Optional ButtonIndex = 1, Optional IconPath As String, Optional IconIndex As Long, Optional ButtonType As THUMBBUTTONFLAGS, Optional Description As String, ProcedureName As String)
    '範囲外なら、エラーで、止める
    If ButtonIndex < 1 Or ButtonIndex > MAX_BUTTONS Then
        Err.Raise vbObjectError + ErrorValue_範囲外THUMBBUTTON, ErrorNAme_範囲外THUMBBUTTON, "ボタンインデックスは 1 ～ 7 にしてください。"
    Else
        '範囲内なら、設定する
        With SettingTHUMBBUTTON(ButtonIndex)
            .ButtonIndex = ButtonIndex
            .ButtonType = ButtonType
            .Description = Description
            .IconIndex = IconIndex
            .IconPath = IconPath
            .ProcedureName = ProcedureName
        End With
    End If
End Property

'***************************************************************************************************
'* 機能    ：ウィンドウのサムネイル表現に埋め込まれたツール バーで使用されるボタン定義情報をリセットします
'---------------------------------------------------------------------------------------------------
'* 引数　　：ButtonIndex        リセットしたいボタンインデックス
'---------------------------------------------------------------------------------------------------
'* 注意事項：ButtonIndex を 0(False) で渡すと、ALL リセットします
'***************************************************************************************************
Public Property Let ClearProperty_THUMBBUTTON(ButtonIndex As Long)
    '範囲外なら、エラーで、止める
    If ButtonIndex < 0 Or ButtonIndex > MAX_BUTTONS Then
        Err.Raise vbObjectError + ErrorValue_範囲外THUMBBUTTON, ErrorNAme_範囲外THUMBBUTTON, "ボタンインデックスは 1 ～ 7 にしてください。"
    ElseIf ButtonIndex = 0 Then
        '0の場合は、ALL Reset
        For ButtonIndex = 1 To MAX_BUTTONS
            SetProperty_THUMBBUTTON(ButtonIndex, vbNullString, 0, THBF_HIDDEN, vbNullString) = vbNullString
        Next
    Else
        '範囲内なら、設定する
        SetProperty_THUMBBUTTON(ButtonIndex, vbNullString, 0, THBF_HIDDEN, vbNullString) = vbNullString
    End If
End Property

'***************************************************************************************************
'* 機能    ：VBA→DLL への受け渡し用structを作成します。
'---------------------------------------------------------------------------------------------------
'* 返り値  ：ボタンを定義一覧情報が詰まった1次元配列(ユーザー定義型)
'* 引数　　：ユーザー定義型：THUMBBUTTON
'---------------------------------------------------------------------------------------------------
'* 注意事項：・DLL側で、シグネチャ（型や順序）が合うようにすること。
'          ：・複数の型を混在させる場合、散在せずまとめるようにすること
'          ：・Type 名は、DLL側の struct と同一にすること
'          ：・Type 内の変数名も、DLL側の struct 内と同一にすること
'          ：・DLL 側の型が、 BOOL であっても、こちら VBA 側は、 long 型にすること
'          ：・2次元配列は基本、渡せません。
'          ：・渡せる引数の数に制限があります。
'***************************************************************************************************
Private Function UseDLL_THUMBBUTTON(Target As THUMBBUTTON) As THUMBBUTTONDATA
    With UseDLL_THUMBBUTTON
        '文字列に関するパラメーター　LongPtrを使うことで、ポインタから文字列を取得するようにします。
        .ProcedureName = StrPtr(Target.ProcedureName)
        .IconPath = StrPtr(Target.IconPath)
        .Description = StrPtr(Target.Description)
        
        '数値
        .ButtonType = Target.ButtonType
        .IconIndex = Target.IconIndex
        .ButtonIndex = Target.ButtonIndex + ButtonID_Correction
    End With
End Function



'***************************************************************************************************
'                                      ■■■ メイン処理 ■■■
'***************************************************************************************************
'* 機能    ：指定ハンドルに、THUMBBUTTON 構造体を埋め込みます
'---------------------------------------------------------------------------------------------------
'* 引数　　：MakeButtonNum      確保するボタン数
'            hwnd               ウィンドウハンドル
'---------------------------------------------------------------------------------------------------
'* 注意事項：・非表示として確保するためこれだけでは見た目上、何も起こりません
'            ・MakeButtonNum を 0 で渡すと、定義自体を削除します
'***************************************************************************************************
Public Property Let InstallationButton(Optional hwnd As LongPtr, MakeButtonNum As Long)
    '未指定ならExcelApplicationを指定
    If hwnd = 0 Then hwnd = Application.hwnd
    
    'DLL内の関数を実行
    InitializeThumbnailButton MakeButtonNum, hwnd
End Property

'***************************************************************************************************
'* 機能    ：指定ボタンIndexに対して、変更を加えます
'---------------------------------------------------------------------------------------------------
'* 引数　　：ButtonIndex    ボタン位置
'            hwnd           ウィンドウハンドル
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前に、「SetProperty_THUMBBUTTON」で定義情報を用意しておくこと
'***************************************************************************************************
Public Property Let UpdateButton(Optional hwnd As LongPtr, ButtonIndex As Long)
    '未指定ならExcelApplicationを指定
    If hwnd = 0 Then hwnd = Application.hwnd
    
    'DLL内の関数を実行
    UpdateThumbnailButton UseDLL_THUMBBUTTON(SettingTHUMBBUTTON(ButtonIndex)), hwnd
End Property
