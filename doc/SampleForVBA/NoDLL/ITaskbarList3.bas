Attribute VB_Name = "ITaskbarList3"
'***************************************************************************************************
'            Windows 7 以降 タスクバーボタンに、ITaskbarList3関連の機能を反映させます
'       おまけとして、「Windows Terminal」用のプログレスバー定義用コマンドも生成させます
'***************************************************************************************************
Option Explicit
Option Private Module



'***************************************************************************************************
'               ■■■ VBA単体 (DispCallFunc × VTable) 内部関数宣言セクション ■■■
'***************************************************************************************************
' 機能     ：ITaskbarList3 インターフェイスを CoCreateInstance で取得し、
'            DispCallFunc 経由で VTable メソッド (SetProgressState / SetProgressValue) を呼び出します
' 参考     ：https://github.com/sancarn/stdVBA/blob/master/src/stdCOM.cls
'***************************************************************************************************
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByVal prgvt As LongPtr, _
    ByVal prgpvarg As LongPtr, _
    ByRef pvargResult As Variant) As Long

Private Declare PtrSafe Function CoCreateInstance Lib "ole32" ( _
    ByRef rclsid As GUID, _
    ByVal pUnkOuter As LongPtr, _
    ByVal dwClsContext As Long, _
    ByRef riid As GUID, _
    ByRef ppv As LongPtr) As Long

Private Declare PtrSafe Function IIDFromString Lib "ole32" ( _
    ByVal lpsz As LongPtr, _
    ByRef riid As GUID) As Long



'***************************************************************************************************
'                                   ■■■ 定数 / 型定義 ■■■
'***************************************************************************************************
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Const CC_STDCALL As Long = 4
Private Const CLSCTX_INPROC_SERVER As Long = 1
Private Const S_OK As Long = 0

Private Const CLSID_TaskbarList As String = "{56FDF344-FD6D-11d0-958A-006097C9A090}"
Private Const IID_ITaskbarList3 As String = "{EA1AFBA6-0097-4908-9580-150FCC282DC6}"
Private Const IID_ITaskbarList3_Alt As String = "{ea1afb91-9e28-4b86-90e9-9e9f8a5eefaf}"

#If Win64 Then
    Private Const PTRSIZE As Long = 8
    Private Const VT_PARAM_PTR As Integer = vbLongLong
#Else
    Private Const PTRSIZE As Long = 4
    Private Const VT_PARAM_PTR As Integer = vbLong
#End If

Private Const VTBL_HRINIT As Long = 3
Private Const VTBL_RELEASE As Long = 2

'ITaskbarList3 (IUnknown + ITaskbarList + ITaskbarList2 を継承した VTable インデックス)
Private Const VTBL_SETPROGRESSVALUE As Long = 9
Private Const VTBL_SETPROGRESSSTATE As Long = 10



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
'                           ■■■ ITaskbarList3  各純メソッドヘルパー ■■■
'***************************************************************************************************
' 機能     ：ITaskbarList3::SetProgressState
'---------------------------------------------------------------------------------------------------
' URL      ：https://learn.microsoft.com/ja-jp/windows/desktop/api/shobjidl_core/nf-shobjidl_core-itaskbarlist3-setprogressstate
'***************************************************************************************************
Private Function SetProgressState(ByVal pTaskbarList As LongPtr, ByVal hwnd As LongPtr, ByVal state As Long) As Long
    Dim vArgs(0 To 1) As Variant
    Dim vTypes(0 To 1) As Integer
    Dim vPtrs(0 To 1) As LongPtr

    vArgs(0) = hwnd
    vArgs(1) = state
    vTypes(0) = VT_PARAM_PTR
    vTypes(1) = vbLong
    vPtrs(0) = VarPtr(vArgs(0))
    vPtrs(1) = VarPtr(vArgs(1))

    SetProgressState = InvokeComMethod(pTaskbarList, VTBL_SETPROGRESSSTATE, 2, vTypes, vPtrs)
End Function

'***************************************************************************************************
' 機能     ：ITaskbarList3::SetProgressValue
'---------------------------------------------------------------------------------------------------
' URL      ：https://learn.microsoft.com/ja-jp/windows/desktop/api/shobjidl_core/nf-shobjidl_core-itaskbarlist3-setprogressvalue
'***************************************************************************************************
Private Function SetProgressValue(ByVal pTaskbarList As LongPtr, ByVal hwnd As LongPtr, ByVal current As Long, ByVal maximum As Long) As Long
    Dim vArgs(0 To 2) As Variant
    Dim vTypes(0 To 2) As Integer
    Dim vPtrs(0 To 2) As LongPtr

    vArgs(0) = hwnd
    vArgs(1) = current
    vArgs(2) = maximum
    vTypes(0) = VT_PARAM_PTR
    vTypes(1) = VT_PARAM_PTR
    vTypes(2) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    vPtrs(1) = VarPtr(vArgs(1))
    vPtrs(2) = VarPtr(vArgs(2))

    SetProgressValue = InvokeComMethod(pTaskbarList, VTBL_SETPROGRESSVALUE, 3, vTypes, vPtrs)
End Function



'***************************************************************************************************
'                    ■■■ VBAから操作しやすいように改良した各種ヘルパー ■■■
'***************************************************************************************************
' 機能     ：SetProgressState/SetProgressValue を統合して、1つのプロシージャで呼び出すようにします
'***************************************************************************************************
Private Sub SetTaskbarProgress(ByVal hwnd As LongPtr, ByVal current As Long, ByVal maximum As Long, ByVal Status As Long)
    Dim pTaskbarList As LongPtr

    pTaskbarList = CreateITaskbarList3()
    If pTaskbarList = 0 Then Exit Sub

    Call TaskbarHrInit(pTaskbarList)
    Call ITaskbarList3.SetProgressState(pTaskbarList, hwnd, Status)

    If Status = TBPF_NORMAL Or Status = TBPF_PAUSED Or Status = TBPF_ERROR Then
        Call ITaskbarList3.SetProgressValue(pTaskbarList, hwnd, current, maximum)
    End If

    Call ComRelease(pTaskbarList)
End Sub



'***************************************************************************************************
'                               ■■■ COM / VTable ヘルパー ■■■
'***************************************************************************************************
' 機能     ：ITaskbarList3 インターフェイスを操作するためのヘルパー群です。
'***************************************************************************************************
Private Function CreateITaskbarList3() As LongPtr
    Dim clsid As GUID
    Dim iid As GUID
    Dim pTaskbarList As LongPtr
    Dim iidList(0 To 1) As String
    Dim i As Long

    If IIDFromString(StrPtr(CLSID_TaskbarList), clsid) <> S_OK Then Exit Function

    iidList(0) = IID_ITaskbarList3
    iidList(1) = IID_ITaskbarList3_Alt

    For i = LBound(iidList) To UBound(iidList)
        If IIDFromString(StrPtr(iidList(i)), iid) = S_OK Then
            If CoCreateInstance(clsid, 0, CLSCTX_INPROC_SERVER, iid, pTaskbarList) = S_OK Then
                CreateITaskbarList3 = pTaskbarList
                Exit Function
            End If
        End If
    Next i
End Function

Private Function InvokeComMethod(ByVal pInterface As LongPtr, ByVal vTableIndex As Long, ByVal paramCount As Long, ByRef vTypes() As Integer, ByRef vPtrs() As LongPtr) As Long
    Dim vResult As Variant
    Dim dispResult As Long

    If paramCount = 0 Then
        dispResult = DispCallFunc(pInterface, vTableIndex * PTRSIZE, CC_STDCALL, vbLong, 0, 0, 0, vResult)
    Else
        dispResult = DispCallFunc(pInterface, vTableIndex * PTRSIZE, CC_STDCALL, vbLong, paramCount, VarPtr(vTypes(0)), VarPtr(vPtrs(0)), vResult)
    End If

    If dispResult <> S_OK Then
        InvokeComMethod = dispResult
    ElseIf IsEmpty(vResult) Then
        InvokeComMethod = S_OK
    Else
        InvokeComMethod = CLng(vResult)
    End If
End Function

Private Function TaskbarHrInit(ByVal pTaskbarList As LongPtr) As Long
    Dim vTypes(0 To 0) As Integer
    Dim vPtrs(0 To 0) As LongPtr

    TaskbarHrInit = InvokeComMethod(pTaskbarList, VTBL_HRINIT, 0, vTypes, vPtrs)
End Function

Private Sub ComRelease(ByVal pInterface As LongPtr)
    Dim vResult As Variant

    If pInterface = 0 Then Exit Sub
    Call DispCallFunc(pInterface, VTBL_RELEASE * PTRSIZE, CC_STDCALL, vbLong, 0, 0, 0, vResult)
End Sub

'***************************************************************************************************
'* 機能    ：タスクバー API 呼び出しの診断（HRESULT を表示）
'* 使い方  ：うまく表示されない場合に 1 回実行して、結果を確認してください
'***************************************************************************************************
Public Sub DiagnoseTaskbarProgress()
    Dim pTaskbarList As LongPtr
    Dim hr As Long
    Dim msg As String

    pTaskbarList = CreateITaskbarList3()
    msg = "hwnd=" & Hex(Application.hwnd) & vbCrLf
    msg = msg & "CreateITaskbarList3=" & Hex(pTaskbarList) & vbCrLf

    If pTaskbarList = 0 Then
        MsgBox msg & "CoCreateInstance に失敗しました。", vbExclamation, "Taskbar Progress 診断"
        Exit Sub
    End If

    hr = TaskbarHrInit(pTaskbarList)
    msg = msg & "HrInit=0x" & Hex(hr) & vbCrLf

    hr = SetProgressState(pTaskbarList, Application.hwnd, TBPF_INDETERMINATE)
    msg = msg & "SetProgressState(INDETERMINATE)=0x" & Hex(hr) & vbCrLf

    hr = SetProgressValue(pTaskbarList, Application.hwnd, 50, 100)
    msg = msg & "SetProgressValue(50,100)=0x" & Hex(hr) & vbCrLf

    Call ComRelease(pTaskbarList)
    MsgBox msg, vbInformation, "Taskbar Progress 診断"
End Sub



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
Public Sub UpdateTaskbarProgress(currentProgress As Long, Optional maxProgress As Long = 100, Optional Status As SetProgressState = TBPF_NORMAL, Optional hwnd As LongPtr)
    'hwnd未指定なら、Excelを指定
    If hwnd = 0 Then hwnd = Application.hwnd

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
    SetProgressValueForWindowsTerminal = "<NUL SET /p =" & Chr(27) & "]9;4;" & state & ";" & ProgressValue & Chr(7)
End Function



'***************************************************************************************************
'                               ■■■ Demo用プロシージャ ■■■
'***************************************************************************************************
' 機能     ：SetProgressState/SetProgressValue Demo
'***************************************************************************************************
Sub demo_UpdateTaskbarProgress()
    UpdateTaskbarProgress 50
End Sub
