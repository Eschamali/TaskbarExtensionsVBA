VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITaskbarList3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'            Windows 7 以降 タスクバーボタンに、ITaskbarList3関連の機能を反映させます
'       おまけとして、「Windows Terminal」用のプログレスバー定義用コマンドも生成させます
'***************************************************************************************************
Option Explicit

'***************************************************************************************************
'               ■■■ VBA単体 (DispCallFunc × VTable) 内部関数宣言セクション ■■■
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

Private Declare PtrSafe Function ExtractIconExW Lib "shell32.dll" ( _
    ByVal lpszFile As LongPtr, _
    ByVal nIconIndex As Long, _
    ByRef phiconLarge As LongPtr, _
    ByRef phiconSmall As LongPtr, _
    ByVal nIcons As Long) As Long

Private Declare PtrSafe Function DestroyIcon Lib "user32.dll" ( _
    ByVal hIcon As LongPtr) As Long

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
Private Const STATE_UNSET As Long = -1

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
Private Const VTBL_THUMBBARADDBUTTONS As Long = 15
Private Const VTBL_THUMBBARUPDATEBUTTONS As Long = 16
Private Const VTBL_SETOVERLAYICON As Long = 18

Private Const MAX_THUMB_BUTTONS As Long = 7

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

Public Enum THUMBBUTTONFLAGS
    THBF_ENABLED = 0
    THBF_DISABLED = 1
    THBF_DISMISSONCLICK = 2
    THBF_NOBACKGROUND = 4
    THBF_HIDDEN = 8
    THBF_NONINTERACTIVE = 16
End Enum



'***************************************************************************************************
'                                   ■■■ メンバ変数 ■■■
'***************************************************************************************************
Private m_pTaskbarList As LongPtr
Private m_initialized As Boolean
Private m_lastState As Long
Private m_lastHwnd As LongPtr

Private m_thumbByHwnd As Object
Private WithEvents m_app As Excel.Application
Attribute m_app.VB_VarHelpID = -1
Private m_appHooked As Boolean

'***************************************************************************************************
'                               ■■■ ライフサイクル ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    m_lastState = STATE_UNSET
End Sub

Private Sub Class_Terminate()
    Call CleanupAllThumbBars
    Call ReleaseComObject
    Set m_app = Nothing
    m_appHooked = False
End Sub

Private Sub EnsureAppQuitHook()
    If Not m_appHooked Then
        Set m_app = Application
        m_appHooked = True
    End If
End Sub

Private Sub m_app_Quit()
    Call CleanupAllThumbBars
End Sub

Private Sub m_app_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    Dim w As Window

    On Error Resume Next
    For Each w In Wb.Windows
        Call RemoveThumbBarForHwnd(w.hwnd)
    Next w
    On Error GoTo 0
End Sub

'***************************************************************************************************
'                               ■■■ Public API ■■■
'***************************************************************************************************
Public Sub UpdateProgress(ByVal currentProgress As Long, Optional ByVal maxProgress As Long = 100, Optional ByVal Status As SetProgressState = TBPF_NORMAL, Optional ByVal hwnd As LongPtr)
    If hwnd = 0 Then hwnd = Application.hwnd
    If Not EnsureInit() Then Exit Sub

    If m_lastState <> CLng(Status) Or m_lastHwnd <> hwnd Then
        Call ComSetProgressState(hwnd, CLng(Status))
        m_lastState = CLng(Status)
        m_lastHwnd = hwnd
    End If

    If Status = TBPF_NORMAL Or Status = TBPF_PAUSED Or Status = TBPF_ERROR Then
        Call ComSetProgressValue(hwnd, currentProgress, maxProgress)
    End If

    If Status = TBPF_NOPROGRESS Then
        m_lastState = STATE_UNSET
    End If
End Sub

Public Sub UpdateOverlayIcon(ByVal IconPath As String, Optional ByVal IconIndex As Long = 0, Optional ByVal Description As String, Optional ByVal hwnd As LongPtr)
    Dim hIcon As LongPtr

    If hwnd = 0 Then hwnd = Application.hwnd
    If Not EnsureInit() Then Exit Sub

    If IconIndex < 0 Then
        Call ComSetOverlayIcon(hwnd, 0)
    Else
        hIcon = LoadOverlayIcon(IconPath, IconIndex)
        If hIcon = 0 Then Exit Sub

        Call ComSetOverlayIcon(hwnd, hIcon, Description)
        Call DestroyIcon(hIcon)
    End If
End Sub

Public Function ProgressSequenceForTerminal(ByVal progress As Long, Optional ByVal state As SetProgressStateForTerminal = TER_NORMAL) As String
    Dim ProgressValue As Integer

    If progress < 0 Then
        ProgressValue = 0
    ElseIf progress > 100 Then
        ProgressValue = 100
    Else
        ProgressValue = CInt(progress)
    End If

    ProgressSequenceForTerminal = "<NUL SET /p =" & Chr(27) & "]9;4;" & state & ";" & ProgressValue & Chr(7)
End Function

Public Sub Diagnose()
    Dim hr As Long
    Dim msg As String
    Dim hwnd As LongPtr

    hwnd = Application.hwnd
    msg = "hwnd=" & Hex(hwnd) & vbCrLf

    If Not EnsureInit() Then
        MsgBox msg & "CoCreateInstance に失敗しました。", vbExclamation, "Taskbar Progress 診断"
        Exit Sub
    End If

    msg = msg & "CreateITaskbarList3=" & Hex(m_pTaskbarList) & vbCrLf
    msg = msg & "HrInit=0x" & Hex(S_OK) & vbCrLf

    hr = ComSetProgressState(hwnd, TBPF_INDETERMINATE)
    msg = msg & "SetProgressState(INDETERMINATE)=0x" & Hex(hr) & vbCrLf

    hr = ComSetProgressValue(hwnd, 50, 100)
    msg = msg & "SetProgressValue(50,100)=0x" & Hex(hr) & vbCrLf

    MsgBox msg, vbInformation, "Taskbar Progress 診断"
End Sub

Public Sub Reset()
    Call CleanupAllThumbBars
    Call ReleaseComObject
    m_initialized = False
    m_lastState = STATE_UNSET
    m_lastHwnd = 0
End Sub

'***************************************************************************************************
'                               ■■■ サムネイルツールバー Public API ■■■
'***************************************************************************************************
Public Sub InitThumbBar(Optional ByVal hwnd As LongPtr)
    Dim handler As ITaskbarSubclassHandler
    Dim key As String

    hwnd = ResolveThumbHwnd(hwnd)
    key = ThumbHwndKey(hwnd)
    If ThumbRegistry.Exists(key) Then Exit Sub

    If Not EnsureInit() Then Exit Sub
    Call EnsureAppQuitHook

    Set handler = New ITaskbarSubclassHandler
    Call handler.ResetThumbButtons
    If ComThumbBarAddButtons(hwnd, MAX_THUMB_BUTTONS, handler.ButtonsPtr) <> S_OK Then Exit Sub
    If Not handler.Attach(hwnd) Then Exit Sub

    ThumbRegistry.Add key, handler
End Sub

Public Sub ConfigureThumbButton(ByVal buttonIndex As Long, ByVal ProcedureName As String, Optional ByVal IconPath As String, Optional ByVal IconIndex As Long = 0, Optional ByVal ButtonType As THUMBBUTTONFLAGS = THBF_ENABLED, Optional ByVal Description As String = vbNullString, Optional ByVal hwnd As LongPtr)
    Dim handler As ITaskbarSubclassHandler

    ValidateThumbButtonIndex buttonIndex
    hwnd = ResolveThumbHwnd(hwnd)
    Set handler = GetThumbHandler(hwnd)
    If handler Is Nothing Then Exit Sub

    If IconPath = "" Then IconPath = Application.Path & "\XLICONS.EXE"

    Call handler.ConfigureButton(buttonIndex, ProcedureName, IconPath, IconIndex, CLng(ButtonType), Description)
End Sub

Public Sub UpdateThumbButton(ByVal buttonIndex As Long, Optional ByVal hwnd As LongPtr)
    Dim handler As ITaskbarSubclassHandler
    Dim hIcon As LongPtr
    Dim IconPath As String
    Dim IconIndex As Long

    ValidateThumbButtonIndex buttonIndex
    hwnd = ResolveThumbHwnd(hwnd)
    Set handler = GetThumbHandler(hwnd)
    If handler Is Nothing Then Exit Sub
    If Not EnsureInit() Then Exit Sub

    IconPath = handler.ConfigIconPath(buttonIndex)
    IconIndex = handler.ConfigIconIndex(buttonIndex)
    hIcon = 0
    If IconIndex >= 0 Then
        hIcon = LoadThumbIcon(IconPath, IconIndex)
    End If

    Call handler.ApplyButtonUpdate(buttonIndex, hIcon)
    Call ComThumbBarUpdateButtons(hwnd, MAX_THUMB_BUTTONS, handler.ButtonsPtr)
End Sub

Public Sub ClearThumbButton(ByVal buttonIndex As Long, Optional ByVal hwnd As LongPtr)
    Dim handler As ITaskbarSubclassHandler

    If buttonIndex = 0 Then
        Call ClearAllThumbButtons(hwnd)
        Exit Sub
    End If

    ValidateThumbButtonIndex buttonIndex
    hwnd = ResolveThumbHwnd(hwnd)
    Set handler = GetThumbHandler(hwnd)
    If handler Is Nothing Then Exit Sub

    Call handler.ClearButton(buttonIndex)
    Call ComThumbBarUpdateButtons(hwnd, MAX_THUMB_BUTTONS, handler.ButtonsPtr)
End Sub

Public Sub ClearAllThumbButtons(Optional ByVal hwnd As LongPtr)
    Dim i As Long

    hwnd = ResolveThumbHwnd(hwnd)
    For i = 1 To MAX_THUMB_BUTTONS
        Call ClearThumbButton(i, hwnd)
    Next i
End Sub

Public Sub ReleaseThumbBar(Optional ByVal hwnd As LongPtr)
    hwnd = ResolveThumbHwnd(hwnd)
    Call RemoveThumbBarForHwnd(hwnd)
End Sub

'***************************************************************************************************
' 機能     ：ITaskbarList3::SetProgressState
'---------------------------------------------------------------------------------------------------
' URL      ：https://learn.microsoft.com/ja-jp/windows/desktop/api/shobjidl_core/nf-shobjidl_core-itaskbarlist3-setprogressstate
'***************************************************************************************************
Private Function ComSetProgressState(ByVal hwnd As LongPtr, ByVal state As Long) As Long
    Dim vArgs(0 To 1) As Variant
    Dim vTypes(0 To 1) As Integer
    Dim vPtrs(0 To 1) As LongPtr

    vArgs(0) = hwnd
    vArgs(1) = state
    vTypes(0) = VT_PARAM_PTR
    vTypes(1) = vbLong
    vPtrs(0) = VarPtr(vArgs(0))
    vPtrs(1) = VarPtr(vArgs(1))

    ComSetProgressState = InvokeComMethod(m_pTaskbarList, VTBL_SETPROGRESSSTATE, 2, vTypes, vPtrs)
End Function

'***************************************************************************************************
' 機能     ：ITaskbarList3::SetProgressValue
'---------------------------------------------------------------------------------------------------
' URL      ：https://learn.microsoft.com/ja-jp/windows/desktop/api/shobjidl_core/nf-shobjidl_core-itaskbarlist3-setprogressvalue
'***************************************************************************************************
Private Function ComSetProgressValue(ByVal hwnd As LongPtr, ByVal current As Long, ByVal maximum As Long) As Long
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

    ComSetProgressValue = InvokeComMethod(m_pTaskbarList, VTBL_SETPROGRESSVALUE, 3, vTypes, vPtrs)
End Function

Private Function ComThumbBarAddButtons(ByVal hwnd As LongPtr, ByVal cButtons As Long, ByVal pButtons As LongPtr) As Long
    Dim vArgs(0 To 2) As Variant
    Dim vTypes(0 To 2) As Integer
    Dim vPtrs(0 To 2) As LongPtr

    vArgs(0) = hwnd
    vArgs(1) = cButtons
    vArgs(2) = pButtons
    vTypes(0) = VT_PARAM_PTR
    vTypes(1) = vbLong
    vTypes(2) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    vPtrs(1) = VarPtr(vArgs(1))
    vPtrs(2) = VarPtr(vArgs(2))

    ComThumbBarAddButtons = InvokeComMethod(m_pTaskbarList, VTBL_THUMBBARADDBUTTONS, 3, vTypes, vPtrs)
End Function

Private Function ComThumbBarUpdateButtons(ByVal hwnd As LongPtr, ByVal cButtons As Long, ByVal pButtons As LongPtr) As Long
    Dim vArgs(0 To 2) As Variant
    Dim vTypes(0 To 2) As Integer
    Dim vPtrs(0 To 2) As LongPtr

    vArgs(0) = hwnd
    vArgs(1) = cButtons
    vArgs(2) = pButtons
    vTypes(0) = VT_PARAM_PTR
    vTypes(1) = vbLong
    vTypes(2) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    vPtrs(1) = VarPtr(vArgs(1))
    vPtrs(2) = VarPtr(vArgs(2))

    ComThumbBarUpdateButtons = InvokeComMethod(m_pTaskbarList, VTBL_THUMBBARUPDATEBUTTONS, 3, vTypes, vPtrs)
End Function

'***************************************************************************************************
' 機能     ：ITaskbarList3::SetOverlayIcon
'---------------------------------------------------------------------------------------------------
' URL      ：https://learn.microsoft.com/ja-jp/windows/win32/api/shobjidl_core/nf-shobjidl_core-itaskbarlist3-setoverlayicon
'***************************************************************************************************
Private Function ComSetOverlayIcon(ByVal hwnd As LongPtr, ByVal hIcon As LongPtr, Optional ByVal Description As String) As Long
    Dim vArgs(0 To 2) As Variant
    Dim vTypes(0 To 2) As Integer
    Dim vPtrs(0 To 2) As LongPtr

    vArgs(0) = hwnd
    vArgs(1) = hIcon
    vArgs(2) = IIf(StrPtr(Description), Description, 0)
    vTypes(0) = VT_PARAM_PTR
    vTypes(1) = VT_PARAM_PTR
    vTypes(2) = vbString
    vPtrs(0) = VarPtr(vArgs(0))
    vPtrs(1) = VarPtr(vArgs(1))
    vPtrs(2) = VarPtr(vArgs(2))

    ComSetOverlayIcon = InvokeComMethod(m_pTaskbarList, VTBL_SETOVERLAYICON, 3, vTypes, vPtrs)
End Function



'***************************************************************************************************
'                               ■■■ COM / VTable ヘルパー ■■■
'***************************************************************************************************
Private Function EnsureInit() As Boolean
    If m_pTaskbarList = 0 Then
        m_pTaskbarList = CreateITaskbarList3()
        If m_pTaskbarList = 0 Then Exit Function
    End If

    If Not m_initialized Then
        If TaskbarHrInit(m_pTaskbarList) <> S_OK Then Exit Function
        m_initialized = True
    End If

    EnsureInit = True
End Function

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

Private Sub ReleaseComObject()
    If m_pTaskbarList <> 0 Then
        Call ComRelease(m_pTaskbarList)
        m_pTaskbarList = 0
    End If
End Sub



'***************************************************************************************************
'                               ■■■ その他 ヘルパー ■■■
'***************************************************************************************************
Private Function GetFileExtension(ByVal filePath As String) As String
    Dim pos As Long

    pos = InStrRev(filePath, ".")
    If pos = 0 Then Exit Function
    GetFileExtension = LCase$(Mid$(filePath, pos + 1))
End Function

Private Function LoadOverlayIcon(ByVal IconPath As String, ByVal IconIndex As Long) As LongPtr
    Dim hIconLarge As LongPtr
    Dim hIconSmall As LongPtr
    Dim extracted As Long
    Dim ext As String
    Dim index As Long

    ext = GetFileExtension(IconPath)
    If ext <> "ico" And ext <> "exe" And ext <> "dll" Then Exit Function

    If ext = "ico" Then
        index = 0
    Else
        index = IconIndex
    End If

    extracted = ExtractIconExW(StrPtr(IconPath), index, hIconLarge, hIconSmall, 1)
    If extracted = 0 Then Exit Function

    If hIconSmall <> 0 Then
        LoadOverlayIcon = hIconSmall
        If hIconLarge <> 0 Then Call DestroyIcon(hIconLarge)
    Else
        LoadOverlayIcon = hIconLarge
    End If
End Function

Private Function ThumbHwndKey(ByVal hwnd As LongPtr) As String
    ThumbHwndKey = CStr(hwnd)
End Function

Private Function ResolveThumbHwnd(Optional ByVal hwnd As LongPtr) As LongPtr
    If hwnd = 0 Then
        On Error Resume Next
        hwnd = ActiveWindow.hwnd
        If hwnd = 0 Then hwnd = Application.hwnd
        On Error GoTo 0
    End If
    ResolveThumbHwnd = hwnd
End Function

Private Function ThumbRegistry() As Object
    If m_thumbByHwnd Is Nothing Then
        Set m_thumbByHwnd = CreateObject("Scripting.Dictionary")
        m_thumbByHwnd.CompareMode = 1
    End If
    Set ThumbRegistry = m_thumbByHwnd
End Function

Private Function GetThumbHandler(ByVal hwnd As LongPtr) As ITaskbarSubclassHandler
    Dim key As String

    key = ThumbHwndKey(hwnd)
    If ThumbRegistry.Exists(key) Then
        Set GetThumbHandler = ThumbRegistry(key)
    End If
End Function

Private Sub RemoveThumbBarForHwnd(ByVal hwnd As LongPtr)
    Dim handler As ITaskbarSubclassHandler
    Dim key As String

    If hwnd = 0 Then Exit Sub
    key = ThumbHwndKey(hwnd)
    If Not ThumbRegistry.Exists(key) Then Exit Sub

    Set handler = ThumbRegistry(key)
    Call handler.ReleaseAllThumbIcons
    Call handler.Detach
    ThumbRegistry.Remove key
    Set handler = Nothing
End Sub

Private Sub DisconnectAppQuitHook()
    Set m_app = Nothing
    m_appHooked = False
End Sub

Private Sub CleanupAllThumbBars()
    Dim keys As Variant
    Dim i As Long
    Dim hwnd As LongPtr

    Call DisconnectAppQuitHook

    If m_thumbByHwnd Is Nothing Then Exit Sub
    If m_thumbByHwnd.Count = 0 Then
        Set m_thumbByHwnd = Nothing
        Exit Sub
    End If

    keys = m_thumbByHwnd.keys
    For i = LBound(keys) To UBound(keys)
        hwnd = CLngPtr(keys(i))
        Call RemoveThumbBarForHwnd(hwnd)
    Next i
    Set m_thumbByHwnd = Nothing
End Sub

Private Function LoadThumbIcon(ByVal IconPath As String, ByVal IconIndex As Long) As LongPtr
    Dim hIconLarge As LongPtr
    Dim hIconSmall As LongPtr
    Dim extracted As Long

    extracted = ExtractIconExW(StrPtr(IconPath), IconIndex, hIconLarge, hIconSmall, 1)
    If extracted = 0 Then Exit Function

    If hIconSmall <> 0 Then
        LoadThumbIcon = hIconSmall
        If hIconLarge <> 0 Then Call DestroyIcon(hIconLarge)
    Else
        LoadThumbIcon = hIconLarge
    End If
End Function

Private Sub ValidateThumbButtonIndex(ByVal buttonIndex As Long)
    If buttonIndex < 1 Or buttonIndex > MAX_THUMB_BUTTONS Then
        Err.Raise vbObjectError + 1, "ITaskbarList3", "ボタンインデックスは 1 ～ 7 にしてください。"
    End If
End Sub
