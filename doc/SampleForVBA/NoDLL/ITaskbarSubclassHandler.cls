VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITaskbarSubclassHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'  サムネイルツールバー WM_COMMAND サブクラス（cls 完結・ToastCallbackCore 方式）
'  BuildFuncPtrCache + VirtualAlloc サンク。Public は SubclassProc (Sub) のみ。
'  THBN_CLICKED のみ VBA。終了時は ITaskbarList3 が Application.Quit / Class_Terminate で Detach。
'***************************************************************************************************
Option Explicit

#If Win64 Then
    Private Const PTRSIZE As Long = 8
    Private Const NullPtr As LongLong = 0^
    Private Const CC_STDCALL As Long = 4
    Private Const vbLongPtr As Long = vbLongLong
#Else
    #Error "CTaskbarSubclassHandler は 64 ビット VBA (x64) が必要です"
#End If

Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const MEM_RELEASE As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const THUNK_SIZE As Long = 256

Private Const WM_COMMAND As Long = &H111
Private Const THBN_CLICKED As Long = &H1800
Private Const THUMB_BUTTON_ID_BASE As Long = 1001
Private Const MAX_THUMB_BUTTONS As Long = 7
Private Const THB_FLAGS As Long = 8
Private Const THB_ICON As Long = 2
Private Const THB_TOOLTIP As Long = 4
Private Const THBF_HIDDEN As Long = 8

Private Type NATIVE_THUMBBUTTON
    dwMask As Long
    iid As Long
    iBitmap As Long
    hIcon As LongPtr
    szTip(0 To 259) As Integer
    dwFlags As Long
End Type

Private Type THUMB_BUTTON_CONFIG
    ProcedureName As String
    IconPath As String
    IconIndex As Long
    ButtonType As Long
    Description As String
End Type

Private Declare PtrSafe Function DestroyIcon Lib "user32.dll" (ByVal hIcon As LongPtr) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, _
    ByVal cc As Long, ByVal vtReturn As Integer, _
    ByVal cActuals As Long, ByRef prgvt As Any, _
    ByRef prgpvarg As Any, ByRef pvargResult As Any) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
    ByVal dest As LongPtr, ByVal src As LongPtr, ByVal length As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef dest As Any, ByRef src As Any, ByVal length As LongPtr)
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" ( _
    ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, _
    ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" ( _
    ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, _
    ByVal uIdSubclass As LongPtr) As Long
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" ( _
    ByVal hwnd As LongPtr, ByVal uMsg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" ( _
    ByVal lpModuleName As LongPtr) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" ( _
    ByVal lpLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As LongPtr, ByVal lpProcName As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub SysFreeString Lib "oleaut32" (ByVal bstr As LongPtr)

Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type TYPEATTR
    GUID As GUID
    lcid As Long
    dwReserved As Long
    memidConstructor As Long
    memidDestructor As Long
    lpstrSchema As LongPtr
    cbSizeInstance As Long
    typekind As Long
    cFuncs As Integer
    cVars As Integer
    cImplTypes As Integer
    cbSizeVft As Integer
    cbAlignment As Integer
    wTypeFlags As Integer
    wMajorVerNum As Integer
    wMinorVerNum As Integer
    tdescAlias As LongPtr
    tdescAlias2 As LongPtr
    idldescType As LongPtr
    idldescType2 As LongPtr
End Type

Private Type FUNCDESC
    memid As Long
    padding1 As Long
    lprgscode As LongPtr
    lprgelemdescParam As LongPtr
    funckind As Long
    invkind As Long
    callconv As Long
    cParams As Integer
    cParamsOpt As Integer
    oVft As Integer
    cScodes As Integer
    elemdescFunc_tdesc As LongPtr
    elemdescFunc_tdesc2 As LongPtr
    elemdescFunc_idldesc As LongPtr
    elemdescFunc_idldesc2 As LongPtr
    wFuncFlags As Integer
    padding2 As Integer
    padding3 As Long
End Type

Private m_buttons(1 To MAX_THUMB_BUTTONS) As NATIVE_THUMBBUTTON
Private m_thumbConfig(1 To MAX_THUMB_BUTTONS) As THUMB_BUTTON_CONFIG
Private m_hwnd As LongPtr
Private m_uIdSubclass As LongPtr
Private m_pSubclassThunk As LongPtr
Private m_capture(0 To 6) As LongPtr
Private m_funcPtrs As Object
Private m_attached As Boolean
Private m_ansiDefSubclassName(0 To 32) As Byte

Private Sub Class_Terminate()
    Call ReleaseAllThumbIcons
    Call Detach
    Set m_funcPtrs = Nothing
End Sub

Friend Sub SetButtonProc(ByVal buttonIndex As Long, ByVal procName As String)
    If buttonIndex >= 1 And buttonIndex <= MAX_THUMB_BUTTONS Then
        m_thumbConfig(buttonIndex).ProcedureName = procName
    End If
End Sub

Friend Sub ResetThumbButtons()
    Dim i As Long
    For i = 1 To MAX_THUMB_BUTTONS
        Call InitHiddenThumbButton(i)
        m_thumbConfig(i).ProcedureName = vbNullString
        m_thumbConfig(i).IconPath = vbNullString
        m_thumbConfig(i).IconIndex = 0
        m_thumbConfig(i).ButtonType = 0
        m_thumbConfig(i).Description = vbNullString
    Next i
End Sub

Friend Function ButtonsPtr() As LongPtr
    ButtonsPtr = VarPtr(m_buttons(1))
End Function

Friend Sub ConfigureButton( _
    ByVal buttonIndex As Long, _
    ByVal ProcedureName As String, _
    ByVal IconPath As String, _
    ByVal IconIndex As Long, _
    ByVal ButtonType As Long, _
    ByVal Description As String)

    If buttonIndex < 1 Or buttonIndex > MAX_THUMB_BUTTONS Then Exit Sub
    With m_thumbConfig(buttonIndex)
        .ProcedureName = ProcedureName
        .IconPath = IconPath
        .IconIndex = IconIndex
        .ButtonType = ButtonType
        .Description = Description
    End With
End Sub

Friend Sub ApplyButtonUpdate(ByVal buttonIndex As Long, ByVal hIcon As LongPtr)
    Dim cfg As THUMB_BUTTON_CONFIG

    If buttonIndex < 1 Or buttonIndex > MAX_THUMB_BUTTONS Then Exit Sub
    cfg = m_thumbConfig(buttonIndex)
    Call ReleaseThumbButtonIcon(buttonIndex)

    With m_buttons(buttonIndex)
        .iid = THUMB_BUTTON_ID_BASE + buttonIndex - 1
        .iBitmap = 0
        .dwMask = THB_FLAGS Or THB_ICON Or THB_TOOLTIP
        .dwFlags = cfg.ButtonType
        Call CopyThumbTip(m_buttons(buttonIndex), cfg.Description)
        .hIcon = hIcon
    End With
End Sub

Friend Sub ClearButton(ByVal buttonIndex As Long)
    If buttonIndex < 1 Or buttonIndex > MAX_THUMB_BUTTONS Then Exit Sub
    Call ReleaseThumbButtonIcon(buttonIndex)
    Call InitHiddenThumbButton(buttonIndex)
    m_thumbConfig(buttonIndex).ProcedureName = vbNullString
End Sub

Friend Function ConfigIconPath(ByVal buttonIndex As Long) As String
    If buttonIndex >= 1 And buttonIndex <= MAX_THUMB_BUTTONS Then
        ConfigIconPath = m_thumbConfig(buttonIndex).IconPath
    End If
End Function

Friend Function ConfigIconIndex(ByVal buttonIndex As Long) As Long
    If buttonIndex >= 1 And buttonIndex <= MAX_THUMB_BUTTONS Then
        ConfigIconIndex = m_thumbConfig(buttonIndex).IconIndex
    End If
End Function

Friend Sub ReleaseAllThumbIcons()
    Dim i As Long
    For i = 1 To MAX_THUMB_BUTTONS
        Call ReleaseThumbButtonIcon(i)
    Next i
End Sub

Private Sub InitHiddenThumbButton(ByVal buttonIndex As Long)
    With m_buttons(buttonIndex)
        .dwMask = THB_FLAGS
        .iid = THUMB_BUTTON_ID_BASE + buttonIndex - 1
        .iBitmap = 0
        .hIcon = 0
        .dwFlags = THBF_HIDDEN
        Call ClearThumbTip(m_buttons(buttonIndex))
    End With
End Sub

Private Sub ClearThumbTip(ByRef btn As NATIVE_THUMBBUTTON)
    Dim i As Long
    For i = LBound(btn.szTip) To UBound(btn.szTip)
        btn.szTip(i) = 0
    Next i
End Sub

Private Sub CopyThumbTip(ByRef btn As NATIVE_THUMBBUTTON, ByVal tip As String)
    Dim i As Long
    Dim ch As Integer

    Call ClearThumbTip(btn)
    For i = 1 To Len(tip)
        If i - 1 > UBound(btn.szTip) Then Exit For
        ch = AscW(Mid$(tip, i, 1))
        btn.szTip(i - 1) = ch
    Next i
End Sub

Private Sub ReleaseThumbButtonIcon(ByVal buttonIndex As Long)
    If m_buttons(buttonIndex).hIcon <> 0 Then
        Call DestroyIcon(m_buttons(buttonIndex).hIcon)
        m_buttons(buttonIndex).hIcon = 0
    End If
End Sub

Friend Function Attach(ByVal hwnd As LongPtr) As Boolean
    Dim pProc As LongPtr
    Dim pCapture As LongPtr

    If hwnd = 0 Then Exit Function
    If m_attached And m_hwnd = hwnd Then
        Attach = True
        Exit Function
    End If

    Call Detach
    m_hwnd = hwnd
    m_uIdSubclass = 1
    pCapture = VarPtr(m_capture(0))

    If m_funcPtrs Is Nothing Then
        Set m_funcPtrs = CreateObject("Scripting.Dictionary")
        m_funcPtrs.CompareMode = 1
    End If
    If m_funcPtrs.Count = 0 Then BuildFuncPtrCache ObjPtr(Me), m_funcPtrs
    If Not m_funcPtrs.Exists("SubclassProc") Then Exit Function

    pProc = m_funcPtrs("SubclassProc")
    m_pSubclassThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If m_pSubclassThunk = 0 Then Exit Function

    If Not WriteSubclassThunkAt(m_pSubclassThunk, ObjPtr(Me), pProc, pCapture, DefSubclassProcAddr()) Then
        Call VirtualFree(m_pSubclassThunk, 0, MEM_RELEASE)
        m_pSubclassThunk = 0
        Exit Function
    End If

    If SetWindowSubclass(hwnd, m_pSubclassThunk, m_uIdSubclass, pCapture) = 0 Then
        Call VirtualFree(m_pSubclassThunk, 0, MEM_RELEASE)
        m_pSubclassThunk = 0
        Exit Function
    End If

    m_attached = True
    Attach = True
End Function

Friend Sub Detach()
    If m_attached And m_hwnd <> 0 And m_pSubclassThunk <> 0 Then
        Call RemoveWindowSubclass(m_hwnd, m_pSubclassThunk, m_uIdSubclass)
    End If

    If m_pSubclassThunk <> 0 Then
        Call VirtualFree(m_pSubclassThunk, 0, MEM_RELEASE)
        m_pSubclassThunk = 0
    End If

    m_hwnd = 0
    m_attached = False
End Sub

Public Sub SubclassProc( _
    ByVal hwnd As LongPtr, _
    ByVal msg As Long, _
    ByVal wParam As LongPtr, _
    ByVal lParam As LongPtr, _
    ByVal uIdSubclass As LongPtr, _
    ByVal pCapture As LongPtr)

    Dim vbaIndex As Long
    Dim procName As String

    vbaIndex = LoWordPtr(wParam) - THUMB_BUTTON_ID_BASE + 1
    If vbaIndex >= 1 And vbaIndex <= MAX_THUMB_BUTTONS Then
        procName = m_thumbConfig(vbaIndex).ProcedureName
        If Len(procName) > 0 Then
            On Error Resume Next
            Application.Run procName
            On Error GoTo 0
        End If
    End If
    m_capture(6) = 0
End Sub

Private Function DefSubclassProcAddr() As LongPtr
    Static cached As LongPtr
    Static nameReady As Boolean

    If Not nameReady Then
        Call AsciiNameToBuffer("DefSubclassProc", m_ansiDefSubclassName)
        nameReady = True
    End If

    If cached = 0 Then
        cached = GetProcAddress(Comctl32ModuleHandle(), VarPtr(m_ansiDefSubclassName(0)))
    End If
    DefSubclassProcAddr = cached
End Function

Private Function Comctl32ModuleHandle() As LongPtr
    Static hMod As LongPtr

    If hMod = 0 Then
        hMod = GetModuleHandle(StrPtr("comctl32.dll"))
        If hMod = 0 Then hMod = LoadLibrary(StrPtr("comctl32.dll"))
    End If
    Comctl32ModuleHandle = hMod
End Function

Private Sub AsciiNameToBuffer(ByVal name As String, ByRef buf() As Byte)
    Dim i As Long

    For i = 1 To Len(name)
        buf(i - 1) = Asc(Mid$(name, i, 1))
    Next i
    buf(Len(name)) = 0
End Sub

Private Function WriteSubclassThunkAt( _
    ByVal addr As LongPtr, _
    ByVal pOwner As LongPtr, _
    ByVal pFunc As LongPtr, _
    ByVal pCapture As LongPtr, _
    ByVal pDefSubclass As LongPtr) As Boolean

    Dim b(0 To THUNK_SIZE - 1) As Byte
    Dim i As Long
    Dim k As Long
    Dim jneNotCommand As Long
    Dim jeThumbClick As Long
    Dim jmpFast As Long
    Dim vbaStart As Long
    Dim fastStart As Long

    If addr = 0 Or pOwner = 0 Or pFunc = 0 Or pCapture = 0 Or pDefSubclass = 0 Then Exit Function

    i = 0
    b(i) = &H81: b(i + 1) = &HFA: b(i + 2) = &H11: b(i + 3) = &H1: b(i + 4) = &H0: b(i + 5) = &H0: i = i + 6  ' cmp edx, WM_COMMAND
    b(i) = &HF: b(i + 1) = &H85: jneNotCommand = i + 2: i = i + 6                                        ' jne fastStart

    b(i) = &H41: b(i + 1) = &H8B: b(i + 2) = &HC0: i = i + 3                                            ' mov eax, r8d
    b(i) = &HC1: b(i + 1) = &HE8: b(i + 2) = &H10: i = i + 3                                            ' shr eax, 16
    b(i) = &H3D: b(i + 1) = &H0: b(i + 2) = &H18: b(i + 3) = &H0: b(i + 4) = &H0: i = i + 5            ' cmp eax, THBN_CLICKED
    b(i) = &HF: b(i + 1) = &H84: jeThumbClick = i + 2: i = i + 6                                         ' je vbaStart

    b(i) = &HE9: jmpFast = i + 1: i = i + 5                                                              ' jmp fastStart

    vbaStart = i
    b(i) = &H53: i = i + 1                                                                                ' push rbx
    b(i) = &H57: i = i + 1                                                                                ' push rdi
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H48: i = i + 4                           ' sub rsp, 48h

    b(i) = &H48: b(i + 1) = &H8B: b(i + 2) = &HBC: b(i + 3) = &H24: b(i + 4) = &H88: b(i + 5) = &H0: b(i + 6) = &H0: b(i + 7) = &H0: i = i + 8
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HF: i = i + 3
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H57: b(i + 3) = &H8: i = i + 4
    b(i) = &H4C: b(i + 1) = &H89: b(i + 2) = &H47: b(i + 3) = &H10: i = i + 4
    b(i) = &H4C: b(i + 1) = &H89: b(i + 2) = &H4F: b(i + 3) = &H18: i = i + 4
    b(i) = &H48: b(i + 1) = &H8B: b(i + 2) = &H84: b(i + 3) = &H24: b(i + 4) = &H80: b(i + 5) = &H0: b(i + 6) = &H0: b(i + 7) = &H0: i = i + 8
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H47: b(i + 3) = &H20: i = i + 4
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H7F: b(i + 3) = &H28: i = i + 4

    b(i) = &H48: b(i + 1) = &H8B: b(i + 2) = &H17: i = i + 3
    b(i) = &H4C: b(i + 1) = &H8B: b(i + 2) = &H47: b(i + 3) = &H8: i = i + 4
    b(i) = &H4C: b(i + 1) = &H8B: b(i + 2) = &H4F: b(i + 3) = &H10: i = i + 4
    b(i) = &H48: b(i + 1) = &H8B: b(i + 2) = &H47: b(i + 3) = &H18: i = i + 4
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5
    b(i) = &H48: b(i + 1) = &H8B: b(i + 2) = &H47: b(i + 3) = &H20: i = i + 4
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H30: i = i + 5
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H7C: b(i + 3) = &H24: b(i + 4) = &H38: i = i + 5
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pOwner: i = i + 8
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pFunc: i = i + 8
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2

    b(i) = &H48: b(i + 1) = &H8B: b(i + 2) = &H47: b(i + 3) = &H30: i = i + 4
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H48: i = i + 4
    b(i) = &H5F: i = i + 1
    b(i) = &H5B: i = i + 1
    b(i) = &HC3: i = i + 1

    fastStart = i
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H28: i = i + 4
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pDefSubclass: i = i + 8
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H28: i = i + 4
    b(i) = &HC3: i = i + 1

    Call PatchRel32(b, jneNotCommand, fastStart)
    Call PatchRel32(b, jeThumbClick, vbaStart)
    Call PatchRel32(b, jmpFast, fastStart)

    For k = 0 To ((i + 7) \ 8) * 8 - 1 Step 8
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k

    WriteSubclassThunkAt = True
End Function

Private Sub PatchRel32(ByRef b() As Byte, ByVal patchAt As Long, ByVal target As Long)
    Dim rel As Long
    rel = target - (patchAt + 4)
    b(patchAt) = rel And &HFF&
    b(patchAt + 1) = (rel \ &H100&) And &HFF&
    b(patchAt + 2) = (rel \ &H10000) And &HFF&
    b(patchAt + 3) = (rel \ &H1000000) And &HFF&
End Sub

Private Sub BuildFuncPtrCache(ByVal pObj As LongPtr, ByVal dict As Object)
    Dim pTI As LongPtr
    Dim hr As Long

    dict.RemoveAll
    hr = CallCOM3(pObj, 4, 0&, 0&, VarPtr(pTI))
    If hr <> 0 Then Exit Sub

    Dim pTA As LongPtr
    hr = CallCOM1(pTI, 3, VarPtr(pTA))
    If hr <> 0 Then GoTo Cleanup

    Dim uTA As TYPEATTR
    CopyMemory uTA, ByVal pTA, LenB(uTA)
    Dim pVT As LongPtr
    CopyMemory pVT, ByVal pObj, LenB(pVT)

    Dim n As Long
    For n = 0 To uTA.cFuncs - 1
        Dim pFD As LongPtr
        hr = CallCOM2(pTI, 5, CLng(n), VarPtr(pFD))
        If hr <> 0 Then GoTo NextF
        Dim uFD As FUNCDESC
        CopyMemory uFD, ByVal pFD, LenB(uFD)
        Dim pB As LongPtr
        pB = 0
        hr = CallCOM5(pTI, 12, uFD.memid, VarPtr(pB), 0, 0, 0)
        If hr = 0 And pB <> 0 Then
            Dim fn As String
            fn = PtrToStrW(pB)
            SysFreeString pB
            pB = 0
            Dim pa As LongPtr
            CopyMemory pa, ByVal (pVT + uFD.oVft), LenB(pa)
            dict(fn) = pa
        End If
        CallCOM1V pTI, 20, pFD
NextF:
    Next n
    CallCOM1V pTI, 19, pTA
Cleanup:
    CallCOM0V pTI, 2
End Sub

Private Function PtrToStrW(ByVal p As LongPtr) As String
    Dim cch As Long

    If p = 0 Then Exit Function
    cch = lstrlenW(p)
    If cch = 0 Then Exit Function
    PtrToStrW = String$(cch, vbNullChar)
    RtlMoveMemory StrPtr(PtrToStrW), p, CLngPtr(cch * 2)
End Function

Private Function ReadLongPtrFromBytes(ByRef b() As Byte, ByVal offset As Long) As LongPtr
    ReadLongPtrFromBytes = ReadLongPtr(VarPtr(b(offset)))
End Function

Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PTRSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        Call WritePtrNatively(pa, VarPtr(.sa))
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PTRSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        Call WritePtrNatively(pa, VarPtr(.sa))
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub

Private Function LoWordPtr(ByVal value As LongPtr) As Long
    LoWordPtr = CLng(value) And &HFFFF&
End Function

Private Function CallCOM0V(ByVal p As LongPtr, ByVal idx As Long) As Long
    Dim r As Variant
    CallCOM0V = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 0, 0, 0, r)
End Function

Private Function CallCOM1(ByVal p As LongPtr, ByVal idx As Long, ByVal a1 As LongPtr) As Long
    Dim r As Variant
    Dim v(0) As Variant
    Dim t(0) As Integer
    Dim vp(0) As LongPtr
    v(0) = a1: t(0) = vbLongPtr: vp(0) = VarPtr(v(0))
    CallCOM1 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 1, t(0), vp(0), r)
    If CallCOM1 = 0 Then CallCOM1 = CLng(r)
End Function

Private Sub CallCOM1V(ByVal p As LongPtr, ByVal idx As Long, ByVal a1 As LongPtr)
    Dim r As Variant
    Dim v(0) As Variant
    Dim t(0) As Integer
    Dim vp(0) As LongPtr
    v(0) = a1: t(0) = vbLongPtr: vp(0) = VarPtr(v(0))
    DispCallFunc p, idx * LenB(p), CC_STDCALL, vbLong, 1, t(0), vp(0), r
End Sub

Private Function CallCOM2(ByVal p As LongPtr, ByVal idx As Long, _
    ByVal a1 As Long, ByVal a2 As LongPtr) As Long

    Dim r As Variant
    Dim v(1) As Variant
    Dim t(1) As Integer
    Dim vp(1) As LongPtr
    v(0) = a1: v(1) = a2: t(0) = vbLong: t(1) = vbLongPtr
    vp(0) = VarPtr(v(0)): vp(1) = VarPtr(v(1))
    CallCOM2 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 2, t(0), vp(0), r)
    If CallCOM2 = 0 Then CallCOM2 = CLng(r)
End Function

Private Function CallCOM3(ByVal p As LongPtr, ByVal idx As Long, _
    ByVal a1 As Long, ByVal a2 As Long, ByVal a3 As LongPtr) As Long

    Dim r As Variant
    Dim v(2) As Variant
    Dim t(2) As Integer
    Dim vp(2) As LongPtr
    v(0) = a1: v(1) = a2: v(2) = a3
    t(0) = vbLong: t(1) = vbLong: t(2) = vbLongPtr
    vp(0) = VarPtr(v(0)): vp(1) = VarPtr(v(1)): vp(2) = VarPtr(v(2))
    CallCOM3 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 3, t(0), vp(0), r)
    If CallCOM3 = 0 Then CallCOM3 = CLng(r)
End Function

Private Function CallCOM5(ByVal p As LongPtr, ByVal idx As Long, _
    ByVal a1 As Long, ByVal a2 As LongPtr, ByVal a3 As LongPtr, _
    ByVal a4 As LongPtr, ByVal a5 As LongPtr) As Long

    Dim r As Variant
    Dim v(4) As Variant
    Dim t(4) As Integer
    Dim vp(4) As LongPtr
    v(0) = a1: v(1) = a2: v(2) = a3: v(3) = a4: v(4) = a5
    t(0) = vbLong: t(1) = vbLongPtr: t(2) = vbLongPtr: t(3) = vbLongPtr: t(4) = vbLongPtr
    vp(0) = VarPtr(v(0)): vp(1) = VarPtr(v(1)): vp(2) = VarPtr(v(2))
    vp(3) = VarPtr(v(3)): vp(4) = VarPtr(v(4))
    CallCOM5 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 5, t(0), vp(0), r)
    If CallCOM5 = 0 Then CallCOM5 = CLng(r)
End Function
