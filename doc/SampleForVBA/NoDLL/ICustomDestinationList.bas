Attribute VB_Name = "Mod05_JumplistControl"
'***************************************************************************************************
'                           Windows 7以降にあるジャンプリストをVBAのみで制御します
'                                       (DLLフリー版)
'***************************************************************************************************
Option Explicit

'***************************************************************************************************
'                           ■■■ メンバ変数/ユーザー定義型宣言 ■■■
'***************************************************************************************************
Private Const AppUserModelID_Excel  As String = "Microsoft.Office.EXCEL.EXE.15" 'DeskTopアプリ：Excel
Private JumpListItems               As Collection                               'Jumpリスト情報保持用



'***************************************************************************************************
'               ■■■ VBA単体 (DispCallFunc × VTable) 呼び出し用セクション ■■■
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

' propsys.dll の InitPropVariantFromString は環境によりエクスポートされていない事があるため、
' CoTaskMemAlloc + 自前メモリコピーで VT_LPWSTR 形式の PROPVARIANT を構築する方式を取る。
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" ( _
    ByVal cb As LongPtr) As LongPtr

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As LongPtr, _
    ByVal Source As LongPtr, _
    ByVal Length As LongPtr)

Private Declare PtrSafe Function PropVariantClear Lib "ole32.dll" ( _
    ByRef ppropvar As Any) As Long

'------------------------------------構造体-------------------------------------
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PROPERTYKEY
    fmtid As GUID
    pid As Long
End Type

' PROPVARIANT (ネイティブサイズ: x86=16byte, x64=24byte)
#If Win64 Then
Private Type PROPVARIANT
    vt As Integer
    wReserved1 As Integer
    wReserved2 As Integer
    wReserved3 As Integer
    pVal As LongPtr        ' 8 byte (union)
    padding As LongPtr     ' 8 byte (union 残り)
End Type
#Else
Private Type PROPVARIANT
    vt As Integer
    wReserved1 As Integer
    wReserved2 As Integer
    wReserved3 As Integer
    pVal As LongPtr        ' 4 byte (union)
    padding As Long        ' 4 byte (union 残り)
End Type
#End If

'------------------------------------COM定数-------------------------------------
Private Const CC_STDCALL As Long = 4
Private Const CLSCTX_INPROC_SERVER As Long = 1
Private Const S_OK As Long = 0

#If Win64 Then
    Private Const PTRSIZE As Long = 8
    Private Const VT_PARAM_PTR As Integer = vbLongLong
#Else
    Private Const PTRSIZE As Long = 4
    Private Const VT_PARAM_PTR As Integer = vbLong
#End If

' --- GUID 文字列定義 ---
' CoClassのCLSIDはインターフェースのIIDとは別物。以下はShObjIdl_core.hに準拠
Private Const CLSID_DestinationList            As String = "{77F10CF0-3DB5-4966-B520-B7C54FD35ED6}"
Private Const IID_ICustomDestinationList       As String = "{6332DEBF-87B5-4670-90C0-5E57B408A49E}"

Private Const CLSID_EnumerableObjectCollection As String = "{2D3468C1-36A7-43B6-AC24-D3F02FD9607A}"
Private Const IID_IObjectCollection            As String = "{5632B1A4-E38A-400A-928A-D4CD63230295}"
Private Const IID_IObjectArray                 As String = "{92CA9DCD-5622-4BBA-A805-5E9F541BD8C9}"

Private Const CLSID_ShellLink                  As String = "{00021401-0000-0000-C000-000000000046}"
Private Const IID_IShellLinkW                  As String = "{000214F9-0000-0000-C000-000000000046}"

Private Const IID_IPropertyStore               As String = "{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}"

' --- Vtable インデックス定義 ---
' IUnknown
Private Const VTBL_RELEASE As Long = 2

' ICustomDestinationList (ShObjIdl_core.h の定義順)
Private Const VTBL_SETAPPID As Long = 3
Private Const VTBL_BEGINLIST As Long = 4
Private Const VTBL_APPENDCATEGORY As Long = 5
Private Const VTBL_APPENDKNOWNCOLL As Long = 6
Private Const VTBL_ADDUSERTASKS As Long = 7
Private Const VTBL_COMMITLIST As Long = 8
Private Const VTBL_GETREMOVEDDESTS As Long = 9
Private Const VTBL_DELETELIST As Long = 10
Private Const VTBL_ABORTLIST As Long = 11

' IObjectArray
Private Const VTBL_GETCOUNT As Long = 3
Private Const VTBL_GETAT As Long = 4

' IObjectCollection (IObjectArray継承)
Private Const VTBL_ADDOBJECT As Long = 5
Private Const VTBL_CLEAR As Long = 8

' IShellLinkW
Private Const VTBL_SETDESCRIPTION As Long = 7
Private Const VTBL_SETARGUMENTS As Long = 11
Private Const VTBL_SETICONLOCATION As Long = 17
Private Const VTBL_SETPATH As Long = 20

' IPropertyStore
Private Const VTBL_SETVALUE As Long = 6
Private Const VTBL_COMMIT As Long = 7



'***************************************************************************************************
'                           ■■■ ジャンプリスト制御公開関数 ■■■
'***************************************************************************************************
'* 機能    ：ジャンプリストに登録するショートカット情報をメモリ上の Collection に蓄積します。
'---------------------------------------------------------------------------------------------------
'* 引数    ：表示名        ジャンプリストに表示するタイトル
'*           実行パス      起動する対象プログラムの実行パス
'*           コマンド引数  コマンドライン引数 (Optional)
'*           カテゴリ名    独自のカテゴリ名。指定しない（空文字）場合は「タスク」に登録されます (Optional)
'*           説明文        マウスホバー時に表示されるツールチップ文 (Optional)
'*           アイコンパス  アイコンを保持するファイルパス (.exe / .dll / .ico) (Optional)
'*           アイコンIndex アイコンの格納インデックス番号 (Optional)
'***************************************************************************************************
Public Sub Registration(ByVal 表示名 As String, ByVal 実行パス As String, Optional ByVal コマンド引数 As String, Optional ByVal カテゴリ名 As String, Optional ByVal 説明文 As String, Optional ByVal アイコンパス As String, Optional ByVal アイコンIndex As Long)
    ' アイコン未指定時はExcelのアイコンファイルを適用
    If アイコンパス = "" Then アイコンパス = Application.Path & "\XLICONS.EXE"

    ' Collectionを初期化
    If JumpListItems Is Nothing Then Set JumpListItems = New Collection

    ' VBAの制限(標準モジュールUDTをCollectionに追加できない)を回避するため、内部Dictionaryを使用
    Dim target As Dictionary
    Set target = New Dictionary
    
    target.Add "カテゴリ名", カテゴリ名
    target.Add "表示名", 表示名
    target.Add "実行パス", 実行パス
    target.Add "コマンド引数", コマンド引数
    target.Add "アイコンパス", アイコンパス
    target.Add "アイコンIndex", アイコンIndex
    target.Add "説明文", 説明文
    
    JumpListItems.Add target
End Sub

'***************************************************************************************************
'* 機能    ：Registration で蓄積した設定情報を COM API 経由でジャンプリストに一括反映します。
'---------------------------------------------------------------------------------------------------
'* 引数    ：TargetApplicationModelUserID   アプリ固有の AppUserModelID (既定値：Excel)
'***************************************************************************************************
Public Sub Import(Optional TargetApplicationModelUserID As String = AppUserModelID_Excel)
    If JumpListItems Is Nothing Then Set JumpListItems = New Collection

    Debug.Print "=== JumpList Import Start ==="
    Debug.Print "AppUserModelID: " & TargetApplicationModelUserID
    Debug.Print "Items Count: " & JumpListItems.Count

    Dim pDestList As LongPtr
    pDestList = CreateCustomDestinationList()
    If pDestList = 0 Then
        Debug.Print "CreateCustomDestinationList failed!"
        Exit Sub
    End If
    Debug.Print "CustomDestinationList ptr: 0x" & Hex(pDestList)

    Dim hr As Long
    hr = SetAppID(pDestList, TargetApplicationModelUserID)
    Debug.Print "SetAppID result: 0x" & Hex(hr)
    If hr <> S_OK Then
        Call ComRelease(pDestList)
        Exit Sub
    End If

    ' BeginListで初期設定を開始 (削除済みリストは今回は受け流して解放する)
    Dim maxSlots As Long
    Dim iidObjArray As GUID
    Dim pRemovedItems As LongPtr
    
    If IIDFromString(StrPtr(IID_IObjectArray), iidObjArray) <> S_OK Then
        Debug.Print "IIDFromString(IObjectArray) failed!"
        Call ComRelease(pDestList)
        Exit Sub
    End If

    hr = BeginList(pDestList, maxSlots, iidObjArray, pRemovedItems)
    Debug.Print "BeginList result: 0x" & Hex(hr) & " maxSlots: " & maxSlots
    If hr <> S_OK Then
        Call ComRelease(pDestList)
        Exit Sub
    End If

    If pRemovedItems <> 0 Then
        Call ComRelease(pRemovedItems)
    End If

    ' データが蓄積されている場合のみ追加処理を実行
    If JumpListItems.Count > 0 Then
        ' --- 1. カスタムカテゴリの追加 ---
        Dim categories As New Collection
        Dim item As Dictionary
        Dim cat As Variant
        
        On Error Resume Next
        For Each item In JumpListItems
            Dim itemCat As String
            itemCat = item("カテゴリ名")
            If itemCat <> "" Then
                categories.Add itemCat, itemCat
            End If
        Next item
        On Error GoTo 0

        For Each cat In categories
            Debug.Print "Adding Category: " & cat
            Dim pCollection As LongPtr
            pCollection = CreateObjectCollection()
            If pCollection <> 0 Then
                For Each item In JumpListItems
                    If item("カテゴリ名") = cat Then
                        Dim pLink As LongPtr
                        pLink = CreateShellLink(item)
                        If pLink <> 0 Then
                            hr = AddObjectToCollection(pCollection, pLink)
                            Debug.Print "AddObjectToCollection result: 0x" & Hex(hr)
                            Call ComRelease(pLink)
                        End If
                    End If
                Next item
                
                ' カテゴリ追加
                hr = AppendCategoryToDestList(pDestList, CStr(cat), pCollection)
                Debug.Print "AppendCategoryToDestList result: 0x" & Hex(hr)
                Call ComRelease(pCollection)
            End If
        Next cat

        ' --- 2. タスク(カテゴリ名なし)の追加 ---
        Dim hasTasks As Boolean
        hasTasks = False
        For Each item In JumpListItems
            If item("カテゴリ名") = "" Then
                hasTasks = True
                Exit For
            End If
        Next item

        If hasTasks Then
            Debug.Print "Adding Tasks Section..."
            Dim pTaskCollection As LongPtr
            pTaskCollection = CreateObjectCollection()
            If pTaskCollection <> 0 Then
                For Each item In JumpListItems
                    If item("カテゴリ名") = "" Then
                        Dim pTaskLink As LongPtr
                        pTaskLink = CreateShellLink(item)
                        If pTaskLink <> 0 Then
                            hr = AddObjectToCollection(pTaskCollection, pTaskLink)
                            Debug.Print "AddObjectToCollection (Task) result: 0x" & Hex(hr)
                            Call ComRelease(pTaskLink)
                        End If
                    End If
                Next item
                
                ' タスクリスト追加
                hr = AddUserTasksToDestList(pDestList, pTaskCollection)
                Debug.Print "AddUserTasksToDestList result: 0x" & Hex(hr)
                Call ComRelease(pTaskCollection)
            End If
        End If
    End If

    ' コミットしてジャンプリストを反映
    hr = CommitList(pDestList)
    Debug.Print "CommitList result: 0x" & Hex(hr)
    Call ComRelease(pDestList)

    ' 登録情報をリセット
    Set JumpListItems = New Collection
    Debug.Print "=== JumpList Import End ==="
End Sub

'***************************************************************************************************
'* 機能    ：対象アプリケーションのカスタムジャンプリストを完全にクリア(削除)します。
'---------------------------------------------------------------------------------------------------
'* 引数    ：TargetApplicationModelUserID   アプリ固有の AppUserModelID (既定値：Excel)
'***************************************************************************************************
Public Sub Clear(Optional TargetApplicationModelUserID As String = AppUserModelID_Excel)
    Dim pDestList As LongPtr
    pDestList = CreateCustomDestinationList()
    If pDestList = 0 Then Exit Sub

    Call DeleteList(pDestList, TargetApplicationModelUserID)
    Call ComRelease(pDestList)
    
    ' メモリコレクションも初期化
    Set JumpListItems = New Collection
End Sub


'***************************************************************************************************
'                           ■■■ COM 呼び出し・操作ヘルパー関数群 ■■■
'***************************************************************************************************

' PROPVARIANT を VT_LPWSTR 形式で初期化する (InitPropVariantFromString の代替)
' 解放は PropVariantClear (CoTaskMemFree が走る) が行うので、CoTaskMemAlloc で確保する事が重要。
Private Function InitPropVariantAsLPWSTR(ByVal s As String, ByRef pv As PROPVARIANT) As Long
    Const VT_LPWSTR As Integer = 31
    Const E_OUTOFMEMORY As Long = &H8007000E

    ' バイト長 = (文字数 + 終端NUL) * 2
    Dim cb As LongPtr
    cb = (LenB(s) + 2)

    Dim p As LongPtr
    p = CoTaskMemAlloc(cb)
    If p = 0 Then
        InitPropVariantAsLPWSTR = E_OUTOFMEMORY
        Exit Function
    End If

    ' BSTR本体 (StrPtrの指す領域) には終端NULも含まれているので、丸ごとコピー
    Call CopyMemory(p, StrPtr(s), cb)

    pv.vt = VT_LPWSTR
    pv.pVal = p
    InitPropVariantAsLPWSTR = S_OK
End Function

' VTable経由のCOMメソッド呼び出し
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

' IUnknown::Release
Private Sub ComRelease(ByVal pInterface As LongPtr)
    Dim vResult As Variant
    If pInterface = 0 Then Exit Sub
    Call DispCallFunc(pInterface, VTBL_RELEASE * PTRSIZE, CC_STDCALL, vbLong, 0, 0, 0, vResult)
End Sub

' IUnknown::QueryInterface
Private Function QueryInterface(ByVal pInterface As LongPtr, ByRef riid As GUID, ByRef ppv As LongPtr) As Long
    Dim vArgs(0 To 1) As Variant
    Dim vTypes(0 To 1) As Integer
    Dim vPtrs(0 To 1) As LongPtr
    
    vArgs(0) = VarPtr(riid)
    vArgs(1) = VarPtr(ppv)
    vTypes(0) = VT_PARAM_PTR
    vTypes(1) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    vPtrs(1) = VarPtr(vArgs(1))
    
    QueryInterface = InvokeComMethod(pInterface, 0, 2, vTypes, vPtrs)
End Function

' ICustomDestinationList インスタンス生成
Private Function CreateCustomDestinationList() As LongPtr
    Dim clsid As GUID
    Dim iid As GUID
    Dim pDestList As LongPtr
    Dim hr As Long
    
    hr = IIDFromString(StrPtr(CLSID_DestinationList), clsid)
    If hr <> S_OK Then
        Debug.Print "  CreateCustomDestinationList: IIDFromString(CLSID) failed: 0x" & Hex(hr)
        Exit Function
    End If
    
    ' CLSIDパース結果のダンプ
    Debug.Print "  CLSID Dump: Data1=" & Hex(clsid.Data1) & " Data2=" & Hex(clsid.Data2) & " Data3=" & Hex(clsid.Data3) & _
                " Data4=" & Hex(clsid.Data4(0)) & Hex(clsid.Data4(1)) & Hex(clsid.Data4(2)) & Hex(clsid.Data4(3)) & _
                Hex(clsid.Data4(4)) & Hex(clsid.Data4(5)) & Hex(clsid.Data4(6)) & Hex(clsid.Data4(7))
    
    hr = IIDFromString(StrPtr(IID_ICustomDestinationList), iid)
    If hr <> S_OK Then
        Debug.Print "  CreateCustomDestinationList: IIDFromString(IID) failed: 0x" & Hex(hr)
        Exit Function
    End If
    
    ' CLSCTX_ALL (23) での生成を試行
    hr = CoCreateInstance(clsid, 0, 23, iid, pDestList)
    If hr <> S_OK Then
        Debug.Print "  CreateCustomDestinationList: CoCreateInstance failed: 0x" & Hex(hr)
        Exit Function
    End If
    
    CreateCustomDestinationList = pDestList
End Function

' IObjectCollection インスタンス生成
Private Function CreateObjectCollection() As LongPtr
    Dim clsid As GUID
    Dim iid As GUID
    Dim pCollection As LongPtr
    Dim hr As Long
    
    hr = IIDFromString(StrPtr(CLSID_EnumerableObjectCollection), clsid)
    If hr <> S_OK Then
        Debug.Print "  CreateObjectCollection: IIDFromString(CLSID) failed: 0x" & Hex(hr)
        Exit Function
    End If
    
    hr = IIDFromString(StrPtr(IID_IObjectCollection), iid)
    If hr <> S_OK Then
        Debug.Print "  CreateObjectCollection: IIDFromString(IID) failed: 0x" & Hex(hr)
        Exit Function
    End If
    
    hr = CoCreateInstance(clsid, 0, CLSCTX_INPROC_SERVER, iid, pCollection)
    If hr <> S_OK Then
        Debug.Print "  CreateObjectCollection: CoCreateInstance failed: 0x" & Hex(hr)
        Exit Function
    End If
    
    CreateObjectCollection = pCollection
End Function

' IShellLinkW インスタンス生成と各種設定
Private Function CreateShellLink(item As Dictionary) As LongPtr
    Dim clsid As GUID
    Dim iid As GUID
    Dim pLink As LongPtr
    Dim hr As Long
    
    hr = IIDFromString(StrPtr(CLSID_ShellLink), clsid)
    If hr <> S_OK Then
        Debug.Print "  IIDFromString(CLSID_ShellLink) failed: 0x" & Hex(hr)
        Exit Function
    End If
    
    hr = IIDFromString(StrPtr(IID_IShellLinkW), iid)
    If hr <> S_OK Then
        Debug.Print "  IIDFromString(IID_IShellLinkW) failed: 0x" & Hex(hr)
        Exit Function
    End If
    
    hr = CoCreateInstance(clsid, 0, CLSCTX_INPROC_SERVER, iid, pLink)
    If hr <> S_OK Then
        Debug.Print "  CoCreateInstance(ShellLink) failed: 0x" & Hex(hr)
        Exit Function
    End If
    
    Debug.Print "  ShellLink instance created: 0x" & Hex(pLink)
    
    ' 各プロパティをセット
    hr = SetShellLinkPath(pLink, item("実行パス"))
    Debug.Print "  SetShellLinkPath result: 0x" & Hex(hr)
    
    If item("コマンド引数") <> "" Then
        hr = SetShellLinkArguments(pLink, item("コマンド引数"))
        Debug.Print "  SetShellLinkArguments result: 0x" & Hex(hr)
    End If
    
    If item("説明文") <> "" Then
        hr = SetShellLinkDescription(pLink, item("説明文"))
        Debug.Print "  SetShellLinkDescription result: 0x" & Hex(hr)
    End If
    
    If item("アイコンパス") <> "" Then
        hr = SetShellLinkIconLocation(pLink, item("アイコンパス"), item("アイコンIndex"))
        Debug.Print "  SetShellLinkIconLocation result: 0x" & Hex(hr)
    End If
    
    ' 表示名(Title)の登録 (IPropertyStore経由で設定)
    If item("表示名") <> "" Then
        Dim pPropStore As LongPtr
        Dim iidPropStore As GUID
        hr = IIDFromString(StrPtr(IID_IPropertyStore), iidPropStore)
        If hr = S_OK Then
            hr = QueryInterface(pLink, iidPropStore, pPropStore)
            Debug.Print "  QueryInterface(IPropertyStore) result: 0x" & Hex(hr) & " ptr: 0x" & Hex(pPropStore)
            If hr = S_OK And pPropStore <> 0 Then
                Dim keyTitle As PROPERTYKEY
                Dim guidTitle As GUID
                Call IIDFromString(StrPtr("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}"), guidTitle)
                keyTitle.fmtid = guidTitle
                keyTitle.pid = 2 ' PKEY_Title
                
                Dim propvar As PROPVARIANT
                hr = InitPropVariantAsLPWSTR(CStr(item("表示名")), propvar)
                Debug.Print "  InitPropVariantAsLPWSTR result: 0x" & Hex(hr)
                If hr = S_OK Then
                    hr = SetPropertyValue(pPropStore, keyTitle, propvar)
                    Debug.Print "  SetPropertyValue result: 0x" & Hex(hr)
                    If hr = S_OK Then
                        hr = CommitPropertyStore(pPropStore)
                        Debug.Print "  CommitPropertyStore result: 0x" & Hex(hr)
                    End If
                    Call PropVariantClear(propvar)
                End If
                Call ComRelease(pPropStore)
            End If
        Else
            Debug.Print "  IIDFromString(IPropertyStore) failed: 0x" & Hex(hr)
        End If
    End If
    
    CreateShellLink = pLink
End Function


'***************************************************************************************************
'                       ■■■ 個別 COM メソッド薄ラッパー関数群 ■■■
'***************************************************************************************************

' ICustomDestinationList::SetAppID
Private Function SetAppID(ByVal pDestList As LongPtr, ByVal pszAppID As String) As Long
    Dim vArgs(0 To 0) As Variant
    Dim vTypes(0 To 0) As Integer
    Dim vPtrs(0 To 0) As LongPtr
    
    vArgs(0) = StrPtr(pszAppID)
    vTypes(0) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    
    SetAppID = InvokeComMethod(pDestList, VTBL_SETAPPID, 1, vTypes, vPtrs)
End Function

' ICustomDestinationList::BeginList
Private Function BeginList(ByVal pDestList As LongPtr, ByRef pcMaxSlots As Long, ByRef riid As GUID, ByRef ppv As LongPtr) As Long
    Dim vArgs(0 To 2) As Variant
    Dim vTypes(0 To 2) As Integer
    Dim vPtrs(0 To 2) As LongPtr
    
    vArgs(0) = VarPtr(pcMaxSlots)
    vArgs(1) = VarPtr(riid)
    vArgs(2) = VarPtr(ppv)
    vTypes(0) = VT_PARAM_PTR
    vTypes(1) = VT_PARAM_PTR
    vTypes(2) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    vPtrs(1) = VarPtr(vArgs(1))
    vPtrs(2) = VarPtr(vArgs(2))
    
    BeginList = InvokeComMethod(pDestList, VTBL_BEGINLIST, 3, vTypes, vPtrs)
End Function

' ICustomDestinationList::AppendCategory
Private Function AppendCategoryToDestList(ByVal pDestList As LongPtr, ByVal pszCategory As String, ByVal poa As LongPtr) As Long
    Dim vArgs(0 To 1) As Variant
    Dim vTypes(0 To 1) As Integer
    Dim vPtrs(0 To 1) As LongPtr
    
    vArgs(0) = StrPtr(pszCategory)
    vArgs(1) = poa
    vTypes(0) = VT_PARAM_PTR
    vTypes(1) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    vPtrs(1) = VarPtr(vArgs(1))
    
    AppendCategoryToDestList = InvokeComMethod(pDestList, VTBL_APPENDCATEGORY, 2, vTypes, vPtrs)
End Function

' ICustomDestinationList::AddUserTasks
Private Function AddUserTasksToDestList(ByVal pDestList As LongPtr, ByVal poa As LongPtr) As Long
    Dim vArgs(0 To 0) As Variant
    Dim vTypes(0 To 0) As Integer
    Dim vPtrs(0 To 0) As LongPtr
    
    vArgs(0) = poa
    vTypes(0) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    
    AddUserTasksToDestList = InvokeComMethod(pDestList, VTBL_ADDUSERTASKS, 1, vTypes, vPtrs)
End Function

' ICustomDestinationList::CommitList
Private Function CommitList(ByVal pDestList As LongPtr) As Long
    Dim vTypes(0 To 0) As Integer
    Dim vPtrs(0 To 0) As LongPtr
    CommitList = InvokeComMethod(pDestList, VTBL_COMMITLIST, 0, vTypes, vPtrs)
End Function

' ICustomDestinationList::DeleteList
Private Function DeleteList(ByVal pDestList As LongPtr, ByVal pszAppID As String) As Long
    Dim vArgs(0 To 0) As Variant
    Dim vTypes(0 To 0) As Integer
    Dim vPtrs(0 To 0) As LongPtr
    
    vArgs(0) = StrPtr(pszAppID)
    vTypes(0) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    
    DeleteList = InvokeComMethod(pDestList, VTBL_DELETELIST, 1, vTypes, vPtrs)
End Function

' IObjectCollection::AddObject
Private Function AddObjectToCollection(ByVal pCollection As LongPtr, ByVal punk As LongPtr) As Long
    Dim vArgs(0 To 0) As Variant
    Dim vTypes(0 To 0) As Integer
    Dim vPtrs(0 To 0) As LongPtr
    
    vArgs(0) = punk
    vTypes(0) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    
    AddObjectToCollection = InvokeComMethod(pCollection, VTBL_ADDOBJECT, 1, vTypes, vPtrs)
End Function

' IShellLinkW::SetPath
Private Function SetShellLinkPath(ByVal pLink As LongPtr, ByVal pszFile As String) As Long
    Dim vArgs(0 To 0) As Variant
    Dim vTypes(0 To 0) As Integer
    Dim vPtrs(0 To 0) As LongPtr
    
    vArgs(0) = StrPtr(pszFile)
    vTypes(0) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    
    SetShellLinkPath = InvokeComMethod(pLink, VTBL_SETPATH, 1, vTypes, vPtrs)
End Function

' IShellLinkW::SetArguments
Private Function SetShellLinkArguments(ByVal pLink As LongPtr, ByVal pszArgs As String) As Long
    Dim vArgs(0 To 0) As Variant
    Dim vTypes(0 To 0) As Integer
    Dim vPtrs(0 To 0) As LongPtr
    
    vArgs(0) = StrPtr(pszArgs)
    vTypes(0) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    
    SetShellLinkArguments = InvokeComMethod(pLink, VTBL_SETARGUMENTS, 1, vTypes, vPtrs)
End Function

' IShellLinkW::SetDescription
Private Function SetShellLinkDescription(ByVal pLink As LongPtr, ByVal pszName As String) As Long
    Dim vArgs(0 To 0) As Variant
    Dim vTypes(0 To 0) As Integer
    Dim vPtrs(0 To 0) As LongPtr
    
    vArgs(0) = StrPtr(pszName)
    vTypes(0) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    
    SetShellLinkDescription = InvokeComMethod(pLink, VTBL_SETDESCRIPTION, 1, vTypes, vPtrs)
End Function

' IShellLinkW::SetIconLocation
Private Function SetShellLinkIconLocation(ByVal pLink As LongPtr, ByVal pszIconPath As String, ByVal iIcon As Long) As Long
    Dim vArgs(0 To 1) As Variant
    Dim vTypes(0 To 1) As Integer
    Dim vPtrs(0 To 1) As LongPtr
    
    vArgs(0) = StrPtr(pszIconPath)
    vArgs(1) = iIcon
    vTypes(0) = VT_PARAM_PTR
    vTypes(1) = vbLong
    vPtrs(0) = VarPtr(vArgs(0))
    vPtrs(1) = VarPtr(vArgs(1))
    
    SetShellLinkIconLocation = InvokeComMethod(pLink, VTBL_SETICONLOCATION, 2, vTypes, vPtrs)
End Function

' IPropertyStore::SetValue
Private Function SetPropertyValue(ByVal pPropStore As LongPtr, ByRef key As PROPERTYKEY, ByRef propvar As PROPVARIANT) As Long
    Dim vArgs(0 To 1) As Variant
    Dim vTypes(0 To 1) As Integer
    Dim vPtrs(0 To 1) As LongPtr
    
    vArgs(0) = VarPtr(key)
    vArgs(1) = VarPtr(propvar)
    vTypes(0) = VT_PARAM_PTR
    vTypes(1) = VT_PARAM_PTR
    vPtrs(0) = VarPtr(vArgs(0))
    vPtrs(1) = VarPtr(vArgs(1))
    
    SetPropertyValue = InvokeComMethod(pPropStore, VTBL_SETVALUE, 2, vTypes, vPtrs)
End Function

' IPropertyStore::Commit
Private Function CommitPropertyStore(ByVal pPropStore As LongPtr) As Long
    Dim vTypes(0 To 0) As Integer
    Dim vPtrs(0 To 0) As LongPtr
    CommitPropertyStore = InvokeComMethod(pPropStore, VTBL_COMMIT, 0, vTypes, vPtrs)
End Function



'***************************************************************************************************
'                               ■■■ Demo用プロシージャ ■■■
'***************************************************************************************************
Sub demo_JumplistControl()
    ' 一旦クリア
    Call Clear
    
    ' タスク登録
    Call Registration("メモ帳を起動", "notepad.exe", "", "", "Windows標準のメモ帳を起動します", "shell32.dll", 2)
    Call Registration("電卓を起動", "calc.exe", "", "", "Windows標準の電卓を起動します", "shell32.dll", 23)
    
    ' カスタムカテゴリ登録
    Call Registration("開発ドキュメント", "notepad.exe", "README.md", "お気に入り", "READMEファイルを開きます", "shell32.dll", 22)
    Call Registration("Excelヘルプ", "https://learn.microsoft.com/ja-jp/", "", "お気に入り", "Microsoft学習サイトを開きます", "shell32.dll", 14)
    
    ' ジャンプリストへ反映
    Call Import
    
    MsgBox "ジャンプリストの登録が完了しました！" & vbCrLf & _
           "タスクバーの Excel アイコンを右クリックして確認してください。", vbInformation, "ジャンプリスト登録完了"
End Sub

Sub demo_JumplistClear()
    Call Clear
    MsgBox "ジャンプリストをクリアしました！", vbInformation, "ジャンプリストクリア完了"
End Sub
