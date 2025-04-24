//設定がまとまってるヘッダーファイルを指定
#include "TaskbarProgress.h" 

//よく使う名前定義を用意する
using namespace winrt;
using namespace winrt::Windows::UI::Notifications;
using namespace winrt::Windows::Data::Xml::Dom;
using namespace Gdiplus;



//***************************************************************************************************
//                           ■■■ ThumbButtonInfo クラス を扱う準備 ■■■
//***************************************************************************************************
#define MAX_BUTTONS 7                                                   //配置可能なボタンの上限数
#define ButtonID_Correction 1001                                        //ボタンIDの採番開始番号

static ITaskbarList3* g_taskbar = nullptr;                              //ITaskbarList3オブジェクト
static THUMBBUTTON g_btns[MAX_BUTTONS] = {};                            //ボタン情報格納用
static std::wstring g_procNames[MAX_BUTTONS];                           //コールバック用プロシージャ名の格納用
static HWND g_hwnd = nullptr;                                           //InitializeThumbnailButton で呼び出したウィンドウハンドルを保持します
static VbaCallback g_thumbButtonCallbacks[MAX_BUTTONS] = { nullptr };   // 各ボタン用の関数(VBA内プロシージャ)ポインタを7つ保持



//***************************************************************************************************
//                                 ■■■ 内部のヘルパー関数 ■■■
//***************************************************************************************************
//* 機能　　 ：内部でバッジ名に変換する関数
//---------------------------------------------------------------------------------------------------
//* 引数　　 ：badgeValue    任意の整数
//* 返り値　 ：引数の数値に応じたバッジ名
//---------------------------------------------------------------------------------------------------
//* URL      ：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/badges
//***************************************************************************************************
static std::wstring GetBadgeValueString(int badgeValue)
{
    switch (badgeValue) {
    case -1:  return L"activity";
    case -2:  return L"alert";
    case -3:  return L"alarm";
    case -4:  return L"available";
    case -5:  return L"away";
    case -6:  return L"busy";
    case -7:  return L"newMessage";
    case -8:  return L"paused";
    case -9:  return L"playing";
    case -10: return L"unavailable";
    case -11: return L"error";
    case -12: return L"attention";
    default:
        if (badgeValue <= -13) return L"none";
        else return std::to_wstring(badgeValue);
    }
}

//***************************************************************************************************
//* 機能　　 ：タスクバーのボタンUI準備ヘルパー
//***************************************************************************************************
void EnsureTaskbarInterface()
{
    if (!g_taskbar) {
        CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
        CoCreateInstance(CLSID_TaskbarList, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&g_taskbar));
        if (g_taskbar) g_taskbar->HrInit();
    }
}

//***************************************************************************************************
//* 機能　　 ：事前に設定した hwnd に起きたことが、全部ここに届きます。
//---------------------------------------------------------------------------------------------------
//* 引数　 　：hwnd      メッセージを受け取ったウィンドウのハンドル(サブクラスに登録したhwnd)
//             msg       メッセージの種類。Excelで言う、イベントの種類です（例：WM_COMMAND, WM_PAINT, WM_CLOSE など）
//             wParam    メッセージによって意味が異なる補助データ　その1
//             lParam    メッセージによって意味が異なる補助データ　その2
//---------------------------------------------------------------------------------------------------
//* 機能説明 ：Excelで言う、全イベント処理がここに集約されてるイメージです。イベントごとの処理は、Switch文がやりやすいです。
//***************************************************************************************************
LRESULT CALLBACK SubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    //switch文で、イベントごとに「やりたい処理」を書く
    switch (msg)
    {
        //タスクバーのサムネイルボタンをクリックすると、Windows は WM_COMMAND メッセージを送ってきます。
        case WM_COMMAND:
            //このイベントはボタンがクリックされた通知か判定します(今回は THBN_CLICKED となる)
            if (HIWORD(wParam) == THBN_CLICKED) {
                //補正処理
                int buttonIndex = LOWORD(wParam) - ButtonID_Correction;

                //VBA内のプロシージャ名のポインタを直接実行する
                g_thumbButtonCallbacks[buttonIndex]();
                return 0;
            }

            break;

        //他のイベントは、何もしません
        default:
            break;
    }

    //他のイベントは、既定の処理へ
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

//***************************************************************************************************
//* 機能　　 ：Win32アプリ用、通知バッチアイコンを作成します
//---------------------------------------------------------------------------------------------------
//* 引数　 　：number    描画する数字。99を超える場合、"99+"と表示します
//---------------------------------------------------------------------------------------------------
//* 詳細説明 ：
//***************************************************************************************************
HICON CreateBadgeIcon(int number)
{
    // GDI+ 初期化
    static bool initialized = false;
    static GdiplusStartupInput gdiplusStartupInput;
    static ULONG_PTR gdiplusToken;
    if (!initialized) {
        GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, nullptr);
        initialized = true;
    }

    // 描画サイズ（アイコンサイズ）
    const int size = 32;
    Bitmap bmp(size, size, PixelFormat32bppARGB);
    Graphics g(&bmp);

    g.SetSmoothingMode(SmoothingModeAntiAlias);
    g.Clear(Color(0, 0, 0, 0)); // 透明背景

    // 赤い円
    SolidBrush redBrush(Color(255, 255, 0, 0)); // 赤
    g.FillEllipse(&redBrush, 0, 0, size, size);

    // 白文字
    WCHAR buf[4];
    if (number <= 99)
        wsprintf(buf, L"%d", number);
    else
        lstrcpyW(buf, L"99+");

    FontFamily fontFamily(L"Segoe UI");
    Gdiplus::Font font(&fontFamily, 14, FontStyleBold, UnitPixel);
    SolidBrush whiteBrush(Color(255, 255, 255)); // 白

    RectF layoutRect(0, 0, size, size);
    StringFormat format;
    format.SetAlignment(StringAlignmentCenter);
    format.SetLineAlignment(StringAlignmentCenter);

    g.DrawString(buf, -1, &font, layoutRect, &format, &whiteBrush);

    // アイコンに変換
    HICON hIcon = nullptr;
    bmp.GetHICON(&hIcon);
    return hIcon;
}



//***************************************************************************************************
//                                 ■■■ VBAから使用する関数 ■■■
//***************************************************************************************************
//* 機能　　 ：指定アプリハンドルのタスクバーに、プログレスバーの値とステータスを指定します。
//---------------------------------------------------------------------------------------------------
//* 引数　 　：・hwnd     タスクバーを適用させるハンドル
//             ・current  現在値
//             ・maximum  最大値
//             ・status　 数値によって、色変更、不確定にできます。
//***************************************************************************************************
void __stdcall SetTaskbarProgress(HWND hwnd, unsigned long current, unsigned long maximum, long status)
{
	// ITaskbarList3インターフェースを取得
	ITaskbarList3* pTaskbarList = nullptr;
	HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pTaskbarList));
	if (FAILED(hr)) {
		return;
	}

	// タスクバーの進捗状態を設定
	pTaskbarList->SetProgressState(hwnd, static_cast<TBPFLAG>(status));

	// 進捗値を設定
	if (status == TBPF_NORMAL || status == TBPF_PAUSED || status == TBPF_ERROR) {
		pTaskbarList->SetProgressValue(hwnd, current, maximum);
	}

	// リソースの解放
	pTaskbarList->Release();

}

//***************************************************************************************************
//* 機能　　 ：指定アプリハンドルのタスク バー ボタンにオーバーレイを適用して、アプリケーションの状態または通知をユーザーに示します。
//---------------------------------------------------------------------------------------------------
//* 引数　 　：・hwnd        タスクバーを適用させるハンドル
//             ・filePath    アイコンを含むファイルフルパス
//             ・iconIndex   dll等のアイコンセットを読み込んだ際の、読み込み位置
//             ・description アクセシビリティ向け説明文
//---------------------------------------------------------------------------------------------------
//* 機能説明 ：オーバーレイの削除には、iconIndex を -1 以下にします。
//***************************************************************************************************
void __stdcall SetTaskbarOverlayIcon(HWND hwnd, const wchar_t* filePath, int iconIndex, const wchar_t* description)
{
    // ITaskbarList3インターフェースを取得
    ITaskbarList3* pTaskbarList = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pTaskbarList));
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to create ITaskbarList3 instance.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
        return;
    }

    // iconIndexが0未満の場合、アイコンを削除する
    if (iconIndex < 0) {
        hr = pTaskbarList->SetOverlayIcon(hwnd, NULL, NULL);
        if (FAILED(hr)) {
            MessageBoxW(nullptr, L"FFailed to remove overlay icon.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
        }
        pTaskbarList->Release();
        return;
    }

    HICON hIcon = NULL;
    std::wstring path(filePath);
    std::wstring extension = path.substr(path.find_last_of(L".") + 1);

    if (extension == L"ico") {
        // .icoファイルからアイコンをロード
        hIcon = (HICON)LoadImage(NULL, filePath, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE | LR_SHARED);
        if (hIcon == NULL) {
            MessageBoxW(nullptr, L"Failed to load .ico file.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
            pTaskbarList->Release();
            return;
        }
    }
    else if (extension == L"exe") {
        // .exeファイルからアイコンをインデックス指定でロード
        hIcon = ExtractIcon(NULL, filePath, iconIndex);
        if (hIcon == NULL || hIcon == (HICON)1) {
            MessageBoxW(nullptr, L"Failed to extract icon from .exe file.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
            pTaskbarList->Release();
            return;
        }
    }
    else if (extension == L"dll") {
        // .dllファイルからアイコンをインデックス指定でロード
        HMODULE hModule = LoadLibraryEx(filePath, NULL, LOAD_LIBRARY_AS_DATAFILE);
        if (hModule == NULL) {
            MessageBoxW(nullptr, L"Failed to load .dll file.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
            pTaskbarList->Release();
            return;
        }

        hIcon = (HICON)LoadImage(hModule, MAKEINTRESOURCE(iconIndex), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
        if (hIcon == NULL) {
            MessageBoxW(nullptr, L"Failed to load icon from resource.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
            FreeLibrary(hModule);
            pTaskbarList->Release();
            return;
        }

        FreeLibrary(hModule);
    }
    else {
        MessageBoxW(nullptr, L"Unsupported file type.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
        pTaskbarList->Release();
        return;
    }

    // タスクバーにオーバーレイアイコンを設定
    hr = pTaskbarList->SetOverlayIcon(hwnd, hIcon, description);
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to set overlay icon.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
    }

    // アイコンを解放
    DestroyIcon(hIcon);

    // リソースの解放
    pTaskbarList->Release();
}

//***************************************************************************************************
//* 機能　　 ：指定アプリIDのタスク バー ボタンにオーバーレイを適用して、アプリケーションの状態または通知をユーザーに示します
//---------------------------------------------------------------------------------------------------
//* 引数　 　：・badgeValue        タスクバーを適用させるハンドル
//             ・appUserModelID    appUserModelID
//---------------------------------------------------------------------------------------------------
//* 機能説明 ：アプリハンドルではなく、appUserModelID で指定するパターンです。
//* 注意事項 ：・WinRT API環境のあるOSが必要です
//             ・現時点では、デスクトップアプリに対しては効果ありません。
//***************************************************************************************************
void __stdcall SetTaskbarOverlayBadge(int badgeValue, const wchar_t* appUserModelID)
{
    // COMの初期化
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (hr == RPC_E_CHANGED_MODE) {
        // 既に異なるアパートメント モードで初期化されている場合は、そのまま続行
    }
    else if (FAILED(hr)) {
        wchar_t errorMsg[256];
        swprintf_s(errorMsg, 256, L"COM初期化に失敗しました。HRESULT: 0x%08X", hr);
        MessageBoxW(nullptr, errorMsg, L"エラー", MB_OK | MB_ICONERROR);
        return;
    }

    try {
        // バッジの値を文字列に変換
        std::wstring badgeValueStr = GetBadgeValueString(badgeValue);
        std::wstring xmlString = L"<badge value=\"" + badgeValueStr + L"\"/>";

        // XMLの読み込み
        XmlDocument doc;
        doc.LoadXml(winrt::hstring(xmlString));

        // バッジ通知オブジェクトの作成
        BadgeNotification badge(doc);

        // 指定したAppIDの通知マネージャを取得
        auto notifier = BadgeUpdateManager::CreateBadgeUpdaterForApplication(winrt::hstring(appUserModelID));
        notifier.Update(badge);
    }
    catch (...) {
        // エラー処理
        MessageBoxW(nullptr, L"バッジ通知の表示に失敗しました。", L"Badge Error", MB_OK | MB_ICONERROR);
    }
}

//***************************************************************************************************
//* 機能　　 ：指定アプリハンドルにオーバーレイを適用して、アプリケーションの状態または通知をユーザーに示します
//---------------------------------------------------------------------------------------------------
//* 引数　 　：・badgeValue        タスクバーを適用させるバッチ
//             ・hwnd              ウィンドウハンドル
//---------------------------------------------------------------------------------------------------
//* 詳細説明 ：SetTaskbarOverlayBadge を win32 アプリでも扱えるようにしたものです
//* 注意事項 ：仕組みは、SetOverlayIcon + GDI+ でメモリ上にアイコンを描画し、HICONを生成 で実現してます 
//***************************************************************************************************
void __stdcall SetTaskbarOverlayBadgeForWin32(LONG badgeValue, HWND hwnd)
{
    // ITaskbarList3インターフェースを取得
    ITaskbarList3* pTaskbarList = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pTaskbarList));
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to create ITaskbarList3 instance.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
        return;
    }

    // iconIndexが0以下の場合、アイコンを削除する
    if (badgeValue <= 0) {
        hr = pTaskbarList->SetOverlayIcon(hwnd, NULL, NULL);
        if (FAILED(hr)) {
            MessageBoxW(nullptr, L"FFailed to remove overlay icon.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
        }
        pTaskbarList->Release();
        return;
    }

    //反映
    HICON icon = CreateBadgeIcon(badgeValue);
    pTaskbarList->SetOverlayIcon(hwnd, icon,NULL);
}

//***************************************************************************************************
//* 機能　　 ： 指定したウィンドウハンドルにボタン情報を確保します。
//---------------------------------------------------------------------------------------------------
//* 引数　 　： hwnd            ウィンドウハンドル
//---------------------------------------------------------------------------------------------------
//* 詳細説明 ： 上限である7個の非表示ボタンを作ります。以降の設定変更等は、「UpdateThumbnailButton」で
//* 注意事項 ： ・非表示として確保するので、この処理だけでは見た目上、何も起こりません
//              ・実行中のウィンドウハンドルで、1回のみ呼び出すこと。複数の呼び出しは、予期せぬ挙動を招きます。
//***************************************************************************************************
void __stdcall InitializeThumbnailButton(HWND hwnd) 
{
    //初期化処理
    EnsureTaskbarInterface();

    //非表示として、ボタン情報を確保する
    for (int i = 0; i < MAX_BUTTONS; ++i) {
        g_btns[i].dwMask = THB_FLAGS;
        g_btns[i].dwFlags = THBF_HIDDEN;
        g_btns[i].iId = i + ButtonID_Correction;
        g_btns[i].hIcon = NULL;
        g_btns[i].szTip[0] = L'\0';
    }

    //反映処理
    g_taskbar->ThumbBarAddButtons(hwnd, MAX_BUTTONS, g_btns);

    // HWND を保持
    g_hwnd = hwnd;

    // 対象のウィンドウハンドル(hwnd)をサブクラス化して、様々なイベント処理に対応させる
    SetWindowSubclass(hwnd, SubclassProc, 1, 0);
}

//***************************************************************************************************
//* 機能　　 ： 指定したウィンドウハンドルにボタン情報を変更します。
//---------------------------------------------------------------------------------------------------
//* 引数　 　： data     ユーザー定義型：THUMBBUTTONDATA
//              callback 呼び出すVBA内のプロシージャ名のポインタ
//***************************************************************************************************
void __stdcall UpdateThumbnailButton(const THUMBBUTTONDATA* data, VbaCallback callback)
{
    //初期化
    EnsureTaskbarInterface();

    //範囲外のボタンIDなら、何もしない
    if (!data || data->ButtonIndex  < 0 + ButtonID_Correction || data->ButtonIndex  >= MAX_BUTTONS + ButtonID_Correction) return;

    //指定ボタンIDに対して、どんな有効なデータが含まれているか伝える
    THUMBBUTTON& btn = g_btns[data->ButtonIndex - ButtonID_Correction];
    btn.iId = data->ButtonIndex;                        //ツール バー内で一意のボタンのアプリケーション定義識別子。念の為、1001から刻む
    btn.dwMask = THB_FLAGS | THB_ICON | THB_TOOLTIP;    //メンバーに有効なデータが含まれているかを指定する THUMBBUTTONMASK 値の組み合わせ。https://learn.microsoft.com/ja-jp/windows/win32/api/shobjidl_core/ne-shobjidl_core-thumbbuttonmask
    btn.dwFlags = (THUMBBUTTONFLAGS)data->ButtonType;   //THUMBBUTTON によって、ボタンの特定の状態と動作を制御する

    // ツールチップ
    if (data->Description) {
        wcsncpy_s(btn.szTip, data->Description, ARRAYSIZE(btn.szTip));
    }

    // アイコン
    HICON hIcon = NULL;
    if (data->IconPath) {
        ExtractIconExW(data->IconPath, data->IconIndex, NULL, &hIcon, 1);
    }
    btn.hIcon = hIcon;

    // コールバック用にプロシージャ名のポインタを保持
    g_thumbButtonCallbacks[data->ButtonIndex - ButtonID_Correction] = callback;

    //変更を適用
    g_taskbar->ThumbBarUpdateButtons(g_hwnd, MAX_BUTTONS, g_btns);
}

//***************************************************************************************************
//* 機能　　 ： ジャンプリスト制御に使った変数を解放します
//---------------------------------------------------------------------------------------------------
//* 引数　 　： ※割愛します
//***************************************************************************************************
void Cleanup_Jumplist(ICustomDestinationList* pDestList, IObjectCollection* pTasks, IShellLinkW* pLink){
    if (pLink) pLink->Release();
    if (pTasks) pTasks->Release();
    if (pDestList) pDestList->Release();
    CoUninitialize();
}

//***************************************************************************************************
//* 機能　　 ： 高度なジャンプリストを作成します
//---------------------------------------------------------------------------------------------------
//* 引数　 　： RegistrationData     ユーザー定義型：JumpListData
//***************************************************************************************************
void __stdcall Registration_Jumplist(const JumpListData* RegistrationData)
{
    //必要な変数を用意→初期化
    HRESULT hr;
    ICustomDestinationList* pDestList = nullptr;
    IObjectCollection* pTasks = nullptr;
    IShellLinkW* pLink = nullptr;

    //-------------ジャンプリスト関連COMオブジェクトの準備に関わるお作法-------------
    hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) return;

    hr = CoCreateInstance(CLSID_DestinationList, nullptr, CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&pDestList));
    if (FAILED(hr)) { Cleanup_Jumplist(pDestList, pTasks, pLink); return; }
    //-------------------------------------------------------------------------------

    //ジャンプリストの設定先の ApplicationModelUserID の設定先を反映
    hr = pDestList->SetAppID(RegistrationData->ApplicationModelUserID);
    if (FAILED(hr)) { Cleanup_Jumplist(pDestList, pTasks, pLink); return; }

    //ジャンプリスト編集のセッションを開始
    UINT cMinSlots;
    IObjectArray* poaRemoved;
    hr = pDestList->BeginList(&cMinSlots, IID_PPV_ARGS(&poaRemoved));
    if (FAILED(hr)) { Cleanup_Jumplist(pDestList, pTasks, pLink); return; }

    //ジャンプリスト登録データ用オブジェクトを用意
    hr = CoCreateInstance(CLSID_EnumerableObjectCollection, nullptr, CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&pTasks));
    if (FAILED(hr)) { Cleanup_Jumplist(pDestList, pTasks, pLink); return; }

    //タスクリンク作成。
    hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&pLink));
    if (FAILED(hr)) { Cleanup_Jumplist(pDestList, pTasks, pLink); return; }
    
    //作成したタスクに対して、パラメーターを設定
    pLink->SetPath(RegistrationData->FilePath);                                        //実行パス
    pLink->SetArguments(RegistrationData->cmdArguments);                               //引数
    pLink->SetIconLocation(RegistrationData->iconPath, RegistrationData->IconIndex);   //アイコン設定
    pLink->SetDescription(RegistrationData->Description);                              //アクセシビリティ用説明文

    //ジャンプリストに、追加のメタデータ付与制御(ピン留め出来ないようにする等)
    IPropertyStore* pPropStore;
    hr = pLink->QueryInterface(IID_PPV_ARGS(&pPropStore));
    if (SUCCEEDED(hr)) {
        //-----------------------PROPVARIANT の 設定値を生成------------------------
        //BOOL：TRUE
        PROPVARIANT varBoolTrue;
        PropVariantInit(&varBoolTrue);
        varBoolTrue.vt = VT_BOOL;
        varBoolTrue.boolVal = VARIANT_TRUE;

        //BOOL：FALSE
        PROPVARIANT varBoolFalse;
        PropVariantInit(&varBoolFalse);
        varBoolFalse.vt = VT_BOOL;
        varBoolFalse.boolVal = VARIANT_FALSE;

        //String：タスク名に該当
        PROPVARIANT varTitle;
        InitPropVariantFromString(RegistrationData->taskName, &varTitle);
        //--------------------------------------------------------------------------

        //------------------------メタデータを設定/適用-----------------------------
        //URL　https://learn.microsoft.com/ja-jp/windows/win32/properties/software-bumper
        //pPropStore->SetValue(PKEY_AppUserModel_PreventPinning, varBoolTrue);    //ピン留め、一覧から削除　を効かなくします
        pPropStore->SetValue(PKEY_Title, varTitle);                             //タスク名を設定します
    
        //適用
        pPropStore->Commit();
        //--------------------------------------------------------------------------

        //変数、オブジェクトを解放
        PropVariantClear(&varTitle);
        PropVariantClear(&varBoolTrue);
        pPropStore->Release();
    }

    //指定した「ApplicationModelUserID」(pTasks)に、設定したタスク(pLink->XXX)を入れる。
    pTasks->AddObject(pLink);

    //指定したカテゴリ名称群に、上記で定義した pTasks を入れる 
    if (RegistrationData->categoryName == nullptr || wcslen(RegistrationData->categoryName) == 0) {
        // カテゴリ名が未指定 → Tasks に追加（ピン留めできない）
        pDestList->AddUserTasks(pTasks);
    }
    else {
        // カテゴリ名が指定されている → 任意カテゴリ名で追加（ピン留め可能性あり）
        pDestList->AppendCategory(RegistrationData->categoryName, pTasks);
    }

    //ジャンプリスト登録
    pDestList->CommitList();
}
