//設定がまとまってるヘッダーファイルを指定
#include "TaskbarProgress.h" 

//よく使う名前定義を用意する
using namespace winrt;
using namespace winrt::Windows::UI::Notifications;
using namespace winrt::Windows::Data::Xml::Dom;



//***************************************************************************************************
//                           ■■■ ThumbButtonInfo クラス を扱う準備 ■■■
//***************************************************************************************************
#define MAX_BUTTONS 7                           //配置可能なボタンの上限数
#define ButtonID_Correction 1001                //ボタンIDの採番開始番号

static ITaskbarList3* g_taskbar = nullptr;      //ITaskbarList3オブジェクト
static THUMBBUTTON g_btns[MAX_BUTTONS] = {};    //ボタン情報格納用
static std::wstring g_procNames[MAX_BUTTONS];   //コールバック用プロシージャ名の格納用



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
void EnsureTaskbarInterface() {
    if (!g_taskbar) {
        CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
        CoCreateInstance(CLSID_TaskbarList, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&g_taskbar));
        if (g_taskbar) g_taskbar->HrInit();
    }
}

//***************************************************************************************************
//* 機能　　 ：引数にあるプロシージャ名で、VBA マクロを実行します
//---------------------------------------------------------------------------------------------------
//* 引数　 　：Index     プロシージャ名があるIndex値
//***************************************************************************************************
void ExecuteVBAProcByIndex(int index) {
    //プロシージャ名未登録あるいは、インデックスの範囲外なら、ここで終了
    if (index < 0 || index >= 7 || g_procNames[index].empty()) return;

    //詳細メッセージ、取得用(For Debug)
    EXCEPINFO excepInfo;
    memset(&excepInfo, 0, sizeof(EXCEPINFO));  // 初期化

    // 1. ExcelのCLSIDを取得
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    // 恐らく、Excelがインストールされてない場合
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get CLSID for Excel", L"Error", MB_OK);
        return;
    }

    // 2. 既存のExcelインスタンスを取得
    IDispatch* pExcelApp = nullptr;
    hr = GetActiveObject(clsid, nullptr, (IUnknown**)&pExcelApp);
    // 起動中のExcelがない場合
    if (FAILED(hr) || !pExcelApp) {
        MessageBoxW(nullptr, L"Failed to get active Excel instance", L"Error", MB_OK);

        CoUninitialize();
        return;
    }

    // 3. 「 Run メソッド」のDISPIDの取得
    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"Run");  // 実行するメソッド名(VBAのApplication.Run 相当)
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    //Runメソッドの取得に失敗した場合
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get DISPID for Run method", L"Error", MB_OK);
        return;
    }

    // 4. Application.Run メソッドの引数を設定。
    CComVariant macroName(g_procNames[index].c_str());  //実行したいマクロ(プロシージャ)名
    //　初期化
    DISPPARAMS params = {};
    VARIANTARG arg;
    VariantInit(&arg);
    //　実行マクロ名を設定
    _bstr_t procName(macroName);
    //　パラメーターの仕様を定義
    arg.vt = VT_BSTR;
    arg.bstrVal = procName;
    params.rgvarg = &arg;
    params.cArgs = 1;

    // 5. マクロの呼び出し
    CComVariant result;
    hr = pExcelApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &result, &excepInfo, nullptr);

    //-------------以降は、デバッグ用-------------
    // 現在のExcelインスタンス内に、指定マクロがないと想定
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get Excel macro", L"Error", MB_OK);
    }

    //MessageBoxでDISPPARAMSの内容を確認
    std::wstring debugMessage;

    // cArgsの確認
    debugMessage += L"Number of arguments: " + std::to_wstring(params.cArgs) + L"\n";

    // rgvarg の中身を文字列化
    for (UINT i = 0; i < params.cArgs; ++i) {
        VARIANT& arg = params.rgvarg[i];

        if (arg.vt == VT_BSTR) {
            debugMessage += L"Argument " + std::to_wstring(i) + L": " + arg.bstrVal + L"\n";
        }
        else {
            debugMessage += L"Argument " + std::to_wstring(i) + L": [not a BSTR]\n";
        }
    }

    // rgvarg の中身を確認
    MessageBoxW(nullptr, debugMessage.c_str(), L"DISPPARAMS Debug", MB_OK);

     //エラーが起こったら、エラーコードと詳細メッセージ(ある場合)を表示。
    if (FAILED(hr)) {
        std::wstring errorMessage = L"Invoke failed. HRESULT: " + std::to_wstring(hr);

        if (excepInfo.bstrDescription) {
            errorMessage += L"\nException: " + std::wstring(excepInfo.bstrDescription);
            SysFreeString(excepInfo.bstrDescription);  // リソース解放
        }

        MessageBoxW(nullptr, errorMessage.c_str(), L"Error1", MB_OK);
    }
    else {
        _com_error err(hr);
        MessageBoxW(nullptr, err.ErrorMessage(), L"Info", MB_OK);
    }

    //-------------ここまでが、デバッグ用-------------

    //後始末
    pExcelApp->Release();
    CoUninitialize();
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

                //VBA内のプロシージャ名を実行する準備へ
                ExecuteVBAProcByIndex(buttonIndex);
                return 0;
            }

            break;

        //他のイベントは、何もしません
        default:
            break;
    }

    //他のイベントは、既定の処理へ
    return DefWindowProc(hwnd, msg, wParam, lParam);
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
//* 機能　　 ： 指定したウィンドウハンドルにボタン情報を確保します。
//---------------------------------------------------------------------------------------------------
//* 引数　 　： buttonCount     確保するボタン数
//              hwnd            ウィンドウハンドル
//---------------------------------------------------------------------------------------------------
//* 注意事項 ： 非表示として確保するので、この処理だけでは見た目上、何も起こりません
//***************************************************************************************************
void __stdcall InitializeThumbnailButton(LONG buttonCount, HWND hwnd) {
    //初期化処理
    EnsureTaskbarInterface();

    //0以下で渡されたら、ボタン自体を削除し、処理終了
    if (buttonCount <= 0) {
        memset(g_btns, 0, sizeof(g_btns));
        g_taskbar->ThumbBarAddButtons(hwnd, 0, nullptr);

        //サブクラス化、解除
        RemoveWindowSubclass;
        return;
    }

    //上限を超えてたら、何もしない
    if (buttonCount > MAX_BUTTONS) return;

    //非表示として、ボタン情報を確保する
    for (int i = 0; i < MAX_BUTTONS; ++i) {
        g_btns[i].dwMask = THB_FLAGS;
        g_btns[i].dwFlags = THBF_HIDDEN;
        g_btns[i].iId = i + ButtonID_Correction;
        g_btns[i].hIcon = NULL;
        g_btns[i].szTip[0] = L'\0';
    }

    //反映処理
    g_taskbar->ThumbBarAddButtons(hwnd, buttonCount, g_btns);

    // 対象のウィンドウハンドル(hwnd)をサブクラス化して、様々なイベント処理に対応させる
    SetWindowSubclass(hwnd, SubclassProc, 1, 0);
}

//***************************************************************************************************
//* 機能　　 ： 指定したウィンドウハンドルにボタン情報を変更します。
//---------------------------------------------------------------------------------------------------
//* 引数　 　： data     ユーザー定義型：THUMBBUTTONDATA
//              hwnd     ウィンドウハンドル
//---------------------------------------------------------------------------------------------------
//* 注意事項 ： 非表示として確保するので、この処理だけでは見た目上、何も起こりません
//***************************************************************************************************
void __stdcall UpdateThumbnailButton(const THUMBBUTTONDATA* data, HWND hwnd) {
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

    // コールバック用にプロシージャ名を保持
    if (data->ProcedureName) {
        g_procNames[data->ButtonIndex - ButtonID_Correction] = data->ProcedureName;
    }

    //変更を適用
    g_taskbar->ThumbBarUpdateButtons(hwnd, MAX_BUTTONS, g_btns);
}