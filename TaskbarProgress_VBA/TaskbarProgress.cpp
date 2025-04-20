//設定がまとまってるヘッダーファイルを指定
#include "TaskbarProgress.h" 

//よく使う名前定義を用意する
using namespace winrt;
using namespace winrt::Windows::UI::Notifications;
using namespace winrt::Windows::Data::Xml::Dom;

// グローバル変数として保持（後で呼び出す）
static CallbackFunc g_callback = nullptr;

// ボタンIDの定義（複数ボタン対応を見越して定義）
#define THUMB_BTN_ID 1001



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
//* 機能　　 ：タスクバーボタンが押されたときの通知を受け取り、VBA関数を呼び出すウィンドウプロシージャ。
//---------------------------------------------------------------------------------------------------
//* 引数　 　：※割愛します
//---------------------------------------------------------------------------------------------------
//* 機能説明 ：サブクラスプロシージャ（ボタン押下などのメッセージを受け取る）
//***************************************************************************************************
LRESULT CALLBACK SubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    if (msg == WM_COMMAND) {
        if (LOWORD(wParam) == THUMB_BTN_ID && g_callback) {
            // ボタンが押されたとき、VBA から渡された関数を実行
            (*g_callback)();
        }
    }
    // その他のメッセージは既定の処理へ
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

//***************************************************************************************************
//* 機能　　 ： VBA 側からコールバック関数ポインタを登録するための関数
//---------------------------------------------------------------------------------------------------
//* 引数　 　： callback     実行させたいVBA関数名(文字列ではなく、アドレス)
//---------------------------------------------------------------------------------------------------
//* 機能説明 ：VBA から関数ポインタを登録するためのエクスポート関数。
//***************************************************************************************************
void __stdcall SetThumbButtonCallback(CallbackFunc callback)
{
    g_callback = callback;
}

//***************************************************************************************************
//* 機能　　 ： 指定したウィンドウハンドルにボタンを追加＆サブクラス化(メイン処理)
//---------------------------------------------------------------------------------------------------
//* 引数　 　： hwnd     ウィンドウハンドル
//---------------------------------------------------------------------------------------------------
//* 機能説明 ：ウィンドウハンドルをもとに、タスクバーにボタンを追加する処理。
//             引数は基本、VBA の Application.hwnd を渡すこと
//***************************************************************************************************
void __stdcall AddThumbButton(HWND hwnd)
{
    // サブクラス化してメッセージフックを開始
    SetWindowSubclass(hwnd, SubclassProc, 1, 0);

    // タスクバーインターフェースの取得
    ITaskbarList3* pTaskbar = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pTaskbar));
    if (SUCCEEDED(hr)) {
        pTaskbar->HrInit();

        // ボタン情報を設定
        THUMBBUTTON thumbButton = {};
        thumbButton.iId = THUMB_BTN_ID;
        thumbButton.dwMask = THB_FLAGS | THB_TOOLTIP;
        thumbButton.dwFlags = THBF_ENABLED;
        wcscpy_s(thumbButton.szTip, L"VBAマクロ実行");

        // ボタンを追加
        pTaskbar->ThumbBarAddButtons(hwnd, 1, &thumbButton);
        pTaskbar->Release();
    }
}
