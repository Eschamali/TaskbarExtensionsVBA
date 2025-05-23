﻿/***************************************************************************************************
 *									アイコン オーバーレイ
 ***************************************************************************************************
 * 以下の機能を提供します
 * ・タスク バー ボタンを使用して、特定の通知と状態アイコンを表示させ、ユーザーに伝えることができます。
 * ・UWPアプリ向けのバッジ通知も対応
 * ・通知数的なバッチも対応(予定)
 * ・VBAからの操作に対応
 *
 * 利用API: ITaskbarList3::SetOverlayIcon、BadgeNotification
 *
 * URL
 * https://learn.microsoft.com/ja-jp/windows/win32/shell/taskbar-extensions#icon-overlays
 * https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/badges
 ***************************************************************************************************/



 //設定がまとまってるヘッダーファイルを指定
#include "common.h" 



 // C++の名前空間（namespace）を省略して使えるようにする
using namespace winrt;
using namespace winrt::Windows::UI::Notifications;
using namespace winrt::Windows::Data::Xml::Dom;
using namespace Gdiplus;



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
//* 機能　　 ：Win32アプリ用、通知バッチアイコンを作成します
//---------------------------------------------------------------------------------------------------
//* 引数　 　：number    描画する数字。99を超える場合、"99+"と表示します
//---------------------------------------------------------------------------------------------------
//* 詳細説明 ：
//***************************************************************************************************
static HICON CreateBadgeIcon(int number)
{
    // GDI+ 初期化
    static bool initialized = false;
    static GdiplusStartupInput gdiplusStartupInput;
    static ULONG_PTR gdiplusToken;
    if (!initialized) {
        GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, nullptr);
        initialized = true;
    }

    const int canvasSize = 32;  // 正方形キャンバス（アイコン全体）
    const int ellipseWidth = 32;
    const int ellipseHeight = (number <= 9) ? 32 : 28;
    int FontSize;

    Bitmap bmp(canvasSize, canvasSize, PixelFormat32bppARGB);
    Graphics g(&bmp);
    g.SetSmoothingMode(SmoothingModeAntiAlias);
    g.Clear(Color(0, 0, 0, 0));  // 完全透明背景

    // --- 楕円の中心座標計算（キャンバス中央に配置） ---
    int x = (canvasSize - ellipseWidth) / 2;
    int y = (canvasSize - ellipseHeight) / 2;

    // --- 赤い楕円描画 ---
    SolidBrush redBrush(Color(255, 255, 0, 0));  // 赤色
    g.FillEllipse(&redBrush, x, y, ellipseWidth, ellipseHeight);

    // --- 白文字の準備 ---
    WCHAR buf[4];
    if (number <= 99)
        wsprintf(buf, L"%d", number);
    else
        lstrcpyW(buf, L"99+");

    //フォントサイズ決め
    if (number <= 9)
        FontSize = 24;
    else if (number <= 99)
        FontSize = 20;
    else
        FontSize = 14;

    FontFamily fontFamily(L"Segoe UI");
    Gdiplus::Font font(&fontFamily, FontSize, FontStyleBold, UnitPixel);
    SolidBrush whiteBrush(Color(255, 255, 255));  // 白

    // --- 文字描画の範囲を楕円内に限定 ---
    RectF layoutRect((REAL)x, (REAL)y, (REAL)ellipseWidth, (REAL)ellipseHeight);
    StringFormat format;
    format.SetAlignment(StringAlignmentCenter);
    format.SetLineAlignment(StringAlignmentCenter);

    g.DrawString(buf, -1, &font, layoutRect, &format, &whiteBrush);

    // アイコンに変換して返す
    HICON hIcon = nullptr;
    bmp.GetHICON(&hIcon);
    return hIcon;
}



//***************************************************************************************************
//                                 ■■■ VBAから使用する関数 ■■■
//***************************************************************************************************
//* 機能　　 ：指定アプリハンドルのタスクバーボタンに、オーバーレイアイコンを適用して、アプリケーションの状態または通知をユーザーに示します。
//---------------------------------------------------------------------------------------------------
//* 引数　 　：・hwnd        タスクバーを適用させるハンドル
//             ・filePath    アイコンを含むファイルフルパス
//             ・iconIndex   dll等のアイコンセットを読み込んだ際の、読み込み位置
//             ・description アクセシビリティ向け説明文
//---------------------------------------------------------------------------------------------------
//* 機能説明 ：オーバーレイアイコンの削除は、iconIndex を -1 以下にします。
//***************************************************************************************************
void __stdcall SetTaskbarOverlayIcon(HWND hwnd, const wchar_t* filePath, int iconIndex, const wchar_t* description)
{     
    //数値→文字列変換用変数
    wchar_t buffer[256];

    // ITaskbarList3インターフェースを取得
    ITaskbarList3* pTaskbarList = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pTaskbarList));
    if (FAILED(hr)) {
        // 初期化失敗したので、イベントビュアーへ記録
        swprintf(buffer, 256, L"Failed to create ITaskbarList3 instance.\nErrorCode：0x%08X", hr);
        WriteToEventViewer(2, SourceName, buffer, EVENTLOG_ERROR_TYPE, 0, TRUE);
        return;
    }

    // iconIndexが0未満の場合、アイコンを削除して、処理終了
    if (iconIndex < 0) {
        hr = pTaskbarList->SetOverlayIcon(hwnd, NULL, NULL);
        if (FAILED(hr)) {
            // アイコン削除に失敗したので、イベントビュアーへ記録
            swprintf(buffer, 256, L"Failed to remove overlay icon.\nErrorCode：0x%08X", hr);
            WriteToEventViewer(3, SourceName, buffer, EVENTLOG_ERROR_TYPE, 0, TRUE);
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
            // アイコン読み込み失敗したので、イベントビュアーへ記録
            WriteToEventViewer(4, SourceName, L"Failed to load .ico file.",  EVENTLOG_ERROR_TYPE, 0, TRUE);
            pTaskbarList->Release();
            return;
        }
    }
    else if (extension == L"exe") {
        // .exeファイルからアイコンをインデックス指定でロード
        hIcon = ExtractIcon(NULL, filePath, iconIndex);
        if (hIcon == NULL || hIcon == (HICON)1) {
            // .exeファイルからアイコン抽出失敗時、イベントビュアーへ記録
            WriteToEventViewer(5, SourceName, L"Failed to extract icon from .exe file.",  EVENTLOG_ERROR_TYPE, 0, TRUE);
            pTaskbarList->Release();
            return;
        }
    }
    else if (extension == L"dll") {
        // .dllファイルからアイコンをインデックス指定でロード
        HMODULE hModule = LoadLibraryEx(filePath, NULL, LOAD_LIBRARY_AS_DATAFILE);
        if (hModule == NULL) {
            // .dllファイルからアイコン抽出失敗時、イベントビュアーへ記録
            WriteToEventViewer(6, SourceName, L"Failed to load .dll file.", EVENTLOG_ERROR_TYPE, 0, TRUE);
            pTaskbarList->Release();
            return;
        }

        hIcon = (HICON)LoadImage(hModule, MAKEINTRESOURCE(iconIndex), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
        if (hIcon == NULL) {
            //アイコンリソースの抽出失敗時、イベントビュアーへ記録
            WriteToEventViewer(7, SourceName, L"Failed to load icon from resource.", EVENTLOG_ERROR_TYPE, 0, TRUE);
            FreeLibrary(hModule);
            pTaskbarList->Release();
            return;
        }

        FreeLibrary(hModule);
    }
    else {
        //iconデータでないファイルを渡された時
        WriteToEventViewer(8, SourceName, L"Unsupported file type.", EVENTLOG_ERROR_TYPE, 0, TRUE);
        pTaskbarList->Release();
        return;
    }

    // タスクバーにオーバーレイアイコンを設定
    hr = pTaskbarList->SetOverlayIcon(hwnd, hIcon, description);
    if (FAILED(hr)) {
        //オーバレイ反映に失敗した時
        WriteToEventViewer(9, SourceName, L"Failed to set overlay icon.", EVENTLOG_ERROR_TYPE, 0, TRUE);
    }

    // アイコンを解放
    DestroyIcon(hIcon);

    // リソースの解放
    pTaskbarList->Release();
}

//***************************************************************************************************
//* 機能　　 ：UWPアプリ向けにある通知数バッチを「ITaskbarList3::SetOverlayIcon」を使って、再現します
//---------------------------------------------------------------------------------------------------
//* 引数　 　：・badgeValue        数字を指定
//             ・hwnd              ウィンドウハンドル
//---------------------------------------------------------------------------------------------------
//* 詳細説明 ：・SetTaskbarOverlayBadge を win32 アプリでも扱えるようにしたものです。現時点では、通知数的に扱えます
//*            ・仕組みは、SetOverlayIcon + GDI + でメモリ上にアイコンを描画し、HICONを生成 で実現してます
//             ・0以下でアイコン削除、1～99で、そのまんまの表示、100以上で、"99+" と表示します
//***************************************************************************************************
void __stdcall SetTaskbarOverlayBadgeForWin32(LONG badgeValue, HWND hwnd)
{
    //数値→文字列変換用変数
    wchar_t buffer[256];

    // ITaskbarList3インターフェースを取得
    ITaskbarList3* pTaskbarList = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pTaskbarList));
    if (FAILED(hr)) {
        // 初期化失敗したので、イベントビュアーへ記録
        swprintf(buffer, 256, L"Failed to create ITaskbarList3 instance.\nErrorCode：0x%08X", hr);
        WriteToEventViewer(2, SourceName, buffer, EVENTLOG_ERROR_TYPE, 0, TRUE);
        return;
    }

    // badgeValueが0以下の場合、アイコンを削除して、処理終了
    if (badgeValue <= 0) {
        hr = pTaskbarList->SetOverlayIcon(hwnd, NULL, NULL);
        if (FAILED(hr)) {
            // アイコン削除に失敗したので、イベントビュアーへ記録
            swprintf(buffer, 256, L"Failed to remove overlay icon.\nErrorCode：0x%08X", hr);
            WriteToEventViewer(3, SourceName, buffer, EVENTLOG_ERROR_TYPE, 0, TRUE);
        }
        pTaskbarList->Release();
        return;
    }

    //内部でアイコン生成
    HICON icon = CreateBadgeIcon(badgeValue);

    //反映処理
    pTaskbarList->SetOverlayIcon(hwnd, icon, NULL);
}

//***************************************************************************************************
//* 機能　　 ：指定AppUserModelIDのタスク バー ボタンにオーバーレイを適用して、アプリケーションの状態または通知をユーザーに示します
//---------------------------------------------------------------------------------------------------
//* 引数　 　：・badgeValue        タスクバーを適用させるハンドル
//             ・appUserModelID    AppUserModelID
//---------------------------------------------------------------------------------------------------
//* 機能説明 ：アプリハンドルではなく、AppUserModelID で指定するパターンです。
//* 注意事項 ：・WinRT API環境のあるOSが必要です
//             ・現時点では、デスクトップアプリに対しては効果ありません。
//***************************************************************************************************
void __stdcall SetTaskbarOverlayBadge(int badgeValue, const wchar_t* appUserModelID)
{
    //数値→文字列変換用変数
    wchar_t buffer[256];

    // COMの初期化
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (hr == RPC_E_CHANGED_MODE) {
        // 既に異なるアパートメント モードで初期化されている場合は、そのまま続行
    }
    else if (FAILED(hr)) {
        swprintf(buffer, 256, L"COM初期化に失敗しました。HRESULT: 0x%08X", hr);
        WriteToEventViewer(10, SourceName, buffer, EVENTLOG_ERROR_TYPE, 0, TRUE);
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
        WriteToEventViewer(11, SourceName, L"バッジ通知の表示に失敗しました。", EVENTLOG_ERROR_TYPE, 0, TRUE);
    }
}
