//***************************************************************************************************
//									サムネイルツールバー
//***************************************************************************************************
// 以下の機能を提供します
// ・サムネイルプレビュー下部へのボタン追加
// ・VBA コールバック連携対応
// 
// 利用WindowsAPI
// ・ITaskbarList3::ThumbBarAddButtons
// ・ITaskbarList3::ThumbBarUpdateButtons
// 
// URL
// https://learn.microsoft.com/ja-jp/windows/win32/shell/taskbar-extensions#thumbnail-toolbars
//***************************************************************************************************



//設定がまとまってるヘッダーファイルを指定
#include "common.h" 



//***************************************************************************************************
//                                  ■■■ 静的変数/定数 ■■■
//***************************************************************************************************
static constexpr int MAX_BUTTONS = 7;                                   //配置可能なボタンの上限数
static constexpr int ButtonID_Correction = 1001;                        //ボタンIDの採番開始番号

static ITaskbarList3* g_taskbar = nullptr;                              //ITaskbarList3オブジェクト
static THUMBBUTTON g_btns[MAX_BUTTONS] = {};                            //ボタン情報格納用
static std::wstring g_procNames[MAX_BUTTONS];                           //コールバック用プロシージャ名の格納用
static HWND g_hwnd = nullptr;                                           //InitializeThumbnailButton で呼び出したウィンドウハンドルを保持します
static VbaCallback g_thumbButtonCallbacks[MAX_BUTTONS] = { nullptr };   //各ボタン用の関数(VBA内プロシージャ)ポインタを7つ保持



//***************************************************************************************************
//                                 ■■■ 内部のヘルパー関数 ■■■
//***************************************************************************************************
//* 機能　　 ：ITaskbarList3 インターフェースを初期化して使えるようにするためのものです
//***************************************************************************************************
static void EnsureTaskbarInterface()
{
    //数値→文字列変換用変数
    wchar_t buffer[256];

    if (!g_taskbar) {
        HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
        if (FAILED(hr)) {
            // 初期化失敗したので、イベントビュアーへ記録
            swprintf(buffer, 256, L"CoInitializeEx failed.\nErrorCode：0x%08X", hr);
            WriteToEventViewer(1, SourceName, buffer, EVENTLOG_ERROR_TYPE,0,TRUE);
            return;
        }

        hr = CoCreateInstance(CLSID_TaskbarList, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&g_taskbar));
        if (SUCCEEDED(hr) && g_taskbar) {
            g_taskbar->HrInit();
        }
        else {
            // 処理失敗したので、イベントビュアーへ記録
            swprintf(buffer, 256, L"CoCreateInstance failed.\nErrorCode：0x%08X", hr);
            WriteToEventViewer(2, SourceName, buffer, EVENTLOG_ERROR_TYPE, 0, TRUE);
        }
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
//* 詳細説明 ：Excelで言う、全イベント処理がここに集約されてるイメージです。イベントごとの処理は、Switch文がやりやすいです。
//***************************************************************************************************
static LRESULT CALLBACK SubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
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
//                                 ■■■ VBAから使用する関数 ■■■
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
    if (!data || data->ButtonIndex < 0 + ButtonID_Correction || data->ButtonIndex >= MAX_BUTTONS + ButtonID_Correction) return;

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

    //インデックス値が負になってたら、アイコンなしのままにします
    if (data->IconPath && data->IconIndex >= 0) {
        ExtractIconExW(data->IconPath, data->IconIndex, NULL, &hIcon, 1);
    }
    else {
        DestroyIcon(btn.hIcon);
    }
    btn.hIcon = hIcon;

    // コールバック用にプロシージャ名のポインタを保持
    g_thumbButtonCallbacks[data->ButtonIndex - ButtonID_Correction] = callback;

    //変更を適用
    g_taskbar->ThumbBarUpdateButtons(g_hwnd, MAX_BUTTONS, g_btns);
}
