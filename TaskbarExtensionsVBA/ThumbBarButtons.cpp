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

constexpr const wchar_t* EXCEL_DESK_CLASS_NAME = L"XLDESK";                 //"XLMAIN"ウィンドウの子名称
constexpr const wchar_t* EXCEL_SHEET_CLASS_NAME = L"EXCEL7";                //"XLDESK"の子名称
constexpr const wchar_t* EXCEL_APPLICATION_CLASS_NAME = L"Application";     //"Application"のオブジェクト名称
constexpr const wchar_t* EXCEL_APPLICATION_RUN_MethodName = L"Run";         //"Application.Run"のメソッド名称



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
//* 機能　　：引数に従った Application オブジェクトを取得します
//---------------------------------------------------------------------------------------------------
//* 引数　　：※割愛します
//---------------------------------------------------------------------------------------------------
//* 詳細説明：WorkbookからApplicationを取得するために使います
//***************************************************************************************************
static HRESULT GetProperty(IDispatch* pDisp, const wchar_t* propName, CComVariant& result) {
    if (!pDisp) return E_POINTER;
    OLECHAR* name = (OLECHAR*)propName;
    DISPID dispID;
    HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) return hr;
    DISPPARAMS params = { NULL, NULL, 0, 0 };
    return pDisp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &result, NULL, NULL);
}

//***************************************************************************************************
//* 機能　　 ：引数にあるインデックスを基に、VBA マクロを実行します
//---------------------------------------------------------------------------------------------------
//* 引数　 　：Index     プロシージャ名があるIndex値
//***************************************************************************************************
void ExecuteVBAProcByIndex(int index) {
    //1.プロシージャ名未登録あるいは、インデックスの範囲外なら、ここで終了
    if (index < 0 || index >= 7 || g_procNames[index].empty()) return;

    //2.COMの初期化
    HRESULT hrInit = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (hrInit == RPC_E_CHANGED_MODE) {
        // 既に異なるアパートメント モードで初期化されている場合は、そのまま続行
        MessageBoxW(nullptr, L"既に異なるアパートメント モードで初期化済み", L"INFO", MB_OK);
    }
    else if (FAILED(hrInit)) {
        wchar_t errorMsg[256];
        swprintf_s(errorMsg, 256, L"COM初期化に失敗しました。HRESULT: 0x%08X", hrInit);
        MessageBoxW(nullptr, errorMsg, L"エラー", MB_OK);
        return;
    }

    //wchar_t title[256];
    //GetWindowTextW(g_hwnd, title, 256);
    //MessageBoxW(nullptr, title, L"対象ウィンドウのタイトル", MB_OK);

    //wchar_t className[256];
    //GetClassNameW(g_hwnd, className, 256);
    //MessageBoxW(nullptr, className, L"対象ウィンドウのクラス名", MB_OK);

    //MessageBoxW(nullptr, g_procNames[index].c_str(), L"対象のプロシージャ名", MB_OK);


    //---------- 3. 孫ウィンドウ経由で、Excel Applicationオブジェクト取得 ----------
    CComPtr<IDispatch> pExcelDispatch;
    HRESULT hr = E_FAIL; // 見つからなかった場合のデフォルト

    // 3 - 1. XLMAINウィンドウの子である「XLDESK」ウィンドウを探す
    HWND hXlDesk = FindWindowExW(g_hwnd, NULL, EXCEL_DESK_CLASS_NAME, NULL);
    if (hXlDesk) {
        // 3 - 2. XLDESKの子である「EXCEL7」ウィンドウを探す
        HWND hExcel7 = FindWindowExW(hXlDesk, NULL, EXCEL_SHEET_CLASS_NAME, NULL);
        if (hExcel7) {
            // 3 - 4. EXCEL7ウィンドウから直接Workbookオブジェクトを取得
            CComPtr<IDispatch> pWorkbookDisp;
            hr = AccessibleObjectFromWindow(hExcel7, OBJID_NATIVEOM, IID_IDispatch, (void**)&pWorkbookDisp);

            if (SUCCEEDED(hr) && pWorkbookDisp) {
                // 3 - 5. WorkbookオブジェクトからApplicationオブジェクトを取得
                CComVariant varApp;
                hr = GetProperty(pWorkbookDisp, EXCEL_APPLICATION_CLASS_NAME, varApp);
                if (SUCCEEDED(hr) && varApp.vt == VT_DISPATCH) {
                    pExcelDispatch = varApp.pdispVal; // 成功！
                }
            }
        }
        else {
            hr = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
    }
    else {
        hr = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }
    // --- ここまで ---

    if (SUCCEEDED(hr) && pExcelDispatch) {
        //成功！
    }
    else {
        _com_error err(hr);
        wchar_t buf[512];
        const wchar_t* reason = L"不明なエラー";
        if (!hXlDesk) reason = L"子ウィンドウ 'XLDESK' が見つかりません";
        else if (!FindWindowExW(hXlDesk, NULL, L"EXCEL7", NULL)) reason = L"孫ウィンドウ 'EXCEL7' が見つかりません";
        else reason = L"EXCEL7からオブジェクト取得に失敗しました";

        swprintf_s(buf, L"エラー理由: %s\nHRESULT=0x%08X\n%s", reason, hr, err.ErrorMessage());
        MessageBoxW(nullptr, buf, L"エラー", MB_OK);

        return;
    }

    // 4. "Run"メソッドのDISPIDを取得する
    DISPID dispidRun;
    OLECHAR* runMethodName = const_cast<OLECHAR*>(EXCEL_APPLICATION_RUN_MethodName); // 実行するメソッド名(VBAのApplication.Run 相当)
    hr = pExcelDispatch->GetIDsOfNames(IID_NULL, &runMethodName, 1, LOCALE_USER_DEFAULT, &dispidRun);

    //　Runメソッドの取得に失敗した場合
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get DISPID for Run method", L"Error", MB_OK);
        return;
    }

    // 5. Application.Run メソッドの引数を設定
    CComVariant macroNameArg(g_procNames[index].c_str());   //引数として渡すのは、VBA側で整形した「'ブック名'!プロシージャ名」

    // 6. Application.Run を呼び出す準備
    CComVariant argsArray[1] = { macroNameArg };
    DISPPARAMS params = { argsArray, nullptr, 1, 0 };

    // 7. "Run"メソッドを起動する
    //　詳細メッセージ、取得用
    EXCEPINFO excepInfo;
    memset(&excepInfo, 0, sizeof(EXCEPINFO));

    CComVariant result;
    hr = pExcelDispatch->Invoke(dispidRun, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,&params, &result, &excepInfo, nullptr);


    //-------------以降は、デバッグ用-------------
    //if (FAILED(hr)) {
    //    // ... 万が一のInvokeエラー処理 ...
    //    _com_error err(hr);
    //    wchar_t errorMsg[256];
    //    swprintf_s(errorMsg, 256, L"HRESULT: 0x%08X", hr);
    //    MessageBoxW(nullptr, errorMsg, L"エラー", MB_OK);

    //    swprintf_s(errorMsg, 256, L"%s", err.ErrorMessage());
    //    MessageBoxW(nullptr, errorMsg, L"根本的エラー", MB_OK);

    //}

    // 現在のExcelインスタンス内に、指定マクロがないと想定
    //if (FAILED(hr)) {
    //    MessageBoxW(nullptr, L"Failed to get Excel macro", L"Error", MB_OK);
    //}

    ////MessageBoxでDISPPARAMSの内容を確認
    //std::wstring debugMessage;

    //// cArgsの確認
    //debugMessage += L"Number of arguments: " + std::to_wstring(params.cArgs) + L"\n";

    //// rgvarg の中身を文字列化
    //for (UINT i = 0; i < params.cArgs; ++i) {
    //    VARIANT& arg = params.rgvarg[i];

    //    if (arg.vt == VT_BSTR) {
    //        debugMessage += L"Argument " + std::to_wstring(i) + L": " + arg.bstrVal + L"\n";
    //    }
    //    else {
    //        debugMessage += L"Argument " + std::to_wstring(i) + L": [not a BSTR]\n";
    //    }
    //}

    //// rgvarg の中身を確認
    //MessageBoxW(nullptr, debugMessage.c_str(), L"DISPPARAMS Debug", MB_OK);

    ////エラーが起こったら、エラーコードと詳細メッセージ(ある場合)を表示。
    //if (FAILED(hr)) {
    //    std::wstring errorMessage = L"Invoke failed. HRESULT: " + std::to_wstring(hr);

    //    if (excepInfo.bstrDescription) {
    //        errorMessage += L"\nException: " + std::wstring(excepInfo.bstrDescription);
    //        SysFreeString(excepInfo.bstrDescription);  // リソース解放
    //    }

    //    MessageBoxW(nullptr, errorMessage.c_str(), L"Error1", MB_OK);
    //}
    //else {
    //    _com_error err(hr);
    //    MessageBoxW(nullptr, err.ErrorMessage(), L"Info", MB_OK);
    //}

    //-------------ここまでが、デバッグ用-------------

    if (SUCCEEDED(hrInit)) {
        CoUninitialize();
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
//              callback 呼び出すVBA内のプロシージャ名
//***************************************************************************************************
void __stdcall UpdateThumbnailButton(const THUMBBUTTONDATA* data, const wchar_t* callback)
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

    // コールバック用にプロシージャ名を保持
    g_procNames[data->ButtonIndex - ButtonID_Correction] = callback;

    //変更を適用
    g_taskbar->ThumbBarUpdateButtons(g_hwnd, MAX_BUTTONS, g_btns);
}
