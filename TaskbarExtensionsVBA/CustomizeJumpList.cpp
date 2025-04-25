/***************************************************************************************************
 *								ジャンプ リストのカスタマイズ
 ***************************************************************************************************
 * 以下の機能を提供します
 * ・「ピン留め」、「最近」以外のカスタムなカテゴリを作成して、そこに任意のタスクを追加
 * ・通常操作では登録できないショートカットファイルも、ジャンプリストへ追加可能
 * ・VBAからの操作に対応
 *
 * 利用API: ICustomDestinationList
 * 
 * URL
 * https://learn.microsoft.com/ja-jp/windows/win32/shell/taskbar-extensions#customizing-jump-lists
 ***************************************************************************************************/



 //設定がまとまってるヘッダーファイルを指定
#include "common.h" 



//***************************************************************************************************
//                           ■■■ 静的ユーザー定義型/定数 ■■■
//***************************************************************************************************
static struct JumpListDataSafe
{
    std::wstring categoryName;
    std::wstring taskName;
    std::wstring FilePath;
    std::wstring cmdArguments;
    std::wstring iconPath;
    std::wstring Description;
    LONG IconIndex;
};

static std::vector<JumpListDataSafe> g_JumpListEntries;	//ジャンプリストデータ保持



//***************************************************************************************************
//                               ■■■ 内部のヘルパー関数 ■■■
//***************************************************************************************************
//* 機能　　 ： ジャンプリストに追加した情報を消去します
//---------------------------------------------------------------------------------------------------
//* 引数　 　： ※割愛
//---------------------------------------------------------------------------------------------------
//* 詳細説明 ： ポインタによるクリア方式を取っています
//***************************************************************************************************
static void CleanupJumpListTask(ICustomDestinationList* pDestList, IObjectCollection* pTasks) {
    if (pTasks) pTasks->Release();
    if (pDestList) pDestList->Release();

    // 蓄積したエントリをクリア
    g_JumpListEntries.clear();
    CoUninitialize();
}



//***************************************************************************************************
//                                 ■■■ VBAから使用する関数 ■■■
//***************************************************************************************************
//* 機能　　 ： ジャンプリストに追加する設定情報を蓄積します
//---------------------------------------------------------------------------------------------------
//* 引数　 　： RegistrationData     ユーザー定義型：JumpListData
//---------------------------------------------------------------------------------------------------
//* 詳細説明 ： VBAで、2次元配列を渡すのがほぼ不可能のため、DLL側のグローバル変数を利用して、予め設定情報を2次元配列的に保存していきます
//***************************************************************************************************
void __stdcall AddJumpListTask(const JumpListData* data) {
    if (data == nullptr) return;

    //中身の値そのものをコピー
    JumpListDataSafe safeData;
    if (data->categoryName) safeData.categoryName = data->categoryName;
    if (data->taskName) safeData.taskName = data->taskName;
    if (data->FilePath) safeData.FilePath = data->FilePath;
    if (data->cmdArguments) safeData.cmdArguments = data->cmdArguments;
    if (data->iconPath) safeData.iconPath = data->iconPath;
    if (data->Description) safeData.Description = data->Description;
    safeData.IconIndex = data->IconIndex;

    //設定情報を蓄積
    g_JumpListEntries.push_back(std::move(safeData));
}

//***************************************************************************************************
//* 機能　　 ： 蓄積されたジャンプリスト設定情報を元にジャンプリストを作成します
//---------------------------------------------------------------------------------------------------
//* 引数　 　： ApplicationModelUserID   ジャンプリスト追加先のApplicationModelUserID
//---------------------------------------------------------------------------------------------------
//* 注意事項 ：・空の設定情報のまま実行すると、ジャンプリストの中身をクリアします。
//             ・設定値に問題(無効な引数等)があった場合、不整合を防ぐため、設定情報をクリアします
//***************************************************************************************************
void __stdcall CommitJumpList(const wchar_t* ApplicationModelUserID)
{
    //必要な変数を用意→初期化
    HRESULT hr;
    ICustomDestinationList* pDestList = nullptr;
    IObjectCollection* pTasks = nullptr;
    IShellLinkW* pLink = nullptr;

    //-------------ジャンプリスト関連COMオブジェクトの準備に関わるお作法-------------
    hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) return;

    hr = CoCreateInstance(CLSID_DestinationList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pDestList));
    if (FAILED(hr)) { CleanupJumpListTask(pDestList, pTasks); return; }
    //-------------------------------------------------------------------------------

    //ジャンプリストの設定先の ApplicationModelUserID の設定先を反映
    pDestList->SetAppID(ApplicationModelUserID);

    //ジャンプリスト編集のセッションを開始
    UINT cMinSlots;
    IObjectArray* poaRemoved;
    hr = pDestList->BeginList(&cMinSlots, IID_PPV_ARGS(&poaRemoved));
    if (FAILED(hr)) { CleanupJumpListTask(pDestList, pTasks); return; }

    //ジャンプリスト登録データ用オブジェクトを用意
    hr = CoCreateInstance(CLSID_EnumerableObjectCollection, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pTasks));
    if (FAILED(hr)) { CleanupJumpListTask(pDestList, pTasks); return; }

    //カテゴリ名、収集準備
    std::map<std::wstring, CComPtr<IObjectCollection>> categoryTasks;

    //設定情報を読み込む処理へ
    for (const auto& entry : g_JumpListEntries) {
        // カテゴリ名が未登録なら新規登録
        if (categoryTasks.find(entry.categoryName) == categoryTasks.end()) {
            CComPtr<IObjectCollection> pNewCollection;
            hr = CoCreateInstance(CLSID_EnumerableObjectCollection, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pNewCollection));
            if (FAILED(hr)) continue;
            categoryTasks[entry.categoryName] = pNewCollection;
        }

        //hellLinkW オブジェクト（ショートカットリンク）を作成。
        hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pLink));
        if (FAILED(hr)) continue;   //ShellLinkW オブジェクト（ショートカットリンク）を生成しようと試みて、もし失敗したらそのエントリの処理をスキップして次のループへ進む
        
        // ---- Separator or Normal Task 判定(ジャンプ リストの タスク セクションに区切り記号を挿入する際の判定用) ----
        bool isSeparator = entry.FilePath.empty();
        if (!isSeparator) {
            //作成したタスクに対して、パラメーターを設定(通常タスク)
            pLink->SetPath(entry.FilePath.c_str());                                                     //実行パス
            if (entry.cmdArguments.c_str()) pLink->SetArguments(entry.cmdArguments.c_str());            //引数
            if (entry.iconPath.c_str()) pLink->SetIconLocation(entry.iconPath.c_str(), entry.IconIndex);//アイコン設定
            if (entry.Description.c_str()) pLink->SetDescription(entry.Description.c_str());            //ツールチップ
        }
        else {
            // Separator の場合は SetPath(nullptr)として、セパレーターとして認識できるようにする
            pLink->SetPath(nullptr);
        }

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

            //「FilePath.empty()」に応じた真偽値
            PROPVARIANT varBoolSeparator;
            PropVariantInit(&varBoolSeparator);
            varBoolSeparator.vt = VT_BOOL;
            varBoolSeparator.boolVal = isSeparator ? VARIANT_TRUE : VARIANT_FALSE;

            //String：タスク名に該当
            PROPVARIANT varTitle;
            InitPropVariantFromString(entry.taskName.c_str(), &varTitle);
            //--------------------------------------------------------------------------

            //------------------------メタデータを設定/適用-----------------------------
            //URL　https://learn.microsoft.com/ja-jp/windows/win32/properties/software-bumper
            //pPropStore->SetValue(PKEY_AppUserModel_PreventPinning, varBoolTrue);            //ピン留め、一覧から削除　を効かなくします
            pPropStore->SetValue(PKEY_AppUserModel_IsDestListSeparator, varBoolSeparator);  //ジャンプリストの 「タスク」セクションに区切り記号を挿入します。
            pPropStore->SetValue(PKEY_Title, varTitle);                                     //タスク名を設定します

            //適用
            pPropStore->Commit();
            //--------------------------------------------------------------------------

            //変数、オブジェクトを解放
            PropVariantClear(&varTitle);
            PropVariantClear(&varBoolTrue);
            pPropStore->Release();
        }
        //指定した「ApplicationModelUserID」(pTasks)に、設定したタスク(pLink->XXX)を入れる。
        categoryTasks[entry.categoryName]->AddObject(pLink);

        //用が済んだので解放
        pLink->Release();
    }

    // エントリがない場合 → ジャンプリストをクリア
    if (g_JumpListEntries.empty()) {
        // 何も追加せず、空データによる CommitList で「クリア」
        pDestList->CommitList();
    }
    else {
        // まとめてジャンプリストに追加
        for (const auto& [category, tasks] : categoryTasks) {
            CComPtr<IObjectArray> pObjectArray;
            hr = tasks->QueryInterface(IID_PPV_ARGS(&pObjectArray));
            if (FAILED(hr)) continue;

            //カテゴリ名が空のものは AddUserTasks に
            if (category.empty()) {
                pDestList->AddUserTasks(pObjectArray);
            }
            //カテゴリ名がある場合は、カテゴリ名ごと追加処理
            else {
                pDestList->AppendCategory(category.c_str(), pObjectArray);
            }
        }

        //ジャンプリスト反映
        pDestList->CommitList();
    }

    //クリーンアップ処理
    CleanupJumpListTask(pDestList, pTasks);
}
