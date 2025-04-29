/***************************************************************************************************
 *											進行状況バー
 ***************************************************************************************************
 * 以下の機能を提供します
 * ・タスクバーに、進行状況バーを表示できます。
 * ・赤、黄カラーに対応し、操作が一時停止しているか、エラーが発生し、ユーザーの介入が必要であることを示すこともできます。
 * ・VBAからの操作に対応
 *
 * 利用API: ITaskbarList3::SetProgressState、ITaskbarList3::SetProgressValue
 *
 * URL		  
 * https://learn.microsoft.com/ja-jp/windows/win32/shell/taskbar-extensions#progress-bars
 ***************************************************************************************************/



//設定がまとまってるヘッダーファイルを指定
#include "common.h" 



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
		//数値→文字列変換用変数
		wchar_t buffer[256];

		// 処理失敗したので、イベントビュアーへ記録
		swprintf(buffer, 256, L"SetTaskbarProgress 関数にて、CoCreateInstance failed.\nErrorCode：0x%08X", hr);
		WriteToEventViewer(2, SourceName, buffer, EVENTLOG_ERROR_TYPE, 0, TRUE);

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
