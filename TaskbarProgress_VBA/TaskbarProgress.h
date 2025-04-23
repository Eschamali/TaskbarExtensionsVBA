#pragma once									//おまじない

//必要なライブラリ等を読み込む
#include <shobjidl.h>							//ITaskbarList3に使用
#include <winrt/base.h>							//WindowsRT APIベース
#include <winrt/Windows.UI.Notifications.h>		//WindowsRT APIの通知関連
#include <winrt/Windows.Data.Xml.Dom.h>			//WindowsRT APIのxml操作関連
#include <atlbase.h>                            //Excelインスタンス制御関連
#include <comdef.h>                             //デバッグによるエラーチェック用
#pragma comment(lib, "comctl32.lib")            //サブクラス関連


//外部参照設定つまりはVBAからでもアクセスできるようにする設定。おまじないと思ってください。
//詳細→https://liclog.net/vba-dll-create-1/
#ifdef TaskbarProgressVBA_EXPORTS
#define TaskbarProgressVBA_API __declspec(dllexport)
#else
#define TaskbarProgressVBA_API __declspec(dllimport)
#endif


//関数ポインタの型を定義し、VBA内のプロシージャ名を呼び出せるようにする
typedef void(__stdcall* VbaCallback)();


// 構造体で、定義します。
// ※VBA側で、シグネチャ（型や順序）が合うようにすること。例外として、BOOLはlongで渡さないと上手くいきません
#pragma pack(4)
struct THUMBBUTTONDATA
{
    const wchar_t* IconPath;
    const wchar_t* Description;
    LONG ButtonType;
    LONG IconIndex;
    LONG ButtonIndex;
};
#pragma pack()


//VBAで扱いたい関数を宣言
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarProgress(HWND hwnd, unsigned long current, unsigned long maximum, long status);
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarOverlayIcon(HWND hwnd, const wchar_t* filePath, int iconIndex, const wchar_t* description);
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarOverlayBadge(int badgeValue, const wchar_t* appUserModelID);
extern "C" TaskbarProgressVBA_API void __stdcall InitializeThumbnailButton(HWND hwnd);
extern "C" TaskbarProgressVBA_API void __stdcall UpdateThumbnailButton(const THUMBBUTTONDATA* data, VbaCallback callback);
