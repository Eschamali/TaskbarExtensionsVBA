#pragma once									//おまじない

//必要なライブラリ等を読み込む
#include <windows.h>                            //WindowsAPI全般
#include <gdiplus.h>                            //内部でアイコン描画
#include <shobjidl.h>							//ITaskbarList3に使用
#include <winrt/base.h>							//WindowsRT APIベース
#include <winrt/Windows.UI.Notifications.h>		//WindowsRT APIの通知関連
#include <winrt/Windows.Data.Xml.Dom.h>			//WindowsRT APIのxml操作関連
#include <atlbase.h>                            //Excelインスタンス制御関連
#include <comdef.h>                             //デバッグによるエラーチェック用
#include <propkey.h>                            //ジャンプリストの制御パラメーター定数関連
#include <propvarutil.h>                        //同じ関数で色んな型の値（文字列、数値、ブールなど）を扱うための、汎用型 PROPVARIANT 
#pragma comment(lib, "comctl32.lib")            //サブクラス関連
#pragma comment(lib, "gdiplus.lib")             //内部でアイコン描画


//外部参照設定つまりはVBAからでもアクセスできるようにする設定。おまじないと思ってください。
//詳細→https://liclog.net/vba-dll-create-1/
#ifdef TaskbarExtensions_EXPORTS
#define TaskbarExtensions_API __declspec(dllexport)
#else
#define TaskbarExtensions_API __declspec(dllimport)
#endif


//関数ポインタの型を定義し、VBA内のプロシージャ名を呼び出せるようにする
typedef void(__stdcall* VbaCallback)();


// VBA→DLLやり取り用ユーザー定義型を定義します。
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

struct JumpListData
{
    const wchar_t* categoryName;
    const wchar_t* taskName;
    const wchar_t* FilePath;
    const wchar_t* cmdArguments;
    const wchar_t* iconPath;
    const wchar_t* Description;
    LONG IconIndex;
};
#pragma pack()

//DLL内 専用構造体
struct JumpListDataSafe
{
    std::wstring categoryName;
    std::wstring taskName;
    std::wstring FilePath;
    std::wstring cmdArguments;
    std::wstring iconPath;
    std::wstring Description;
    LONG IconIndex;
};


//VBAで扱いたい関数を宣言
extern "C" TaskbarExtensions_API void __stdcall SetTaskbarProgress(HWND hwnd, unsigned long current, unsigned long maximum, long status);
extern "C" TaskbarExtensions_API void __stdcall SetTaskbarOverlayIcon(HWND hwnd, const wchar_t* filePath, int iconIndex, const wchar_t* description);
extern "C" TaskbarExtensions_API void __stdcall SetTaskbarOverlayBadge(int badgeValue, const wchar_t* appUserModelID);
extern "C" TaskbarExtensions_API void __stdcall SetTaskbarOverlayBadgeForWin32(LONG badgeValue, HWND hwnd);
extern "C" TaskbarExtensions_API void __stdcall InitializeThumbnailButton(HWND hwnd);
extern "C" TaskbarExtensions_API void __stdcall UpdateThumbnailButton(const THUMBBUTTONDATA* data, VbaCallback callback);
extern "C" TaskbarExtensions_API void __stdcall AddJumpListTask(const JumpListData* data);
extern "C" TaskbarExtensions_API void __stdcall CommitJumpList(const wchar_t* ApplicationModelUserID);
