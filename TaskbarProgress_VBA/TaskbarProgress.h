#pragma once									//おまじない

//必要なライブラリ等を読み込む
#include <shobjidl.h>							//ITaskbarList3に使用
#include <winrt/base.h>							//WindowsRT APIベース
#include <winrt/Windows.UI.Notifications.h>		//WindowsRT APIの通知関連
#include <winrt/Windows.Data.Xml.Dom.h>			//WindowsRT APIのxml操作関連
#pragma comment(lib, "comctl32.lib")

//外部参照設定つまりはVBAからでもアクセスできるようにする設定。おまじないと思ってください。
//詳細：https://liclog.net/vba-dll-create-1/
#ifdef TaskbarProgressVBA_EXPORTS
#define TaskbarProgressVBA_API __declspec(dllexport)
#else
#define TaskbarProgressVBA_API __declspec(dllimport)
#endif

// コールバック関数型の定義（VBA から渡される関数）
typedef void(__stdcall* CallbackFunc)();

//VBAで扱いたい関数を宣言
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarProgress(HWND hwnd, unsigned long current, unsigned long maximum, long status);
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarOverlayIcon(HWND hwnd, const wchar_t* filePath, int iconIndex, const wchar_t* description);
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarOverlayBadge(int badgeValue, const wchar_t* appUserModelID);
extern "C" TaskbarProgressVBA_API void __stdcall SetThumbButtonWithIconEx(HWND hwnd, CallbackFunc callback, const wchar_t* iconPath, int iconIndex);
