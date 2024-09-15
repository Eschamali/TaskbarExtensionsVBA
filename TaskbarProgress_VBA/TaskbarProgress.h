#include "windows.h"

#ifdef TaskbarProgressVBA_EXPORTS
#define TaskbarProgressVBA_API __declspec(dllexport)
#else
#define TaskbarProgressVBA_API __declspec(dllimport)
#endif


//VBAÇ≈àµÇ¢ÇΩÇ¢ä÷êîÇêÈåæ
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarProgress(HWND hwnd, unsigned long current, unsigned long maximum, long status);
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarOverlayIcon(HWND hwnd, const wchar_t* filePath, int iconIndex, const wchar_t* description);