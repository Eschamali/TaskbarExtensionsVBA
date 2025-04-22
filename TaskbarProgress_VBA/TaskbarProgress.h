#pragma once									//���܂��Ȃ�

//�K�v�ȃ��C�u��������ǂݍ���
#include <shobjidl.h>							//ITaskbarList3�Ɏg�p
#include <winrt/base.h>							//WindowsRT API�x�[�X
#include <winrt/Windows.UI.Notifications.h>		//WindowsRT API�̒ʒm�֘A
#include <winrt/Windows.Data.Xml.Dom.h>			//WindowsRT API��xml����֘A
#pragma comment(lib, "comctl32.lib")

//�O���Q�Ɛݒ�܂��VBA����ł��A�N�Z�X�ł���悤�ɂ���ݒ�B���܂��Ȃ��Ǝv���Ă��������B
//�ڍׁ�https://liclog.net/vba-dll-create-1/
#ifdef TaskbarProgressVBA_EXPORTS
#define TaskbarProgressVBA_API __declspec(dllexport)
#else
#define TaskbarProgressVBA_API __declspec(dllimport)
#endif

// �\���̂ŁA��`���܂��B
// ��VBA���ŁA�V�O�l�`���i�^�⏇���j�������悤�ɂ��邱�ƁB��O�Ƃ��āABOOL��long�œn���Ȃ��Ə�肭�����܂���
#pragma pack(4)
struct THUMBBUTTONDATA
{
    const wchar_t* ProcedureName;
    const wchar_t* IconPath;
    const wchar_t* Description;
    LONG ButtonType;
    LONG IconIndex;
    LONG ButtonIndex;
};
#pragma pack()


//VBA�ň��������֐���錾
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarProgress(HWND hwnd, unsigned long current, unsigned long maximum, long status);
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarOverlayIcon(HWND hwnd, const wchar_t* filePath, int iconIndex, const wchar_t* description);
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarOverlayBadge(int badgeValue, const wchar_t* appUserModelID);
extern "C" TaskbarProgressVBA_API void __stdcall InitializeThumbnailButton(LONG buttonCount, HWND hwnd);
extern "C" TaskbarProgressVBA_API void __stdcall UpdateThumbnailButton(const THUMBBUTTONDATA* data, HWND hwnd);
