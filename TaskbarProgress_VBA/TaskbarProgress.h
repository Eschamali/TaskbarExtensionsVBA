#pragma once									//���܂��Ȃ�

//�K�v�ȃ��C�u��������ǂݍ���
#include <shobjidl.h>							//ITaskbarList3�Ɏg�p
#include <winrt/base.h>							//WindowsRT API�x�[�X
#include <winrt/Windows.UI.Notifications.h>		//WindowsRT API�̒ʒm�֘A
#include <winrt/Windows.Data.Xml.Dom.h>			//WindowsRT API��xml����֘A
#pragma comment(lib, "comctl32.lib")

//�O���Q�Ɛݒ�܂��VBA����ł��A�N�Z�X�ł���悤�ɂ���ݒ�B���܂��Ȃ��Ǝv���Ă��������B
//�ڍׁFhttps://liclog.net/vba-dll-create-1/
#ifdef TaskbarProgressVBA_EXPORTS
#define TaskbarProgressVBA_API __declspec(dllexport)
#else
#define TaskbarProgressVBA_API __declspec(dllimport)
#endif

// �R�[���o�b�N�֐��^�̒�`�iVBA ����n�����֐��j
typedef void(__stdcall* CallbackFunc)();

//VBA�ň��������֐���錾
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarProgress(HWND hwnd, unsigned long current, unsigned long maximum, long status);
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarOverlayIcon(HWND hwnd, const wchar_t* filePath, int iconIndex, const wchar_t* description);
extern "C" TaskbarProgressVBA_API void __stdcall SetTaskbarOverlayBadge(int badgeValue, const wchar_t* appUserModelID);
extern "C" TaskbarProgressVBA_API void __stdcall SetThumbButtonWithIconEx(HWND hwnd, CallbackFunc callback, const wchar_t* iconPath, int iconIndex);
