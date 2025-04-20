//�ݒ肪�܂Ƃ܂��Ă�w�b�_�[�t�@�C�����w��
#include "TaskbarProgress.h" 

//�悭�g�����O��`��p�ӂ���
using namespace winrt;
using namespace winrt::Windows::UI::Notifications;
using namespace winrt::Windows::Data::Xml::Dom;

// �O���[�o���ϐ��Ƃ��ĕێ��i��ŌĂяo���j
static CallbackFunc g_callback = nullptr;

// �{�^��ID�̒�`�i�����{�^���Ή������z���Ē�`�j
#define THUMB_BTN_ID 1001



//***************************************************************************************************
//                                 ������ �����̃w���p�[�֐� ������
//***************************************************************************************************
//* �@�\�@�@ �F�����Ńo�b�W���ɕϊ�����֐�
//---------------------------------------------------------------------------------------------------
//* �����@�@ �FbadgeValue    �C�ӂ̐���
//* �Ԃ�l�@ �F�����̐��l�ɉ������o�b�W��
//---------------------------------------------------------------------------------------------------
//* URL      �Fhttps://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/badges
//***************************************************************************************************
static std::wstring GetBadgeValueString(int badgeValue)
{
    switch (badgeValue) {
    case -1:  return L"activity";
    case -2:  return L"alert";
    case -3:  return L"alarm";
    case -4:  return L"available";
    case -5:  return L"away";
    case -6:  return L"busy";
    case -7:  return L"newMessage";
    case -8:  return L"paused";
    case -9:  return L"playing";
    case -10: return L"unavailable";
    case -11: return L"error";
    case -12: return L"attention";
    default:
        if (badgeValue <= -13) return L"none";
        else return std::to_wstring(badgeValue);
    }
}



//***************************************************************************************************
//                                 ������ VBA����g�p����֐� ������
//***************************************************************************************************
//* �@�\�@�@ �F�w��A�v���n���h���̃^�X�N�o�[�ɁA�v���O���X�o�[�̒l�ƃX�e�[�^�X���w�肵�܂��B
//---------------------------------------------------------------------------------------------------
//* �����@ �@�F�Ehwnd     �^�X�N�o�[��K�p������n���h��
//             �Ecurrent  ���ݒl
//             �Emaximum  �ő�l
//             �Estatus�@ ���l�ɂ���āA�F�ύX�A�s�m��ɂł��܂��B
//***************************************************************************************************
void __stdcall SetTaskbarProgress(HWND hwnd, unsigned long current, unsigned long maximum, long status)
{
	// ITaskbarList3�C���^�[�t�F�[�X���擾
	ITaskbarList3* pTaskbarList = nullptr;
	HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pTaskbarList));
	if (FAILED(hr)) {
		return;
	}

	// �^�X�N�o�[�̐i����Ԃ�ݒ�
	pTaskbarList->SetProgressState(hwnd, static_cast<TBPFLAG>(status));

	// �i���l��ݒ�
	if (status == TBPF_NORMAL || status == TBPF_PAUSED || status == TBPF_ERROR) {
		pTaskbarList->SetProgressValue(hwnd, current, maximum);
	}

	// ���\�[�X�̉��
	pTaskbarList->Release();

}

//***************************************************************************************************
//* �@�\�@�@ �F�w��A�v���n���h���̃^�X�N �o�[ �{�^���ɃI�[�o�[���C��K�p���āA�A�v���P�[�V�����̏�Ԃ܂��͒ʒm�����[�U�[�Ɏ����܂��B
//---------------------------------------------------------------------------------------------------
//* �����@ �@�F�Ehwnd        �^�X�N�o�[��K�p������n���h��
//             �EfilePath    �A�C�R�����܂ރt�@�C���t���p�X
//             �EiconIndex   dll���̃A�C�R���Z�b�g��ǂݍ��񂾍ۂ́A�ǂݍ��݈ʒu
//             �Edescription �A�N�Z�V�r���e�B����������
//---------------------------------------------------------------------------------------------------
//* �@�\���� �F�I�[�o�[���C�̍폜�ɂ́AiconIndex �� -1 �ȉ��ɂ��܂��B
//***************************************************************************************************
void __stdcall SetTaskbarOverlayIcon(HWND hwnd, const wchar_t* filePath, int iconIndex, const wchar_t* description)
{
    // ITaskbarList3�C���^�[�t�F�[�X���擾
    ITaskbarList3* pTaskbarList = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pTaskbarList));
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to create ITaskbarList3 instance.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
        return;
    }

    // iconIndex��0�����̏ꍇ�A�A�C�R�����폜����
    if (iconIndex < 0) {
        hr = pTaskbarList->SetOverlayIcon(hwnd, NULL, NULL);
        if (FAILED(hr)) {
            MessageBoxW(nullptr, L"FFailed to remove overlay icon.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
        }
        pTaskbarList->Release();
        return;
    }

    HICON hIcon = NULL;
    std::wstring path(filePath);
    std::wstring extension = path.substr(path.find_last_of(L".") + 1);

    if (extension == L"ico") {
        // .ico�t�@�C������A�C�R�������[�h
        hIcon = (HICON)LoadImage(NULL, filePath, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE | LR_SHARED);
        if (hIcon == NULL) {
            MessageBoxW(nullptr, L"Failed to load .ico file.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
            pTaskbarList->Release();
            return;
        }
    }
    else if (extension == L"exe") {
        // .exe�t�@�C������A�C�R�����C���f�b�N�X�w��Ń��[�h
        hIcon = ExtractIcon(NULL, filePath, iconIndex);
        if (hIcon == NULL || hIcon == (HICON)1) {
            MessageBoxW(nullptr, L"Failed to extract icon from .exe file.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
            pTaskbarList->Release();
            return;
        }
    }
    else if (extension == L"dll") {
        // .dll�t�@�C������A�C�R�����C���f�b�N�X�w��Ń��[�h
        HMODULE hModule = LoadLibraryEx(filePath, NULL, LOAD_LIBRARY_AS_DATAFILE);
        if (hModule == NULL) {
            MessageBoxW(nullptr, L"Failed to load .dll file.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
            pTaskbarList->Release();
            return;
        }

        hIcon = (HICON)LoadImage(hModule, MAKEINTRESOURCE(iconIndex), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
        if (hIcon == NULL) {
            MessageBoxW(nullptr, L"Failed to load icon from resource.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
            FreeLibrary(hModule);
            pTaskbarList->Release();
            return;
        }

        FreeLibrary(hModule);
    }
    else {
        MessageBoxW(nullptr, L"Unsupported file type.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
        pTaskbarList->Release();
        return;
    }

    // �^�X�N�o�[�ɃI�[�o�[���C�A�C�R����ݒ�
    hr = pTaskbarList->SetOverlayIcon(hwnd, hIcon, description);
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to set overlay icon.", L"ITaskbarList3 Error", MB_OK | MB_ICONERROR);
    }

    // �A�C�R�������
    DestroyIcon(hIcon);

    // ���\�[�X�̉��
    pTaskbarList->Release();
}

//***************************************************************************************************
//* �@�\�@�@ �F�w��A�v��ID�̃^�X�N �o�[ �{�^���ɃI�[�o�[���C��K�p���āA�A�v���P�[�V�����̏�Ԃ܂��͒ʒm�����[�U�[�Ɏ����܂�
//---------------------------------------------------------------------------------------------------
//* �����@ �@�F�EbadgeValue        �^�X�N�o�[��K�p������n���h��
//             �EappUserModelID    appUserModelID
//---------------------------------------------------------------------------------------------------
//* �@�\���� �F�A�v���n���h���ł͂Ȃ��AappUserModelID �Ŏw�肷��p�^�[���ł��B
//* ���ӎ��� �F�EWinRT API���̂���OS���K�v�ł�
//             �E�����_�ł́A�f�X�N�g�b�v�A�v���ɑ΂��Ă͌��ʂ���܂���B
//***************************************************************************************************
void __stdcall SetTaskbarOverlayBadge(int badgeValue, const wchar_t* appUserModelID)
{
    // COM�̏�����
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (hr == RPC_E_CHANGED_MODE) {
        // ���ɈقȂ�A�p�[�g�����g ���[�h�ŏ���������Ă���ꍇ�́A���̂܂ܑ��s
    }
    else if (FAILED(hr)) {
        wchar_t errorMsg[256];
        swprintf_s(errorMsg, 256, L"COM�������Ɏ��s���܂����BHRESULT: 0x%08X", hr);
        MessageBoxW(nullptr, errorMsg, L"�G���[", MB_OK | MB_ICONERROR);
        return;
    }

    try {
        // �o�b�W�̒l�𕶎���ɕϊ�
        std::wstring badgeValueStr = GetBadgeValueString(badgeValue);
        std::wstring xmlString = L"<badge value=\"" + badgeValueStr + L"\"/>";

        // XML�̓ǂݍ���
        XmlDocument doc;
        doc.LoadXml(winrt::hstring(xmlString));

        // �o�b�W�ʒm�I�u�W�F�N�g�̍쐬
        BadgeNotification badge(doc);

        // �w�肵��AppID�̒ʒm�}�l�[�W�����擾
        auto notifier = BadgeUpdateManager::CreateBadgeUpdaterForApplication(winrt::hstring(appUserModelID));
        notifier.Update(badge);
    }
    catch (...) {
        // �G���[����
        MessageBoxW(nullptr, L"�o�b�W�ʒm�̕\���Ɏ��s���܂����B", L"Badge Error", MB_OK | MB_ICONERROR);
    }
}

//***************************************************************************************************
//* �@�\�@�@ �F�^�X�N�o�[�{�^���������ꂽ�Ƃ��̒ʒm���󂯎��AVBA�֐����Ăяo���E�B���h�E�v���V�[�W���B
//---------------------------------------------------------------------------------------------------
//* �����@ �@�F���������܂�
//---------------------------------------------------------------------------------------------------
//* �@�\���� �F�T�u�N���X�v���V�[�W���i�{�^�������Ȃǂ̃��b�Z�[�W���󂯎��j
//***************************************************************************************************
LRESULT CALLBACK SubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    if (msg == WM_COMMAND) {
        if (LOWORD(wParam) == THUMB_BTN_ID && g_callback) {
            // �{�^���������ꂽ�Ƃ��AVBA ����n���ꂽ�֐������s
            (*g_callback)();
        }
    }
    // ���̑��̃��b�Z�[�W�͊���̏�����
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

//***************************************************************************************************
//* �@�\�@�@ �F VBA ������R�[���o�b�N�֐��|�C���^��o�^���邽�߂̊֐�
//---------------------------------------------------------------------------------------------------
//* �����@ �@�F callback     ���s��������VBA�֐���(������ł͂Ȃ��A�A�h���X)
//---------------------------------------------------------------------------------------------------
//* �@�\���� �FVBA ����֐��|�C���^��o�^���邽�߂̃G�N�X�|�[�g�֐��B
//***************************************************************************************************
void __stdcall SetThumbButtonCallback(CallbackFunc callback)
{
    g_callback = callback;
}

//***************************************************************************************************
//* �@�\�@�@ �F �w�肵���E�B���h�E�n���h���Ƀ{�^����ǉ����T�u�N���X��(���C������)
//---------------------------------------------------------------------------------------------------
//* �����@ �@�F hwnd     �E�B���h�E�n���h��
//---------------------------------------------------------------------------------------------------
//* �@�\���� �F�E�B���h�E�n���h�������ƂɁA�^�X�N�o�[�Ƀ{�^����ǉ����鏈���B
//             �����͊�{�AVBA �� Application.hwnd ��n������
//***************************************************************************************************
void __stdcall AddThumbButton(HWND hwnd)
{
    // �T�u�N���X�����ă��b�Z�[�W�t�b�N���J�n
    SetWindowSubclass(hwnd, SubclassProc, 1, 0);

    // �^�X�N�o�[�C���^�[�t�F�[�X�̎擾
    ITaskbarList3* pTaskbar = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pTaskbar));
    if (SUCCEEDED(hr)) {
        pTaskbar->HrInit();

        // �{�^������ݒ�
        THUMBBUTTON thumbButton = {};
        thumbButton.iId = THUMB_BTN_ID;
        thumbButton.dwMask = THB_FLAGS | THB_TOOLTIP;
        thumbButton.dwFlags = THBF_ENABLED;
        wcscpy_s(thumbButton.szTip, L"VBA�}�N�����s");

        // �{�^����ǉ�
        pTaskbar->ThumbBarAddButtons(hwnd, 1, &thumbButton);
        pTaskbar->Release();
    }
}
