//�ݒ肪�܂Ƃ܂��Ă�w�b�_�[�t�@�C�����w��
#include "TaskbarProgress.h" 

//�悭�g�����O��`��p�ӂ���
using namespace winrt;
using namespace winrt::Windows::UI::Notifications;
using namespace winrt::Windows::Data::Xml::Dom;



//***************************************************************************************************
//                           ������ ThumbButtonInfo �N���X ���������� ������
//***************************************************************************************************
#define MAX_BUTTONS 7                           //�z�u�\�ȃ{�^���̏����
#define ButtonID_Correction 1001                //�{�^��ID�̍̔ԊJ�n�ԍ�

static ITaskbarList3* g_taskbar = nullptr;      //ITaskbarList3�I�u�W�F�N�g
static THUMBBUTTON g_btns[MAX_BUTTONS] = {};    //�{�^�����i�[�p
static std::wstring g_procNames[MAX_BUTTONS];   //�R�[���o�b�N�p�v���V�[�W�����̊i�[�p



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
//* �@�\�@�@ �F�^�X�N�o�[�̃{�^��UI�����w���p�[
//***************************************************************************************************
void EnsureTaskbarInterface() {
    if (!g_taskbar) {
        CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
        CoCreateInstance(CLSID_TaskbarList, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&g_taskbar));
        if (g_taskbar) g_taskbar->HrInit();
    }
}

//***************************************************************************************************
//* �@�\�@�@ �F�����ɂ���v���V�[�W�����ŁAVBA �}�N�������s���܂�
//---------------------------------------------------------------------------------------------------
//* �����@ �@�FIndex     �v���V�[�W����������Index�l
//***************************************************************************************************
void ExecuteVBAProcByIndex(int index) {
    //�v���V�[�W�������o�^���邢�́A�C���f�b�N�X�͈̔͊O�Ȃ�A�����ŏI��
    if (index < 0 || index >= 7 || g_procNames[index].empty()) return;

    //�ڍ׃��b�Z�[�W�A�擾�p(For Debug)
    EXCEPINFO excepInfo;
    memset(&excepInfo, 0, sizeof(EXCEPINFO));  // ������

    // 1. Excel��CLSID���擾
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    // ���炭�AExcel���C���X�g�[������ĂȂ��ꍇ
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get CLSID for Excel", L"Error", MB_OK);
        return;
    }

    // 2. ������Excel�C���X�^���X���擾
    IDispatch* pExcelApp = nullptr;
    hr = GetActiveObject(clsid, nullptr, (IUnknown**)&pExcelApp);
    // �N������Excel���Ȃ��ꍇ
    if (FAILED(hr) || !pExcelApp) {
        MessageBoxW(nullptr, L"Failed to get active Excel instance", L"Error", MB_OK);

        CoUninitialize();
        return;
    }

    // 3. �u Run ���\�b�h�v��DISPID�̎擾
    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"Run");  // ���s���郁�\�b�h��(VBA��Application.Run ����)
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    //Run���\�b�h�̎擾�Ɏ��s�����ꍇ
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get DISPID for Run method", L"Error", MB_OK);
        return;
    }

    // 4. Application.Run ���\�b�h�̈�����ݒ�B
    CComVariant macroName(g_procNames[index].c_str());  //���s�������}�N��(�v���V�[�W��)��
    //�@������
    DISPPARAMS params = {};
    VARIANTARG arg;
    VariantInit(&arg);
    //�@���s�}�N������ݒ�
    _bstr_t procName(macroName);
    //�@�p�����[�^�[�̎d�l���`
    arg.vt = VT_BSTR;
    arg.bstrVal = procName;
    params.rgvarg = &arg;
    params.cArgs = 1;

    // 5. �}�N���̌Ăяo��
    CComVariant result;
    hr = pExcelApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &result, &excepInfo, nullptr);

    //-------------�ȍ~�́A�f�o�b�O�p-------------
    // ���݂�Excel�C���X�^���X���ɁA�w��}�N�����Ȃ��Ƒz��
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get Excel macro", L"Error", MB_OK);
    }

    //MessageBox��DISPPARAMS�̓��e���m�F
    std::wstring debugMessage;

    // cArgs�̊m�F
    debugMessage += L"Number of arguments: " + std::to_wstring(params.cArgs) + L"\n";

    // rgvarg �̒��g�𕶎���
    for (UINT i = 0; i < params.cArgs; ++i) {
        VARIANT& arg = params.rgvarg[i];

        if (arg.vt == VT_BSTR) {
            debugMessage += L"Argument " + std::to_wstring(i) + L": " + arg.bstrVal + L"\n";
        }
        else {
            debugMessage += L"Argument " + std::to_wstring(i) + L": [not a BSTR]\n";
        }
    }

    // rgvarg �̒��g���m�F
    MessageBoxW(nullptr, debugMessage.c_str(), L"DISPPARAMS Debug", MB_OK);

     //�G���[���N��������A�G���[�R�[�h�Əڍ׃��b�Z�[�W(����ꍇ)��\���B
    if (FAILED(hr)) {
        std::wstring errorMessage = L"Invoke failed. HRESULT: " + std::to_wstring(hr);

        if (excepInfo.bstrDescription) {
            errorMessage += L"\nException: " + std::wstring(excepInfo.bstrDescription);
            SysFreeString(excepInfo.bstrDescription);  // ���\�[�X���
        }

        MessageBoxW(nullptr, errorMessage.c_str(), L"Error1", MB_OK);
    }
    else {
        _com_error err(hr);
        MessageBoxW(nullptr, err.ErrorMessage(), L"Info", MB_OK);
    }

    //-------------�����܂ł��A�f�o�b�O�p-------------

    //��n��
    pExcelApp->Release();
    CoUninitialize();
}

//***************************************************************************************************
//* �@�\�@�@ �F���O�ɐݒ肵�� hwnd �ɋN�������Ƃ��A�S�������ɓ͂��܂��B
//---------------------------------------------------------------------------------------------------
//* �����@ �@�Fhwnd      ���b�Z�[�W���󂯎�����E�B���h�E�̃n���h��(�T�u�N���X�ɓo�^����hwnd)
//             msg       ���b�Z�[�W�̎�ށBExcel�Ō����A�C�x���g�̎�ނł��i��FWM_COMMAND, WM_PAINT, WM_CLOSE �Ȃǁj
//             wParam    ���b�Z�[�W�ɂ���ĈӖ����قȂ�⏕�f�[�^�@����1
//             lParam    ���b�Z�[�W�ɂ���ĈӖ����قȂ�⏕�f�[�^�@����2
//---------------------------------------------------------------------------------------------------
//* �@�\���� �FExcel�Ō����A�S�C�x���g�����������ɏW�񂳂�Ă�C���[�W�ł��B�C�x���g���Ƃ̏����́ASwitch�������₷���ł��B
//***************************************************************************************************
LRESULT CALLBACK SubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    //switch���ŁA�C�x���g���ƂɁu��肽�������v������
    switch (msg)
    {
        //�^�X�N�o�[�̃T���l�C���{�^�����N���b�N����ƁAWindows �� WM_COMMAND ���b�Z�[�W�𑗂��Ă��܂��B
        case WM_COMMAND:
            //���̃C�x���g�̓{�^�����N���b�N���ꂽ�ʒm�����肵�܂�(����� THBN_CLICKED �ƂȂ�)
            if (HIWORD(wParam) == THBN_CLICKED) {
                //�␳����
                int buttonIndex = LOWORD(wParam) - ButtonID_Correction;

                //VBA���̃v���V�[�W���������s���鏀����
                ExecuteVBAProcByIndex(buttonIndex);
                return 0;
            }

            break;

        //���̃C�x���g�́A�������܂���
        default:
            break;
    }

    //���̃C�x���g�́A����̏�����
    return DefWindowProc(hwnd, msg, wParam, lParam);
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
//* �@�\�@�@ �F �w�肵���E�B���h�E�n���h���Ƀ{�^�������m�ۂ��܂��B
//---------------------------------------------------------------------------------------------------
//* �����@ �@�F buttonCount     �m�ۂ���{�^����
//              hwnd            �E�B���h�E�n���h��
//---------------------------------------------------------------------------------------------------
//* ���ӎ��� �F ��\���Ƃ��Ċm�ۂ���̂ŁA���̏��������ł͌����ڏ�A�����N����܂���
//***************************************************************************************************
void __stdcall InitializeThumbnailButton(LONG buttonCount, HWND hwnd) {
    //����������
    EnsureTaskbarInterface();

    //0�ȉ��œn���ꂽ��A�{�^�����̂��폜���A�����I��
    if (buttonCount <= 0) {
        memset(g_btns, 0, sizeof(g_btns));
        g_taskbar->ThumbBarAddButtons(hwnd, 0, nullptr);

        //�T�u�N���X���A����
        RemoveWindowSubclass;
        return;
    }

    //����𒴂��Ă���A�������Ȃ�
    if (buttonCount > MAX_BUTTONS) return;

    //��\���Ƃ��āA�{�^�������m�ۂ���
    for (int i = 0; i < MAX_BUTTONS; ++i) {
        g_btns[i].dwMask = THB_FLAGS;
        g_btns[i].dwFlags = THBF_HIDDEN;
        g_btns[i].iId = i + ButtonID_Correction;
        g_btns[i].hIcon = NULL;
        g_btns[i].szTip[0] = L'\0';
    }

    //���f����
    g_taskbar->ThumbBarAddButtons(hwnd, buttonCount, g_btns);

    // �Ώۂ̃E�B���h�E�n���h��(hwnd)���T�u�N���X�����āA�l�X�ȃC�x���g�����ɑΉ�������
    SetWindowSubclass(hwnd, SubclassProc, 1, 0);
}

//***************************************************************************************************
//* �@�\�@�@ �F �w�肵���E�B���h�E�n���h���Ƀ{�^������ύX���܂��B
//---------------------------------------------------------------------------------------------------
//* �����@ �@�F data     ���[�U�[��`�^�FTHUMBBUTTONDATA
//              hwnd     �E�B���h�E�n���h��
//---------------------------------------------------------------------------------------------------
//* ���ӎ��� �F ��\���Ƃ��Ċm�ۂ���̂ŁA���̏��������ł͌����ڏ�A�����N����܂���
//***************************************************************************************************
void __stdcall UpdateThumbnailButton(const THUMBBUTTONDATA* data, HWND hwnd) {
    //������
    EnsureTaskbarInterface();

    //�͈͊O�̃{�^��ID�Ȃ�A�������Ȃ�
    if (!data || data->ButtonIndex  < 0 + ButtonID_Correction || data->ButtonIndex  >= MAX_BUTTONS + ButtonID_Correction) return;

    //�w��{�^��ID�ɑ΂��āA�ǂ�ȗL���ȃf�[�^���܂܂�Ă��邩�`����
    THUMBBUTTON& btn = g_btns[data->ButtonIndex - ButtonID_Correction];
    btn.iId = data->ButtonIndex;                        //�c�[�� �o�[���ň�ӂ̃{�^���̃A�v���P�[�V������`���ʎq�B�O�ׁ̈A1001���獏��
    btn.dwMask = THB_FLAGS | THB_ICON | THB_TOOLTIP;    //�����o�[�ɗL���ȃf�[�^���܂܂�Ă��邩���w�肷�� THUMBBUTTONMASK �l�̑g�ݍ��킹�Bhttps://learn.microsoft.com/ja-jp/windows/win32/api/shobjidl_core/ne-shobjidl_core-thumbbuttonmask
    btn.dwFlags = (THUMBBUTTONFLAGS)data->ButtonType;   //THUMBBUTTON �ɂ���āA�{�^���̓���̏�ԂƓ���𐧌䂷��

    // �c�[���`�b�v
    if (data->Description) {
        wcsncpy_s(btn.szTip, data->Description, ARRAYSIZE(btn.szTip));
    }

    // �A�C�R��
    HICON hIcon = NULL;
    if (data->IconPath) {
        ExtractIconExW(data->IconPath, data->IconIndex, NULL, &hIcon, 1);
    }
    btn.hIcon = hIcon;

    // �R�[���o�b�N�p�Ƀv���V�[�W������ێ�
    if (data->ProcedureName) {
        g_procNames[data->ButtonIndex - ButtonID_Correction] = data->ProcedureName;
    }

    //�ύX��K�p
    g_taskbar->ThumbBarUpdateButtons(hwnd, MAX_BUTTONS, g_btns);
}