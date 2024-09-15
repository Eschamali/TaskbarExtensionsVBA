#include "TaskbarProgress.h" 

//�K�v�Ȋ֐��Z�b�g���C���|�[�g
#include "windows.h"
#include "shobjidl.h"

#include "iostream"  // �f�o�b�O�p

//�w��A�v���n���h���̃^�X�N�o�[�ɁA�v���O���X�o�[�̒l�ƃX�e�[�^�X���w�肵�܂��B
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

//�w��A�v���n���h���̃^�X�N �o�[ �{�^���ɃI�[�o�[���C��K�p���āA�A�v���P�[�V�����̏�Ԃ܂��͒ʒm�����[�U�[�Ɏ����܂��B
void __stdcall SetTaskbarOverlayIcon(HWND hwnd, const wchar_t* filePath, int iconIndex, const wchar_t* description)
{
    // ITaskbarList3�C���^�[�t�F�[�X���擾
    ITaskbarList3* pTaskbarList = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pTaskbarList));
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create ITaskbarList3 instance" << std::endl;
        return;
    }

    // filePath��NULL���󕶎��̏ꍇ�̓A�C�R�����폜����
    if (filePath == nullptr || wcslen(filePath) == 0) {
        hr = pTaskbarList->SetOverlayIcon(hwnd, NULL, NULL);
        if (FAILED(hr)) {
            std::wcerr << L"Failed to remove overlay icon: " << hr << std::endl;
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
            std::wcerr << L"Failed to load .ico file: " << GetLastError() << std::endl;
            pTaskbarList->Release();
            return;
        }
    }
    else if (extension == L"exe") {
        // .exe�t�@�C������A�C�R�����C���f�b�N�X�w��Ń��[�h
        hIcon = ExtractIcon(NULL, filePath, iconIndex);
        if (hIcon == NULL || hIcon == (HICON)1) {
            std::wcerr << L"Failed to extract icon from .exe file: " << GetLastError() << std::endl;
            pTaskbarList->Release();
            return;
        }
    }
    else if (extension == L"dll") {
        // .dll�t�@�C������A�C�R�����C���f�b�N�X�w��Ń��[�h
        HMODULE hModule = LoadLibraryEx(filePath, NULL, LOAD_LIBRARY_AS_DATAFILE);
        if (hModule == NULL) {
            std::wcerr << L"Failed to load .dll file: " << GetLastError() << std::endl;
            pTaskbarList->Release();
            return;
        }

        hIcon = (HICON)LoadImage(hModule, MAKEINTRESOURCE(iconIndex), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
        if (hIcon == NULL) {
            std::wcerr << L"Failed to load icon from resource: " << GetLastError() << std::endl;
            FreeLibrary(hModule);
            pTaskbarList->Release();
            return;
        }

        FreeLibrary(hModule);
    }
    else {
        std::wcerr << L"Unsupported file type: " << extension << std::endl;
        pTaskbarList->Release();
        return;
    }

    // �^�X�N�o�[�ɃI�[�o�[���C�A�C�R����ݒ�
    hr = pTaskbarList->SetOverlayIcon(hwnd, hIcon, description);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to set overlay icon: " << hr << std::endl;
    }

    // �A�C�R�������
    DestroyIcon(hIcon);

    // ���\�[�X�̉��
    pTaskbarList->Release();
}