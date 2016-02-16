#include <wchar.h>
#include "windows.h"
#include <shellapi.h>
#include <string>
#include <sstream>
#include <tchar.h>
#include <stdlib.h>
#include <Mmdeviceapi.h>
#include <Shlobj.h>

#include "PolicyConfig.h"
#include "Propidl.h"
#include "Functiondiscoverykeys_devpkey.h"
#include "QSSIMMNotificationClient.h"

#include "Prefs.h"
#include "Settings.h"
#include "resource.h"

using namespace std;

#define MAX_LOADSTRING 100
#define	WM_USER_SHELLICON WM_USER + 1

#define ID_MENU_ABOUT 100
#define ID_MENU_DONATE 101
#define ID_MENU_DEVICES 120
#define ID_MENU_EXIT 140
#define ID_MENU_SETTINGS 150
#define CYCLE_HOTKEY 200

// Global Variables:
HINSTANCE hInst;
NOTIFYICONDATA nidApp;
HICON hMainIcon;
bool gbDevicesChanged = false;

TCHAR gszTitle[MAX_LOADSTRING];
TCHAR gszWindowClass[MAX_LOADSTRING];
TCHAR gszApplicationToolTip[MAX_LOADSTRING];
HINSTANCE ghInstance;
HWND gOpenWindow = 0;
const int cMaxDevices = 10;
UINT gHotkeyFlags = 0;
wstringstream gVersionString;

CSoundSourcePrefs gPrefs;
CMMNotificationClient * pClient = 0;


// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK	About(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK	HotkeyDlg(HWND, UINT, WPARAM, LPARAM);
void DoSettingsDialog(HINSTANCE hInst, HWND hWnd);

void InstallNotificationCallBack();
int EnumerateDevices();
bool SetDefaultAudioPlaybackDevice(LPCWSTR devID);
bool GetDeviceIcon(HICON * hIcon);
bool IsDefaultAudioPlaybackDevice(const wstring& s);
wstring& GetDefaultAudioPlaybackDevice(wstring& s);

void OpenWebsite (WCHAR * cpURL)
{
      ShellExecute (NULL, L"open", cpURL, NULL, NULL, SW_SHOWNORMAL);
}

void MakeStartupLinkPath(TCHAR * name, TCHAR * path)
{
	HRESULT hres;
	hres = SHGetFolderPath(NULL, CSIDL_STARTUP, NULL, SHGFP_TYPE_CURRENT, path);
	if (SUCCEEDED(hres))
	{
		_tcscat_s(path, MAX_PATH, TEXT("\\"));
		_tcscat_s(path, MAX_PATH, name);
		_tcscat_s(path, MAX_PATH, TEXT(".lnk"));
	}
	else
		*path = TEXT('\0');
}

void EnableAutoStart()
{
	IShellLink* pShellLink = NULL;
	IPersistFile* pPersistFile;
	WIN32_FIND_DATA fd;
    TCHAR szStartupLinkName[MAX_PATH];
	HRESULT hres;

	hres = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_ALL, IID_IShellLink, (void**)&pShellLink);
	if (SUCCEEDED(hres))
	{
		hres = pShellLink->QueryInterface(IID_IPersistFile, (void**)&pPersistFile);
		if (SUCCEEDED(hres))
		{
			HMODULE hModule = GetModuleHandle(NULL);
			WCHAR szExeName[MAX_PATH];
			GetModuleFileNameW(hModule, szExeName, MAX_PATH);
			pShellLink->SetPath(szExeName);
			pShellLink->SetDescription(gszApplicationToolTip);
			pShellLink->SetIconLocation(szExeName, 0);
			hres = pShellLink->QueryInterface(IID_IPersistFile, (void**)&pPersistFile);
			if (SUCCEEDED(hres))
			{
				MakeStartupLinkPath(gszApplicationToolTip, szStartupLinkName);
				pPersistFile->Save(szStartupLinkName, TRUE);
				pPersistFile->Release();
			}
			pShellLink->Release();
		}
	}
}

void DisableAutoStart(void)
{
    TCHAR szStartupLinkName[MAX_PATH];
	HRESULT hres;

	MakeStartupLinkPath(gszApplicationToolTip, szStartupLinkName);
	if (*szStartupLinkName != TEXT('\0'))
		DeleteFile(szStartupLinkName);
}

UINT CheckAutoStart(void)
{
    TCHAR szStartupLinkName[MAX_PATH];
	HRESULT hres;

	MakeStartupLinkPath(gszApplicationToolTip, szStartupLinkName);
	HANDLE hFile = CreateFile(szStartupLinkName, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE != hFile)
	{
		CloseHandle(hFile);
		return 1;
	}
	return 0;
}

wstring& MakeHotkeyString(UINT code, UINT mods, wstring & str)
{
	WCHAR keyname[256];
	UINT scanCode;
	int length = 0;


	if (mods & MOD_SHIFT)
	{
		scanCode = MapVirtualKey(VK_SHIFT, MAPVK_VK_TO_VSC);
		length = GetKeyNameText(scanCode << 16, keyname, 255);
	}
	if (mods & MOD_CONTROL)
	{
		if (length > 0) { keyname[length++] = L'+'; keyname[length] = 0; }
		scanCode = MapVirtualKey(VK_CONTROL, MAPVK_VK_TO_VSC);
		length += GetKeyNameText(scanCode << 16, keyname + length, 255 - length);
	}
	if (mods & MOD_ALT)
	{
		if (length > 0) { keyname[length++] = L'+'; keyname[length] = 0; }
		scanCode = MapVirtualKey(VK_MENU, MAPVK_VK_TO_VSC);
		length += GetKeyNameText(scanCode << 16, keyname + length, 255 - length);
	}
	if (mods & MOD_WIN)
	{
		if (length > 0) { keyname[length++] = L'+'; keyname[length] = 0; }
		scanCode = MapVirtualKey(VK_LWIN, MAPVK_VK_TO_VSC);
		length += GetKeyNameText(scanCode << 16, keyname + length, 255 - length);
	}
	if (code != 0)
	{
		if (length > 0) { keyname[length++] = L'+'; keyname[length] = 0; }
		scanCode = MapVirtualKey(code, MAPVK_VK_TO_VSC);
		length += GetKeyNameText(scanCode << 16, keyname + length, 255 - length);
	}
	str += L" (";
	str += keyname;
	return str += L")";
}

void RemoveHotkeys(HWND hWnd)
{
	UINT count = gPrefs.GetCount();

	for (UINT i = 0; i < count  && i < cMaxDevices; i++)
	{
		if (gPrefs.GetHotkeyEnabled(i))
		{
			UnregisterHotKey(hWnd,i);
		}
	}
	UnregisterHotKey(hWnd, CYCLE_HOTKEY);
}

// Message Box with human readable error if hotkey fails to register.
void HandleHotkeyError(wstring& s)
{
	WCHAR ErrorStr[512] = {0};
	s += L"\n\n";
	DWORD err = GetLastError();
	FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, err, MAKELANGID(LANG_NEUTRAL,SUBLANG_DEFAULT), ErrorStr, 511, 0);
	s += ErrorStr;
	MessageBox(NULL, s.data(), L"Hotkey Error", MB_OK);

}

// Install all hotkeys. Do this on load or when the prefs change.
void InstallHotkeys(HWND hWnd)
{
	UINT count = gPrefs.GetCount();
	wstring s;

	for (UINT i = 0; i < count  && i < cMaxDevices; i++)
	{
		if (!gPrefs.GetIsPresent(i))
			continue;
		if (gPrefs.GetHotkeyEnabled(i))
		{
			UnregisterHotKey(hWnd,i);
			if(0 == RegisterHotKey(hWnd, i, gPrefs.GetHotkeyMods(i) | gHotkeyFlags, gPrefs.GetHotkeyCode(i)))
				HandleHotkeyError(gPrefs.GetName(i, s));
		}
	}
	if (gPrefs.GetCycleKeyEnabled())
	{
		UnregisterHotKey(hWnd, CYCLE_HOTKEY);
		if(0 == RegisterHotKey(hWnd, CYCLE_HOTKEY, gPrefs.GetCycleKeyMods() | gHotkeyFlags, gPrefs.GetCycleKeyCode()))
		{
			s = L"Cycle Hotkey";
			HandleHotkeyError(s);
		}
	}
}

// Function to cycle through visible and cycle-enabled devices.
void CycleDevices()
{
	wstring IdString;
	UINT count = gPrefs.GetCount();
//	UINT start = (gPrefs.GetDefault() + 1);
	UINT start = gPrefs.FindByID(GetDefaultAudioPlaybackDevice(IdString)) + 1;
	UINT j;

	for (UINT i = 0; i < count; i++)
	{
		j = start++ % count;
		gPrefs.GetName(j, IdString);
		if(!gPrefs.GetExcludeFromCycle(j) && !gPrefs.GetIsHidden(j) && gPrefs.GetIsPresent(j))
		{
			SetDefaultAudioPlaybackDevice((gPrefs.GetID(j, IdString)).data());
//			gPrefs.SetDefault(j);
			break;
		}
	}
}

// Check to see if we're already running.
bool AlreadyRunning()
{
#define APPLICATION_INSTANCE_MUTEX_NAME L"{cf5daad3-840e-419c-8be3-3131897500ee}"

    //Make sure at most one instance of the tool is running
    HANDLE hMutexOneInstance(::CreateMutex(NULL, TRUE, APPLICATION_INSTANCE_MUTEX_NAME));
    bool bAlreadyRunning((::GetLastError() == ERROR_ALREADY_EXISTS));
    if (hMutexOneInstance == NULL || bAlreadyRunning)
    {
    	if(hMutexOneInstance)
    	{
    		::ReleaseMutex(hMutexOneInstance);
    		::CloseHandle(hMutexOneInstance);
    	}
    }
	return bAlreadyRunning;
}

// Check OS version and set hotkey flag (MOD_NOREPEAT isn't supported on Vista)
bool CheckOSVersion()
{
	OSVERSIONINFOEX osvi;
	SYSTEM_INFO si;
	BOOL bOsVersionInfoEx;

	ZeroMemory(&si, sizeof(SYSTEM_INFO));
	ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));

	osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
	bOsVersionInfoEx = GetVersionEx((OSVERSIONINFO*) &osvi);

	if (osvi.dwMajorVersion < 6)
	{
		MessageBox(NULL, L"Requires Windows Vista or later.", L"Quick Sound Source", MB_OK);
		return false;
	}
	else if (osvi.dwMinorVersion >=1)
		gHotkeyFlags = MOD_NOREPEAT;

	return true;
}

void MakegVersionString(HINSTANCE hInstance)
{
	DWORD dwVerInfoSize;
	DWORD dwHnd;
	VS_FIXEDFILEINFO *pFixedInfo;	// pointer to fixed file info structure
	UINT    uVersionLen;			// Current length of full version string
	WCHAR AppPathName[MAX_PATH];
	WCHAR AppName[MAX_LOADSTRING+1];

	ZeroMemory(AppName, (MAX_LOADSTRING+1) * sizeof(WCHAR));

	GetModuleFileName(hInstance, AppPathName, MAX_PATH);
	dwVerInfoSize = GetFileVersionInfoSize(AppPathName, &dwHnd);
	void *pBuffer = new char [dwVerInfoSize];
	if(pBuffer == 0)
		return;
	GetFileVersionInfo(AppPathName, dwHnd, dwVerInfoSize, pBuffer);
	VerQueryValue(pBuffer,_T("\\"),(void**)&pFixedInfo,(UINT *)&uVersionLen);

	LoadString(hInstance, IDS_APP_TITLE, AppName, MAX_LOADSTRING);
	gVersionString << AppName
		<< L" " << HIWORD (pFixedInfo->dwProductVersionMS)
		<< L"." << LOWORD (pFixedInfo->dwProductVersionMS)
		<< L"." << HIWORD (pFixedInfo->dwProductVersionLS);
	if(LOWORD (pFixedInfo->dwProductVersionLS) != 0)
		gVersionString << L"." << LOWORD (pFixedInfo->dwProductVersionLS);
}

int APIENTRY _tWinMain(_In_ HINSTANCE hInstance,
					   _In_opt_ HINSTANCE hPrevInstance,
					   _In_ LPTSTR    lpCmdLine,
					   _In_ int       nCmdShow)
{
	if (!CheckOSVersion())
		return 0;

	if (AlreadyRunning())
	{
		MessageBox(NULL, L"An instance of Quick Sound Source is already running.", L"Quick Sound Source", MB_OK);
		return false;
	}

	MSG msg;
	HACCEL hAccelTable;

	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, gszTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_QSS, gszWindowClass, MAX_LOADSTRING);
	LoadString(hInstance, IDS_APPTOOLTIP, gszApplicationToolTip, MAX_LOADSTRING);

	ghInstance = hInstance;
	MyRegisterClass(hInstance);

	CoInitializeEx(NULL, COINIT_MULTITHREADED);

	gPrefs.Init(10);
	if(!gPrefs.Load())
	{
		gPrefs.Save();
		EnableAutoStart();
	}
	EnumerateDevices();

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return FALSE;
	}

	hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_QSS));

	// Main message loop:
	while (int ret = GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	CoUninitialize();
	return (int) msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_QSS);
	wcex.lpszClassName	= gszWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

	return RegisterClassEx(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	HWND hWnd;

	hInst = hInstance; // Store instance handle in our global variable

	hWnd = CreateWindow(gszWindowClass, gszTitle, WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);

	if (!hWnd)
		return FALSE;

	hMainIcon = LoadIcon(hInstance,(LPCTSTR)MAKEINTRESOURCE(IDI_SMALL)); 

	nidApp.cbSize = sizeof(NOTIFYICONDATA);
	nidApp.hWnd = (HWND) hWnd;
	nidApp.uID = IDI_SMALL; 
	nidApp.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
	nidApp.uCallbackMessage = WM_USER_SHELLICON;
	wcsncpy_s(nidApp.szTip, 64, gszApplicationToolTip, _TRUNCATE);
	if (!GetDeviceIcon(&nidApp.hIcon))
		nidApp.hIcon = hMainIcon;
	Shell_NotifyIcon(NIM_ADD, &nidApp); 

	return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	POINT lpClickPoint;
	UINT count;
	UINT devID;
	HMENU hPopMenu;
	wstring MenuString, TempString;

	static UINT s_uTaskbarRestart;

	switch (message)
	{
	case WM_HOTKEY:
		if (wParam == CYCLE_HOTKEY)
			CycleDevices();
		else
			SetDefaultAudioPlaybackDevice((gPrefs.GetID(wParam, TempString)).data());
		return TRUE;
	case WM_CREATE:
		pClient = new CMMNotificationClient(hWnd);
		// load the preferences and fill in audio devices
		InstallNotificationCallBack();
		InstallHotkeys(hWnd);
		return TRUE;
	case WM_USER_NOTIFICATION_DEFAULT:
		DestroyIcon(nidApp.hIcon);
		if (!GetDeviceIcon(&nidApp.hIcon))
			nidApp.hIcon = hMainIcon;
		Shell_NotifyIcon(NIM_MODIFY,  &nidApp);
		return TRUE;
	case WM_USER_NOTIFICATION_ADDED:
	case WM_USER_NOTIFICATION_REMOVED:
	case WM_USER_NOTIFICATION_CHANGED:
		gbDevicesChanged = true;
		EnumerateDevices();
		return TRUE;
	case WM_USER_SHELLICON: 
		// systray msg callback 
		switch(LOWORD(lParam)) 
		{   
		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
			if (IsWindow(gOpenWindow))
			{
				SetForegroundWindow(gOpenWindow);
				return TRUE;
			}
			if (gbDevicesChanged)
			{
				EnumerateDevices();
				gbDevicesChanged = false;
			}
			GetCursorPos(&lpClickPoint);
			hPopMenu = CreatePopupMenu();
			InsertMenu(hPopMenu,0xFFFFFFFF, MF_BYPOSITION|MF_STRING, ID_MENU_ABOUT, _T("About..."));
			InsertMenu(hPopMenu,0xFFFFFFFF, MF_BYPOSITION|MF_STRING, ID_MENU_DONATE, _T("Donate (Opens Web Browser)"));
			InsertMenu(hPopMenu, 0xFFFFFFFF, MF_SEPARATOR|MF_BYPOSITION, 0, NULL);

			count = gPrefs.GetCount();
			for (UINT i = 0; i < count; i++)
			{
				if (gPrefs.GetIsHidden(i) || !gPrefs.GetIsPresent(i))
					continue;
				gPrefs.GetName(i, MenuString);
				if (gPrefs.GetHotkeyEnabled(i))
					MenuString = MakeHotkeyString(gPrefs.GetHotkeyCode(i), gPrefs.GetHotkeyMods(i), MenuString);
				if (IsDefaultAudioPlaybackDevice(gPrefs.GetID(i, TempString)))
					InsertMenu(hPopMenu,0xFFFFFFFF, MF_BYPOSITION|MF_STRING|MF_CHECKED, ID_MENU_DEVICES + i, MenuString.data());
				else
					InsertMenu(hPopMenu,0xFFFFFFFF, MF_BYPOSITION|MF_STRING, ID_MENU_DEVICES + i, MenuString.data());
			}

			gPrefs.GetCycleKeyString(TempString);

			InsertMenu(hPopMenu, 0xFFFFFFFF, MF_SEPARATOR|MF_BYPOSITION, 0, NULL);
			InsertMenu(hPopMenu, 0xFFFFFFFF, MF_BYPOSITION|MF_STRING, ID_MENU_SETTINGS, _T("Settings..."));
			InsertMenu(hPopMenu, 0xFFFFFFFF, MF_SEPARATOR|MF_BYPOSITION, 0, NULL);
			InsertMenu(hPopMenu, 0xFFFFFFFF, MF_BYPOSITION|MF_STRING, ID_MENU_EXIT, _T("Exit"));
			//								
			SetForegroundWindow(hWnd);
			TrackPopupMenu(hPopMenu,TPM_LEFTALIGN|TPM_LEFTBUTTON|TPM_BOTTOMALIGN,lpClickPoint.x, lpClickPoint.y,0,hWnd,NULL);
			return TRUE; 
		}
		break;
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);

		switch (wmId)
		{
		case ID_MENU_ABOUT:
			DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG_ABOUT), hWnd, About);
			return TRUE;
		case ID_MENU_DONATE:
			OpenWebsite(L"https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=PSBXPNT3ZDG2Q");
			return TRUE;
		case ID_MENU_EXIT:
			Shell_NotifyIcon(NIM_DELETE,&nidApp);
			SendMessage(hWnd, WM_CLOSE, 0, 0);
			DestroyWindow(hWnd);
			return TRUE;
		case ID_MENU_SETTINGS:
			RemoveHotkeys(hWnd);
			DoSettingsDialog(hInst, hWnd);
			EnumerateDevices();
			InstallHotkeys(hWnd);
			return TRUE;
		default:
			if (wmId >= (ID_MENU_DEVICES) && wmId < (ID_MENU_DEVICES + cMaxDevices))
			{
				devID = wmId - ID_MENU_DEVICES;
				SetDefaultAudioPlaybackDevice((gPrefs.GetID(devID, TempString)).data());
				return TRUE;
			}
			else
				return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		if(message == s_uTaskbarRestart)
			Shell_NotifyIcon(NIM_ADD, &nidApp);
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return FALSE;
}

void DoSettingsDialog(HINSTANCE hInst, HWND hWnd)
{
	CSoundSourcePrefs TempPrefs;
	INT_PTR rc;

	TempPrefs = gPrefs;
	rc = DialogBoxParam(hInst, MAKEINTRESOURCE(IDD_DIALOG_SETTINGS), hWnd, SettingsDialogProc, (LPARAM ) & TempPrefs);

	if (rc == IDOK)
	{
		gPrefs = TempPrefs;
		gPrefs.Save();
	}
}
// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		gOpenWindow = hDlg;
		MakegVersionString(ghInstance);
		SetWindowText(GetDlgItem(hDlg, IDC_STATIC_APP), gVersionString.str().data());
		return (INT_PTR)TRUE;
	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	case WM_DESTROY:
		gOpenWindow = 0;
		return (INT_PTR)TRUE;
	}
	return (INT_PTR)FALSE;
}

#define EXIT_ON_ERROR(hres)  \
              if (FAILED(hres)) { goto Exit; }

bool GetDeviceIcon(HICON * hIcon)
{
	wstring path;
	IMMDeviceEnumerator *pEnum = NULL;
	IPropertyStore *pStore;
	PROPVARIANT propVariant;
	IMMDevice *pDevice;
	HRESULT hr;

	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum);
	EXIT_ON_ERROR(hr)
	hr = pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
	EXIT_ON_ERROR(hr)
	hr = pDevice->OpenPropertyStore(STGM_READ, &pStore);
	EXIT_ON_ERROR(hr)
	hr = pStore->GetValue(PKEY_DeviceClass_IconPath,  &propVariant);
	path = propVariant.pwszVal;
	PropVariantClear(&propVariant);

	{
		UINT nIconID;
		wstring wsIconPath;
		wstring wsIconId;

		wsIconPath = path;
		UINT pos = wsIconPath.find(L',');
		if (pos != std::string::npos)
		{
			wsIconId = wsIconPath.substr(pos+1, string::npos);
			wsIconPath.erase(pos, string::npos);
			nIconID = (UINT) _wtoi(wsIconId.data());
			*hIcon = ExtractIcon(ghInstance, wsIconPath.data(), nIconID);
		}
	}
Exit:
	SAFE_RELEASE(pStore);
	SAFE_RELEASE(pDevice);
	SAFE_RELEASE(pEnum);

	if (hIcon != 0 && hIcon != (HICON *) 1)
		return true;
	return false;
}

bool SetDefaultAudioPlaybackDevice(LPCWSTR devID)
{	
	IPolicyConfigVista *pPolicyConfig;
	ERole reserved = eConsole;
	HRESULT hr;// = CoInitializeEx(NULL, COINIT_MULTITHREADED);

	//if (SUCCEEDED(hr))
	{
		hr = CoCreateInstance(__uuidof(CPolicyConfigVistaClient), 
			NULL, CLSCTX_ALL, __uuidof(IPolicyConfigVista), (LPVOID *)&pPolicyConfig);
		if (SUCCEEDED(hr))
		{
			hr = pPolicyConfig->SetDefaultEndpoint(devID, reserved);
			pPolicyConfig->Release();
		}
	}
	//CoUninitialize();
	return SUCCEEDED(hr); // fixme make void?
}

void InstallNotificationCallBack()
{
	HRESULT hr;

	IMMDeviceEnumerator *pEnum = NULL;
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum);
	if (SUCCEEDED(hr))
	{
		pEnum->RegisterEndpointNotificationCallback(pClient);
		pEnum->Release();
	}
}

wstring& GetDefaultAudioPlaybackDevice(wstring& s)
{
	HRESULT hr;
	IMMDeviceEnumerator *pEnum = NULL;
	IMMDevice *pDefDevice;
	LPWSTR pwszDefID = NULL;

	s.clear();
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum);
	EXIT_ON_ERROR(hr);
	hr = pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pDefDevice);
	EXIT_ON_ERROR(hr);
	hr = pDefDevice->GetId(&pwszDefID);
	EXIT_ON_ERROR(hr);
	s = pwszDefID;

Exit:
	CoTaskMemFree(pwszDefID);
	SAFE_RELEASE(pEnum);
	SAFE_RELEASE(pDefDevice);
	return s;
}

bool IsDefaultAudioPlaybackDevice(const wstring& s)
{
	HRESULT hr;
	IMMDeviceEnumerator *pEnum = NULL;
	IMMDevice *pDefDevice;
	LPWSTR pwszDefID = NULL;
	bool bMatch = false;

	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum);
	EXIT_ON_ERROR(hr);
	hr = pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pDefDevice);
	EXIT_ON_ERROR(hr);
	hr = pDefDevice->GetId(&pwszDefID);
	EXIT_ON_ERROR(hr);

	if (s == pwszDefID)
		bMatch = true;
Exit:
	CoTaskMemFree(pwszDefID);
	SAFE_RELEASE(pEnum);
	SAFE_RELEASE(pDefDevice);
	return bMatch;
}

int EnumerateDevices()
{
	LPWSTR wstrID = NULL;
	UINT count = 0;
	DevicePrefs SoundDeviceInfo;

	wstring wsDefaultDeviceID(L"");

	SoundDeviceInfo.DeviceID = L"";
	SoundDeviceInfo.HotkeyString = L"";
	SoundDeviceInfo.HasHotkey = false;
	SoundDeviceInfo.IsExcludedFromCycle = false;
	SoundDeviceInfo.IsHidden = false;
	SoundDeviceInfo.IsPresent = false;
	SoundDeviceInfo.KeyCode = 0;
	SoundDeviceInfo.KeyMods = 0;

	HRESULT hr;

	IMMDeviceEnumerator *pEnum = NULL;
	// Create a multimedia device enumerator.
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum);

	EXIT_ON_ERROR(hr);
	IMMDeviceCollection *pDevices;
	// Enumerate the output devices.
	hr = pEnum->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pDevices);
	EXIT_ON_ERROR(hr);
	pDevices->GetCount(&count);
	if (count > cMaxDevices)
		count = cMaxDevices;

	EXIT_ON_ERROR(hr);
	gPrefs.ResetPresent(); // mark all devices not present

	IMMDevice *pDevice;
	IPropertyStore *pStore;
	PROPVARIANT propVariant;

	for (UINT i = 0; i < count; i++)
	{
		hr = pDevices->Item(i, &pDevice);
		EXIT_ON_ERROR(hr);
		hr = pDevice->GetId(&wstrID);
		SoundDeviceInfo.DeviceID = wstrID;
		CoTaskMemFree(wstrID);
		EXIT_ON_ERROR(hr);
		hr = pDevice->OpenPropertyStore(STGM_READ, &pStore);
		EXIT_ON_ERROR(hr);
		PropVariantInit(&propVariant);
		hr = pStore->GetValue(PKEY_Device_FriendlyName, &propVariant);
		EXIT_ON_ERROR(hr);
		SoundDeviceInfo.Name = propVariant.pwszVal;
		gPrefs.Update(&SoundDeviceInfo);
		PropVariantClear(&propVariant);
	}
Exit:
	SAFE_RELEASE(pStore);
	SAFE_RELEASE(pDevice);
	SAFE_RELEASE(pDevices);
	SAFE_RELEASE(pEnum);
	gPrefs.Sort();
	return 0;
}