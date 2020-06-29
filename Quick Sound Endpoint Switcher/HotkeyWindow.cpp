#include "HotkeyWindow.h"

using namespace std;

const wchar_t g_szClassName[] = L"myWindowClass";

//wstring wstrKey;
wstring wstrInstruct1 = L"For best results, press keys individually.";
wstring wstrInstruct2 = L"A second press will toggle a key off";
wstring wstrInstruct3 = L"again.";

bool bShiftKeyUp = true;
bool bControlKeyUp = true;
bool bAltKeyUp = true;
bool bWinKeyUp = true;

UINT HotKeyMods = 0;
UINT HotKeyCode = 0;
LONG hktest = 0;

wstring NewHKStr;
wstring OldHKStr;
LONG HKShift;
LONG HKControl;
LONG HKAlt;
LONG HKWin;
LONG HKKey;

UINT CloseKey;

void MakeHotkeyString(wstring & hotkeystring, UINT key, UINT mods)
{
	WCHAR keyname[128];
	int keylength = 0;

	hotkeystring.clear();

	if (mods & MOD_SHIFT)
	{
		GetKeyNameText(HKShift | 0x2000000, keyname, 128);
		hotkeystring = keyname;
		keylength++;
	}
	if (mods & MOD_CONTROL)
	{
		if (keylength != 0)
			hotkeystring += L"+";
		GetKeyNameText(HKControl | 0x2000000, keyname, 128);
		hotkeystring += keyname;
		keylength++;
	}
	if (mods & MOD_ALT)
	{
		if (keylength != 0)
			hotkeystring += L"+";
		GetKeyNameText(HKAlt | 0x2000000, keyname, 128);
		hotkeystring += keyname;
		keylength++;
	}
	if (mods & MOD_WIN)
	{
		if (keylength != 0)
			hotkeystring += L"+";
		GetKeyNameText(HKWin | 0x2000000, keyname, 128);
		hotkeystring += keyname;
		keylength++;
	}
	if (key != 0)
	{
		if (keylength != 0)
			hotkeystring += L"+";
		GetKeyNameText(HKKey, keyname, 128);
		hotkeystring += keyname;
	}
}

LRESULT CALLBACK WndProc2(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	PAINTSTRUCT ps;
	HDC hdc;

	switch(msg)
	{
	case WM_CREATE:
		break;
	case WM_SHOWWINDOW:
		HotKeyMods = 0;
		HotKeyCode = 0;
		NewHKStr.clear();
		OldHKStr.clear();
		bShiftKeyUp = true;
		bControlKeyUp = true;
		bAltKeyUp = true;
		bWinKeyUp = true;
		break;
	case WM_KEYDOWN:
		switch (wParam)
		{
		case VK_ESCAPE:
			HotKeyMods = 0;
			HotKeyCode = 0;
			CloseKey = wParam;
			DestroyWindow(hWnd);
			break;
		case VK_RETURN:
			UpdateWindow (hWnd);
			DestroyWindow(hWnd);
			CloseKey = wParam;
			break;
		case VK_SHIFT:
			if (bShiftKeyUp)
			{
				bShiftKeyUp = false;
				HotKeyMods ^= MOD_SHIFT;
				HKShift = lParam;
			}
			break;
		case VK_CONTROL:
			if (bControlKeyUp)
			{
				bControlKeyUp = false;
				HotKeyMods ^= MOD_CONTROL;
				HKControl = lParam;
			}
			break;
		case VK_LWIN:
		case VK_RWIN:
			if (bWinKeyUp)
			{
				bWinKeyUp = false;
				HotKeyMods ^= MOD_WIN;
				HKWin = lParam;
			}
			break;
		default:
			hktest = lParam;
			if (wParam == HotKeyCode)
			{
				HotKeyCode = 0;
				break;
			}
			HotKeyCode = wParam;
			HKKey = lParam;
			break;
		}
		break;
	case WM_KEYUP:
		switch (wParam)
		{
		case VK_SHIFT:
			bShiftKeyUp = true;
			break;
		case VK_CONTROL:
			bControlKeyUp = true;
			break;
		case VK_LWIN:
		case VK_RWIN:
			bWinKeyUp = true;
			break;
		default:
			break;
		}
		break;
	case WM_SYSKEYDOWN:
		if (bAltKeyUp)
		{
			bAltKeyUp = false;
			HotKeyMods ^= MOD_ALT;
			HKAlt = lParam;
		}
		break;
	case WM_SYSKEYUP:
		bAltKeyUp = true;
		break;
	case WM_CLOSE:
		DestroyWindow(hWnd);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	case WM_PAINT:
		hdc = BeginPaint (hWnd, &ps);
		TextOut (hdc, 10, 10, wstrInstruct1.data(), wstrInstruct1.length());
		TextOut (hdc, 10, 30, wstrInstruct2.data(), wstrInstruct2.length());
		TextOut (hdc, 10, 50, wstrInstruct3.data(), wstrInstruct3.length());
		TextOut (hdc, 10, 90, NewHKStr.data(), NewHKStr.length());
		EndPaint (hWnd, &ps);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}

	MakeHotkeyString(NewHKStr, HotKeyCode, HotKeyMods);
	if (NewHKStr != OldHKStr)
	{
		OldHKStr = NewHKStr;
		InvalidateRect (hWnd, NULL, TRUE);
		UpdateWindow (hWnd);
	}
	return 0;
}

UINT CreateHotkeyWindow(HWND hWndParent, HINSTANCE hInstance, UINT * HKMods, UINT * HKCode, wstring & HKString)
{
	WNDCLASSEX wc;
	HWND hWnd;
	static bool Registered = false;

	if (Registered == false)
	{
		wc.cbSize        = sizeof(WNDCLASSEX);
		wc.style         = 0;
		wc.lpfnWndProc   = WndProc2;
		wc.cbClsExtra    = 0;
		wc.cbWndExtra    = 0;
		wc.hInstance     = hInstance;
		wc.hIcon         = LoadIcon(NULL, IDI_APPLICATION);
		wc.hCursor       = LoadCursor(NULL, IDC_ARROW);
		wc.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
		wc.lpszMenuName  = NULL;
		wc.lpszClassName = g_szClassName;
		wc.hIconSm       = LoadIcon(NULL, IDI_APPLICATION);

		if(!RegisterClassEx(&wc))
		{
			MessageBox(NULL, L"Hotkey Window Registration Failed!", L"Error!",
				MB_ICONEXCLAMATION | MB_OK);
			return 0;
		}
		else
			Registered = true;
	}

	hWnd = CreateWindowExW(
		WS_EX_CLIENTEDGE,
		g_szClassName,
		L"Press enter to save (escape exits)...",
		WS_DLGFRAME,
		CW_USEDEFAULT, CW_USEDEFAULT, 320, 200,
		hWndParent, NULL, hInstance, NULL);

	if(hWnd == NULL)
	{
		MessageBox(NULL, L"Hotkey Window Creation Failed!", L"Error!",
			MB_ICONEXCLAMATION | MB_OK);
		return 0;
	}

	MSG Msg;
	ShowWindow(hWnd, SW_SHOW);
	UpdateWindow(hWnd);

	//The Message Loop
	while(GetMessage(&Msg, NULL, 0, 0) > 0)
	{
		TranslateMessage(&Msg);
		DispatchMessage(&Msg);
	}

	*HKMods = HotKeyMods;
	*HKCode = HotKeyCode;
	HKString = NewHKStr;

	return CloseKey;
}

