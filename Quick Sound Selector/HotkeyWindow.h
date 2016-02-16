#pragma once

#include <wchar.h>
#include <tchar.h>
#include "windows.h"
#include <string>


UINT CreateHotkeyWindow(HWND hWndParent, HINSTANCE hInstance, UINT * HKMods, UINT * HKCode, std::wstring & HKString);
