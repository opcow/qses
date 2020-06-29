#pragma once
#include "mmdeviceapi.h"
#include <wmsdk.h>
#include <string>
#include "Prefs.h"

#define WM_USER_NOTIFICATION_DEFAULT WM_USER + 2
#define WM_USER_NOTIFICATION_ADDED WM_USER + 3
#define WM_USER_NOTIFICATION_REMOVED WM_USER + 4
#define WM_USER_NOTIFICATION_CHANGED WM_USER + 5

#ifndef SAFE_RELEASE
#define SAFE_RELEASE(punk)  \
              if ((punk) != NULL)  \
                { (punk)->Release(); (punk) = NULL; }
#endif

class CMMNotificationClient :
	public IMMNotificationClient 
{
	LONG _cRef;
	IMMDeviceEnumerator *_pEnumerator;
	HWND mhWnd;
//	CQSESPrefs * mpPrefs;


	// Private function to print device-friendly name
	//HRESULT _PrintDeviceName(LPCWSTR  pwstrId);

public:

	explicit CMMNotificationClient(HWND hWnd) :
		mhWnd(hWnd),
		_cRef(1),
		_pEnumerator(NULL)
	{
	}

	~CMMNotificationClient()
	{
		SAFE_RELEASE(_pEnumerator);
	}


	    // IUnknown methods -- AddRef, Release, and QueryInterface

    ULONG STDMETHODCALLTYPE AddRef()
    {
        return InterlockedIncrement(&_cRef);
    }

    ULONG STDMETHODCALLTYPE Release()
    {
        ULONG ulRef = InterlockedDecrement(&_cRef);
        if (0 == ulRef)
        {
            delete this;
        }
        return ulRef;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(
                                REFIID riid, VOID **ppvInterface)
    {
        if (IID_IUnknown == riid)
        {
            AddRef();
            *ppvInterface = (IUnknown*)this;
        }
        else if (__uuidof(IMMNotificationClient) == riid)
        {
            AddRef();
            *ppvInterface = (IMMNotificationClient*)this;
        }
        else
        {
            *ppvInterface = NULL;
            return E_NOINTERFACE;
        }
        return S_OK;
    }

    // Callback methods for device-event notifications.

    HRESULT STDMETHODCALLTYPE OnDefaultDeviceChanged(
                                EDataFlow flow, ERole role,
                                LPCWSTR pwstrDeviceId)
    {
		switch (role)
		{
		case eConsole:
			SendMessage(mhWnd, WM_USER_NOTIFICATION_DEFAULT, 0, 0);
			break;
		}

        return S_OK;
    }

	HRESULT STDMETHODCALLTYPE OnDeviceAdded(LPCWSTR pwstrDeviceId)
	{
		SendMessage(mhWnd, WM_USER_NOTIFICATION_ADDED, 0, 0);
		return S_OK;
	};

	HRESULT STDMETHODCALLTYPE OnDeviceRemoved(LPCWSTR pwstrDeviceId)
	{
		SendMessage(mhWnd, WM_USER_NOTIFICATION_REMOVED, 0, 0);
		return S_OK;
	}

    HRESULT STDMETHODCALLTYPE OnDeviceStateChanged(
                                LPCWSTR pwstrDeviceId,
                                DWORD dwNewState)
    {
		switch (dwNewState)
		{
		case DEVICE_STATE_ACTIVE:
		case DEVICE_STATE_DISABLED:
		case DEVICE_STATE_NOTPRESENT:
		case DEVICE_STATE_UNPLUGGED:
		case DEVICE_STATEMASK_ALL:
			SendMessage(mhWnd, WM_USER_NOTIFICATION_CHANGED, 0, 0);
			break;
        }
		return S_OK;
	}

    HRESULT STDMETHODCALLTYPE OnPropertyValueChanged(
                                LPCWSTR pwstrDeviceId,
                                const PROPERTYKEY key)
    {
		return S_OK;
    }
};

