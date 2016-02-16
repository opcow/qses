#pragma once

#include <string>
#include "windows.h"
#include <time.h>

#include <boost/archive/tmpdir.hpp>
#include <boost/archive/text_woarchive.hpp>
#include <boost/archive/text_wiarchive.hpp>
#include <boost/serialization/list.hpp>
#include <boost/serialization/string.hpp>
#include <boost/serialization/version.hpp>
#include <boost/serialization/split_member.hpp>

using namespace std;
namespace bs = boost::serialization;

struct DevicePrefs
{
	template<class Archive>
	void serialize(Archive & ar, const unsigned int version)
    {
		ar & Name;
		ar & DeviceID;
		ar & HasHotkey;
		ar & HotkeyString;
		ar & KeyCode;
		ar & KeyMods;
		ar & IsExcludedFromCycle;
		ar & IsHidden;
		ar & TimeStamp;
	}
	UINT KeyMods;
	UINT KeyCode;
	wstring Name;
	wstring DeviceID;
	wstring HotkeyString;
	bool HasHotkey;
	bool IsExcludedFromCycle;
	bool IsHidden;
	bool IsPresent;
	time_t TimeStamp;
};

class CSoundSourcePrefs
{

	friend class boost::serialization::access;

public:

	// class config
	void Add(DevicePrefs *sdi);
	void Update(DevicePrefs *sdi);
	bool SetMax(int max);
	int GetMax() { return mMax; };
	int GetCount() { return mNext; };
	// hotkey
	void SetHotkeyString(int index, const wstring & keystring) { if (index < mNext) mDevices[index].HotkeyString = keystring; }
	wstring& GetHotkeyString(int index, wstring & keyString);
	void SetHotkeyCode(int index, UINT code);
	UINT GetHotkeyCode(int index) { if (index < mNext) return mDevices[index].KeyCode; else return 0; }
	void SetHotkeyMods(int index, UINT mods);
	UINT GetHotkeyMods(int index);
	void EnableHotkey(int index, bool enabled) { if (index < mNext) mDevices[index].HasHotkey = enabled; }
	bool GetHotkeyEnabled(int index);
	// cycle
	void SetExcludeFromCycle(int index, bool b) { if (index < mNext) mDevices[index].IsExcludedFromCycle = b; }
	bool GetExcludeFromCycle(int index);
	void SetCycleKeyString(const WCHAR * s) { mCycleKeyString = s; }
	wstring& GetCycleKeyString(wstring& s) { return s = mCycleKeyString; }
	void SetCycleKeyCode(UINT key) { mKeyCode = key; }
	UINT GetCycleKeyCode() { return mKeyCode; }
	void SetCycleKeyMods(UINT mods) { mKeyMods = mods; }
	UINT GetCycleKeyMods() { return mKeyMods; }
	void EnableCycleKey(bool enabled) { mIsCycleKeyEnabled = enabled; }
	bool GetCycleKeyEnabled() { return mIsCycleKeyEnabled; }
	// misc
	void SetIsHidden(int index, bool b) { if (index < mNext) mDevices[index].IsHidden = b; }
	bool GetIsHidden(int index);
	void SetIsPresent(int index, bool b) { if (index < mNext) mDevices[index].IsPresent = b; }
	bool GetIsPresent(int index);
	//void SetDefault(int index) { if (index < mNext) mDefault = index; }
	//int GetDefault() { return mDefault; }
	void Swap(int d1, int d2);
	void Sort();
	// device info
	bool GetByName(wstring& name, DevicePrefs * sdi);
	int FindByID(const wstring& id);
	wstring& GetName(int index, wstring & s);
	wstring& GetID(int index, wstring& s);
	// prefs class
	bool Init(int max);
	void Clear(int index);
	void ResetPresent();
	// load & save
	bool Load();
	bool Save();

	CSoundSourcePrefs( const CSoundSourcePrefs& other );
	CSoundSourcePrefs() : mDevices(0), mNext(0), mMax(10), /*mDefault(0),*/ mKeyCode(0),
							mKeyMods(0), mCycleKeyString(L""), mIsCycleKeyEnabled(false),
							mDevicesHaveChanged(false) { }
	~CSoundSourcePrefs() { }
	CSoundSourcePrefs& operator=(const CSoundSourcePrefs& source);

private:

	void Erase(int index);
	bool Insert(int index, DevicePrefs * sdi);
	void Remove(int index);
	void RemoveDupes();

	int FindByName(wstring & name);

	template<class Archive>
	void load(Archive & ar, const unsigned int version);
	template<class Archive>
	void save(Archive & ar, const unsigned int version) const;

	DevicePrefs * mDevices;
	int mMax;
	int mNext;
	//int mDefault;
	UINT mKeyCode;
	UINT mKeyMods;
	wstring mCycleKeyString;
	bool mIsCycleKeyEnabled;
	bool mDevicesHaveChanged; // not used
	static const int mLimit = 16;

	BOOST_SERIALIZATION_SPLIT_MEMBER()
};

BOOST_CLASS_VERSION(CSoundSourcePrefs, 1)
BOOST_CLASS_TRACKING( CSoundSourcePrefs, boost::serialization::track_never )
