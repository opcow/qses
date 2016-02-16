#include <iostream>
#include <fstream>
#include <sstream>
#include <Shlobj.h>
#include <string>
#include <windows.h>
#include <wchar.h>

#include "Prefs.h"

using namespace std;

CSoundSourcePrefs::CSoundSourcePrefs( const CSoundSourcePrefs& other ) :
	mMax(other.mMax),
	mNext(other.mNext),
//	mDefault(other.mDefault),
	mKeyCode(other.mKeyCode),
	mKeyMods(other.mKeyMods),
	mCycleKeyString(other.mCycleKeyString),
	mIsCycleKeyEnabled(other.mIsCycleKeyEnabled)
{
	mDevices = new DevicePrefs [mMax];
	for (int i = 0; i < mNext; i++)
		mDevices[i] = other.mDevices[i];
}

CSoundSourcePrefs& CSoundSourcePrefs::operator=( const CSoundSourcePrefs& source )
{

	mMax = source.mMax;
	mNext = source.mNext;
//	mDefault = source.mDefault;
	mKeyCode = source.mKeyCode;
	mKeyMods = source.mKeyMods;
	mCycleKeyString = source.mCycleKeyString;
	mIsCycleKeyEnabled = source.mIsCycleKeyEnabled;

	mDevices = new DevicePrefs [mMax];
	for (int i = 0; i < mNext; i++)
		mDevices[i] = source.mDevices[i];

	return *this;
}

void CSoundSourcePrefs::Add(DevicePrefs *sdi/*, bool def*/)
{
	int index = FindByName(sdi->Name);

	if (index != -1)
	{
		mDevices[index] = *sdi;
		mDevices[index].IsPresent = true;
		time(&mDevices[index].TimeStamp);
		//if (def)
		//	mDefault = index;
	}
	else
	{
		if (mNext == mMax)
			Remove(0);

		mDevices[mNext] = *sdi;
		mDevices[mNext].IsPresent = true;
		time(&mDevices[mNext].TimeStamp);
//		if (def)
//			mDefault = mNext;
		mNext++;
	}
}


void CSoundSourcePrefs::Update(DevicePrefs *sdi/*, bool def*/)
{
	int index = FindByName(sdi->Name);

	if (index != -1)
	{
		mDevices[index].DeviceID = sdi->DeviceID;
		mDevices[index].Name = sdi->Name;
		mDevices[index].IsPresent = true;
		time(&mDevices[index].TimeStamp);
		//if (def)
		//	mDefault = index;
	}
	else
	{
		if (mNext == mMax)
			Remove(0);  //fixme

		mDevices[mNext] = *sdi;
		mDevices[mNext].IsPresent = true;
		time(&mDevices[mNext].TimeStamp);
		//if (def)
		//	mDefault = mNext;
		mNext++;
	}
}

bool CSoundSourcePrefs::SetMax(int max)
{
	if (max < mNext || max > mLimit)
		return false;
	mMax = max;
	return true;
}

wstring& CSoundSourcePrefs::GetHotkeyString(int index, wstring& keyString)
{
	if (index > mNext)
		keyString.clear();

	return keyString = mDevices[index].HotkeyString;
};

void CSoundSourcePrefs::SetHotkeyCode(int index, UINT code)
{
	if (index < mNext) mDevices[index].KeyCode = code;

	if (mDevices[index].KeyCode == 0 && mDevices[index].KeyMods == 0)
		mDevices[index].HasHotkey = false;
}

void CSoundSourcePrefs::SetHotkeyMods(int index, UINT mods)
{
	if (index < mNext) mDevices[index].KeyMods = mods;

	if (mDevices[index].KeyCode == 0 && mDevices[index].KeyMods == 0)
		mDevices[index].HasHotkey = false;
}

UINT CSoundSourcePrefs::GetHotkeyMods(int index)
{
	if (index < mNext)
		return mDevices[index].KeyMods;
	else return 0;
}


bool CSoundSourcePrefs::GetHotkeyEnabled(int index)
{ 
	if (index < mNext)
		return mDevices[index].HasHotkey;
	return false;
}

bool CSoundSourcePrefs::GetExcludeFromCycle(int index)
{ 
	if (index < mNext)
		return mDevices[index].IsExcludedFromCycle;
	return false;
}

bool CSoundSourcePrefs::GetIsHidden(int index)
{ 
	if (index < mNext)
		return mDevices[index].IsHidden;
	return false;
}

bool CSoundSourcePrefs::GetIsPresent(int index)
{ 
	if (index < mNext)
		return mDevices[index].IsPresent;
	return false;
}

void CSoundSourcePrefs::Swap(int d1, int d2)
{ 
	DevicePrefs temp = mDevices[d1];
	mDevices[d1] = mDevices[d2];
	mDevices[d2] = temp;
}

void CSoundSourcePrefs::Sort()
{
	for (int i = 0 ; i < mNext - 1; i++)
		if (mDevices[i].IsPresent == false)
		{
			for(int j = i + 1; j < mNext; j++)
			{
				if (mDevices[j].IsPresent)
					Swap(i, j);
			}
		}
}


bool CSoundSourcePrefs::GetByName(wstring& name, DevicePrefs * sdi)
{
	int index = FindByName(name);

	if(index != -1)
	{
		*sdi = mDevices[index];
		return true;
	}
	return false;
}

int CSoundSourcePrefs::FindByID(const wstring& id)
{
	for (int i = 0; i < mNext; i++)
	{
		if (id == mDevices[i].DeviceID)
			return i;
	}
	return -1;
}

wstring& CSoundSourcePrefs::GetName(int index, wstring& s)
{
	if (index < mNext)
		s = mDevices[index].Name;
	return s;
}

wstring& CSoundSourcePrefs::GetID(int index, wstring& s)
{
	if (index < mNext)
		s = mDevices[index].DeviceID;
	return s;
}

bool CSoundSourcePrefs::Init(int max)
{
	if (max < 1 || max > mLimit)
		return false;

	if (mDevices == 0)
		delete mDevices;

	mDevices = new DevicePrefs[max];

	if (mDevices == 0)
		return false;

	mMax = max;

	for (int i = 0; i < mMax; i++)
		Erase(i);

	return true;
}

void CSoundSourcePrefs::Remove(int index)
{
	if (index >= mNext)
		return;

	if (index < (mNext - 1))
	{
		for (int i = index; i < (mNext - 1); i++)
			mDevices[i] = mDevices[i+1];
	}
	Erase(mNext--);
}

void CSoundSourcePrefs::RemoveDupes()
{
	for (int i = 0 ; i < mNext - 1; i++)
		for (int j = i + 1; j < mNext; j++)
			if (mDevices[i].DeviceID == mDevices[j].DeviceID)
				Remove(j);
}

void CSoundSourcePrefs::Clear(int index)
{
	mDevices[index].HotkeyString.clear();
	mDevices[index].KeyCode = 0;
	mDevices[index].KeyMods = 0;
	mDevices[index].HasHotkey = false;
	mDevices[index].IsExcludedFromCycle = false;
	mDevices[index].IsPresent = false;
	mDevices[index].IsHidden = false;
}

void CSoundSourcePrefs::Erase(int index)
{
	mDevices[index].Name.clear();
	mDevices[index].DeviceID.clear();
	mDevices[index].HotkeyString.clear();
	mDevices[index].KeyCode = 0;
	mDevices[index].KeyMods = 0;
	mDevices[index].HasHotkey = false;
	mDevices[index].IsExcludedFromCycle = false;
	mDevices[index].IsPresent = false;
	mDevices[index].IsHidden = false;
	mDevices[index].TimeStamp = 0;
}

void CSoundSourcePrefs::ResetPresent()
{
	for (int i = 0; i < mMax; i++)
		mDevices[i].IsPresent = false;
}

//bool CSoundSourcePrefs::Insert(int index, DevicePrefs * sdi)
//{
//	if (mNext == mMax || index > mNext)
//		return false;
//
//	for (int i = mNext; i >= index; i--)
//		mDevices[i] = mDevices[i-1];
//	mDevices[index] = *sdi;
//
//	mNext++;
//
//	return true;
//}

int CSoundSourcePrefs::FindByName(wstring& name)
{
	if (mNext != 0)
	{
		for (int i = 0; i < mNext; i++)
		{
			if (name == mDevices[i].Name)
			{
				return i = i;
			}
		}
	}
	return -1;
}

template<class Archive>
void CSoundSourcePrefs::load(Archive & ar, const unsigned int version)
{
	if (version < 1)
		return;
	ar & mNext;
	ar & mKeyCode;
	ar & mKeyMods;
	ar & mCycleKeyString;
	ar & mIsCycleKeyEnabled;

	for (int i = 0; i < mNext; i++)
		ar & mDevices[i];
}
template<class Archive>
void CSoundSourcePrefs::save(Archive & ar, const unsigned int version) const
{
	ar & mNext;
	ar & mKeyCode;
	ar & mKeyMods;
	ar & mCycleKeyString;
	ar & mIsCycleKeyEnabled;

	for (int i = 0; i < mNext; i++)
		ar & mDevices[i];
}

bool CSoundSourcePrefs::Save()
{
	wstring FilePath;
	LPWSTR wstrPrefsPath = 0;

	RemoveDupes();

	SHGetKnownFolderPath(FOLDERID_RoamingAppData, 0, 0, &wstrPrefsPath);

	FilePath = wstrPrefsPath;
	FilePath += L"\\Quick Sound Source";
	CreateDirectoryW(FilePath.data(), 0);
	FilePath += L"\\prefs.ini";

	wofstream output(FilePath.data());
//	wstringstream ss;
	try
	{
		if (!output.is_open())
			throw 1;
//		boost::archive::text_woarchive oa(ss);
		boost::archive::text_woarchive oa(output);
		oa << (*this);
	}
	catch (...) { return false; }

	return true;
}

bool CSoundSourcePrefs::Load()
{
	wstring FilePath;
	LPWSTR wstrPrefsPath = 0;

	SHGetKnownFolderPath(FOLDERID_RoamingAppData, 0, 0, &wstrPrefsPath);

	FilePath = wstrPrefsPath;
	FilePath += L"\\Quick Sound Source";
	BOOL bUsedDefChar = CreateDirectoryW(FilePath.data(), 0);
	FilePath += L"\\prefs.ini";

	std::wifstream input(FilePath.data());
	try
	{
		if (!input.is_open())
			throw 1;
		boost::archive::text_wiarchive ia(input);
		ia >> (*this);
	}
	catch (...) { return false; }

	return true;
}
