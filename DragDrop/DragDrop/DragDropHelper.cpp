
#include "stdafx.h"
#include "DragDropHelper.h"

extern UINT g_uCustomClipbrdFormats;

/*static*/ bool CDragDropHelper::IsRegisteredFormat(CLIPFORMAT cfFormat, LPCTSTR lpszFormat, bool bRegister /*= false*/)
{
	bool bRet = false;
	if (cfFormat >= 0xC000)
	{
		if (bRegister)
			bRet = (cfFormat == RegisterFormat(lpszFormat));
		else
		{
			TCHAR lpszName[128];
			if (::GetClipboardFormatName(cfFormat, lpszName, sizeof(lpszName) / sizeof(lpszName[0])))
				bRet = (0 == _tcsicmp(lpszFormat, lpszName));
		}
	}
	return bRet;
}

/*static*/ DROPIMAGETYPE CDragDropHelper::DropEffectToDropImage(DROPEFFECT dwEffect)
{
	DROPIMAGETYPE nImageType = DROPIMAGE_INVALID;
	dwEffect &= ~DROPEFFECT_SCROLL;
	if (DROPEFFECT_NONE == dwEffect)
		nImageType = DROPIMAGE_NONE;
	else if (dwEffect & DROPEFFECT_MOVE)
		nImageType = DROPIMAGE_MOVE;
	else if (dwEffect & DROPEFFECT_COPY)
		nImageType = DROPIMAGE_COPY;
	else if (dwEffect & DROPEFFECT_LINK)
		nImageType = DROPIMAGE_LINK;
	return nImageType;
}

/*static*/ HGLOBAL CDragDropHelper::CopyGlobalMemory(HGLOBAL hDst, HGLOBAL hSrc, size_t nSize)
{
	ASSERT(hSrc);

	if (0 == nSize)
		nSize = ::GlobalSize(hSrc);
	if (nSize)
	{
		if (NULL == hDst)
			hDst = ::GlobalAlloc(GMEM_MOVEABLE, nSize);
		else if (nSize > ::GlobalSize(hDst))
			hDst = NULL;
		if (hDst)
		{
			::CopyMemory(::GlobalLock(hDst), ::GlobalLock(hSrc), nSize);
			::GlobalUnlock(hDst);
			::GlobalUnlock(hSrc);
		}
	}
	return hDst;
}

/*static*/ bool CDragDropHelper::GetGloablDropFileListData(IN COleDataObject* pIDataObj, OUT CStringList& list)
{
	HDROP       hdrop	= NULL;
	HGLOBAL		hg		= NULL;
	UINT        uNumFiles = 0;
	TCHAR       szNextFile[MAX_PATH] = {0};

	hg = pIDataObj->GetGlobalData(CF_HDROP);

	if (NULL == hg)
		return false;

	hdrop = (HDROP)GlobalLock(hg);

	if (NULL == hdrop)
	{
		GlobalUnlock(hg);
		return false;
	}

	uNumFiles = DragQueryFile(hdrop, -1, NULL, 0);

	for (UINT uFile = 0; uFile < uNumFiles; uFile++)
	{
		if (DragQueryFile(hdrop, uFile, szNextFile, MAX_PATH) > 0)
		{
			list.AddTail(szNextFile);
		}
	}

	GlobalUnlock(hg);
	return true;
}

/*static*/ bool CDragDropHelper::GetGlobalData(LPDATAOBJECT pIDataObj, LPCTSTR lpszFormat, FORMATETC& FormatEtc, STGMEDIUM& StgMedium)
{
	ASSERT(pIDataObj);
	ASSERT(lpszFormat && *lpszFormat);

	bool bRet = false;
	FormatEtc.cfFormat = RegisterFormat(lpszFormat);
	FormatEtc.ptd = NULL;
	FormatEtc.dwAspect = DVASPECT_CONTENT;
	FormatEtc.lindex = -1;
	FormatEtc.tymed = TYMED_HGLOBAL;
	if (SUCCEEDED(pIDataObj->QueryGetData(&FormatEtc)))
	{
		if (SUCCEEDED(pIDataObj->GetData(&FormatEtc, &StgMedium)))
		{
			bRet = (TYMED_HGLOBAL == StgMedium.tymed);
			if (!bRet)
				::ReleaseStgMedium(&StgMedium);
		}
	}
	return bRet;
}

/*static*/ DWORD CDragDropHelper::GetGlobalDataDWord(LPDATAOBJECT pIDataObj, LPCTSTR lpszFormat)
{
	DWORD dwData = 0;
	FORMATETC FormatEtc;
	STGMEDIUM StgMedium;
	if (GetGlobalData(pIDataObj, lpszFormat, FormatEtc, StgMedium))
	{
		ASSERT(::GlobalSize(StgMedium.hGlobal) >= sizeof(DWORD));
		dwData = *(static_cast<LPDWORD>(::GlobalLock(StgMedium.hGlobal)));
		::GlobalUnlock(StgMedium.hGlobal);
		::ReleaseStgMedium(&StgMedium);
	}
	return dwData;
}

/*static*/ bool CDragDropHelper::SetGlobalDataDWord(LPDATAOBJECT pIDataObj, LPCTSTR lpszFormat, DWORD dwData, bool bForce /*= true*/)
{
	bool bSet = false;
	FORMATETC FormatEtc;
	STGMEDIUM StgMedium;
	if (GetGlobalData(pIDataObj, lpszFormat, FormatEtc, StgMedium))
	{
		LPDWORD pData = static_cast<LPDWORD>(::GlobalLock(StgMedium.hGlobal));
		bSet = (*pData != dwData);
		*pData = dwData;
		::GlobalUnlock(StgMedium.hGlobal);
		if (bSet)
			bSet = SUCCEEDED(pIDataObj->SetData(&FormatEtc, &StgMedium, TRUE));
		if (!bSet)
			::ReleaseStgMedium(&StgMedium);
	}
	else if (dwData || bForce)
	{
		StgMedium.hGlobal = ::GlobalAlloc(GMEM_MOVEABLE, sizeof(DWORD));
		if (StgMedium.hGlobal)
		{
			LPDWORD pData = static_cast<LPDWORD>(::GlobalLock(StgMedium.hGlobal));
			*pData = dwData;
			::GlobalUnlock(StgMedium.hGlobal);
			StgMedium.tymed = TYMED_HGLOBAL;
			StgMedium.pUnkForRelease = NULL;
			bSet = SUCCEEDED(pIDataObj->SetData(&FormatEtc, &StgMedium, TRUE));
			if (!bSet)
				::GlobalFree(StgMedium.hGlobal);
		}
	}
	return bSet;
}

/*static*/ bool CDragDropHelper::ClearDescription(DROPDESCRIPTION *pDropDescription)
{
	ASSERT(pDropDescription);

	bool bChanged = 
		pDropDescription->type != DROPIMAGE_INVALID ||
		pDropDescription->szMessage[0] != L'\0' || 
		pDropDescription->szInsert[0] != L'\0'; 
	pDropDescription->type = DROPIMAGE_INVALID;
	pDropDescription->szMessage[0] = pDropDescription->szInsert[0] = L'\0';
	return bChanged;
}

bool CDragDropHelper::IsHaveFlag(COleDataObject* pIDataObj)
{
	
	if (pIDataObj->IsDataAvailable(g_uCustomClipbrdFormats))
	{
		return true;
	}

	return false;
}

bool CDragDropHelper::IsHaveFlag(LPDATAOBJECT pIDataObj)
{
	STGMEDIUM stg = { 0 };
	FORMATETC etc = { g_uCustomClipbrdFormats, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	if (SUCCEEDED(pIDataObj->GetData(&etc, &stg)))
	{
		if (stg.tymed == TYMED_HGLOBAL && stg.hGlobal) {
			return true;
		}
		ReleaseStgMedium(&stg);
	}
	return false;
}

CDropDescription::CDropDescription()
{
	m_lpszInsert = NULL;
	::ZeroMemory(m_pDescriptions, sizeof(m_pDescriptions));
	::ZeroMemory(m_pDescriptions1, sizeof(m_pDescriptions1));
}

CDropDescription::~CDropDescription()
{
	delete [] m_lpszInsert;
	for (unsigned n = 0; n <= nMaxDropImage; n++)
	{
		delete [] m_pDescriptions[n];
		delete [] m_pDescriptions1[n];
	}
}

bool CDropDescription::SetText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1)
{
	ASSERT(IsValidType(nType));

	bool bRet = false;
	if (IsValidType(nType))
	{
		delete [] m_pDescriptions[nType];
		delete [] m_pDescriptions1[nType];
		m_pDescriptions[nType] = m_pDescriptions1[nType] = NULL;
		if (lpszText)
		{
			m_pDescriptions[nType] = new WCHAR[wcslen(lpszText) + 1];
			wcscpy_s(m_pDescriptions[nType], wcslen(lpszText) + 1, lpszText);
			bRet = true;
		}
		if (lpszText1)
		{
			m_pDescriptions1[nType] = new WCHAR[wcslen(lpszText1) + 1];
			wcscpy_s(m_pDescriptions1[nType], wcslen(lpszText1) + 1, lpszText1);
			bRet = true;
		}
	}
	return bRet;
}

bool CDropDescription::SetInsert(LPCWSTR lpszInsert)
{
	delete [] m_lpszInsert;
	m_lpszInsert = NULL;
	if (lpszInsert)
	{
		m_lpszInsert = new WCHAR[wcslen(lpszInsert) + 1];
		wcscpy_s(m_lpszInsert, wcslen(lpszInsert) + 1, lpszInsert);
	}
	return NULL != m_lpszInsert;
}

LPCWSTR CDropDescription::GetText(DROPIMAGETYPE nType, bool b1) const
{
	return IsValidType(nType) ? 
		((b1 && m_pDescriptions1[nType]) ? m_pDescriptions1[nType] : m_pDescriptions[nType]) : NULL;
}

bool CDropDescription::SetDescription(DROPDESCRIPTION *pDropDescription, LPCWSTR lpszMsg) const
{
	return CopyInsert(pDropDescription, m_lpszInsert) | CopyMessage(pDropDescription, lpszMsg);
}

bool CDropDescription::SetDescription(DROPDESCRIPTION *pDropDescription, DROPIMAGETYPE nType) const
{
	return SetDescription(pDropDescription, GetText(nType, HasInsert()));
}

bool CDropDescription::CopyText(DROPDESCRIPTION *pDropDescription, DROPIMAGETYPE nType) const
{
	return CopyMessage(pDropDescription, GetText(nType, HasInsert(pDropDescription)));
}

bool CDropDescription::CopyMessage(DROPDESCRIPTION *pDropDescription, LPCWSTR lpszMsg) const
{
	ASSERT(pDropDescription);
	return CopyDescription(pDropDescription->szMessage, sizeof(pDropDescription->szMessage) / sizeof(pDropDescription->szMessage[0]), lpszMsg);
}

bool CDropDescription::CopyInsert(DROPDESCRIPTION *pDropDescription, LPCWSTR lpszInsert) const
{
	ASSERT(pDropDescription);
	return CopyDescription(pDropDescription->szInsert, sizeof(pDropDescription->szInsert) / sizeof(pDropDescription->szInsert[0]), lpszInsert);
}


bool CDropDescription::CopyDescription(LPWSTR lpszDest, size_t nDstSize, LPCWSTR lpszSrc) const
{
	ASSERT(lpszDest);
	ASSERT(nDstSize > 0);

	bool bChanged = false;
	if (lpszSrc && *lpszSrc)
	{
		if (wcscmp(lpszDest, lpszSrc))
		{
			bChanged = true;
			wcsncpy_s(lpszDest, nDstSize, lpszSrc, _TRUNCATE);
		}
	}
	else if (lpszDest[0])
	{
		bChanged = true;
		lpszDest[0] = L'\0';
	}
	return bChanged;
}