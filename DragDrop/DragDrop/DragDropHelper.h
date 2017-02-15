

#pragma once

/*!
 * @file DragDropHelper.h
 *
 * @author XUECAT
 * @date ÎåÔÂ 2016
 *
 * 
 */

#include <ShlObj.h>
#include <ShObjIdl.h>

class CDragDropHelper
{
public:
	static inline CLIPFORMAT RegisterFormat(LPCTSTR lpszFormat)
	{ return static_cast<CLIPFORMAT>(::RegisterClipboardFormat(lpszFormat)); }

	static unsigned PopCount(unsigned n);
	static bool		IsRegisteredFormat(CLIPFORMAT cfFormat, LPCTSTR lpszFormat, bool bRegister = false);
	static HGLOBAL	CopyGlobalMemory(HGLOBAL hDst, HGLOBAL hSrc, size_t nSize);

	static bool		GetGloablDropFileListData(IN COleDataObject* pIDataObj, OUT CStringList& list);
	static bool		GetGlobalData(LPDATAOBJECT pIDataObj, LPCTSTR lpszFormat, FORMATETC& FormatEtc, STGMEDIUM& StgMedium);
	static DWORD	GetGlobalDataDWord(LPDATAOBJECT pIDataObj, LPCTSTR lpszFormat);
	static bool		SetGlobalDataDWord(LPDATAOBJECT pIDataObj, LPCTSTR lpszFormat, DWORD dwData, bool bForce = true);

	static DROPIMAGETYPE DropEffectToDropImage(DROPEFFECT dwEffect);
	static bool		ClearDescription(DROPDESCRIPTION *pDropDescription);
	static bool		IsHaveFlag(COleDataObject* pIDataObj);
	static bool		IsHaveFlag(LPDATAOBJECT pIDataObj);
};


class CDropDescription 
{
public:
	CDropDescription();
	virtual ~CDropDescription();

	static const DROPIMAGETYPE nMaxDropImage = DROPIMAGE_NOIMAGE;

	inline bool IsValidType(DROPIMAGETYPE nType) const
	{ return (nType >= DROPIMAGE_NONE && nType <= nMaxDropImage); }

	bool	SetInsert(LPCWSTR lpszInsert);
	bool	SetText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1);
	LPCWSTR	GetText(DROPIMAGETYPE nType, bool b1) const;
	bool	SetDescription(DROPDESCRIPTION *pDropDescription, DROPIMAGETYPE nType) const;
	bool	SetDescription(DROPDESCRIPTION *pDropDescription, LPCWSTR lpszMsg) const;
	bool	CopyText(DROPDESCRIPTION *pDropDescription, DROPIMAGETYPE nType) const;
	bool	CopyMessage(DROPDESCRIPTION *pDropDescription, LPCWSTR lpszMsg) const;
	bool	CopyInsert(DROPDESCRIPTION *pDropDescription, LPCWSTR lpszInsert) const;

	inline LPCWSTR GetInsert() const { return m_lpszInsert; } 
	inline bool	CopyInsert(DROPDESCRIPTION *pDropDescription) const 
	{ return CopyInsert(pDropDescription, m_lpszInsert); }

	inline bool HasText(DROPIMAGETYPE nType) const
	{ return IsValidType(nType) && NULL != m_pDescriptions[nType]; }

	inline bool IsTextEmpty(DROPIMAGETYPE nType) const
	{ return !HasText(nType) || L'\0' == m_pDescriptions[nType][0]; }


	inline bool HasInsert() const { return NULL != m_lpszInsert; }
	inline bool IsInsertEmpty() const { return !HasInsert() || L'\0' == *m_lpszInsert; }

	inline bool HasInsert(const DROPDESCRIPTION *pDropDescription) const
	{ return pDropDescription && pDropDescription->szInsert[0] != L'\0'; }

	inline bool	ClearDescription(DROPDESCRIPTION *pDropDescription) const
	{ return CDragDropHelper::ClearDescription(pDropDescription); }

protected:
	bool CopyDescription(LPWSTR lpszDest, size_t nDstSize, LPCWSTR lpszSrc) const;

	LPWSTR m_lpszInsert;
	LPWSTR m_pDescriptions[nMaxDropImage + 1];
	LPWSTR m_pDescriptions1[nMaxDropImage + 1];
};
