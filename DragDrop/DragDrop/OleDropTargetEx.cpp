#include "stdafx.h"
#include "OleDropTargetEx.h"

COleDropTargetEx::COleDropTargetEx()
	: m_pDropTargetHelper(NULL)
	, m_pTargetFun(NULL)
	, m_bAcceptExplorer(false)
	, m_bCanShowDescription(false)
	, m_bUseDropDescription(false)
	, m_bDescriptionUpdated(false)
	, m_bHasFlag(false)
	, m_bEntered(false)
	, m_bHasDragImage(false)
	, m_bTextAllowed(false)
	, m_nAllowEffect(DROPEFFECT_NONE)
{
#if defined(IID_PPV_ARGS)
	::CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pDropTargetHelper));
#else
	::CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER,
		IID_IDropTargetHelper, reinterpret_cast<LPVOID*>(&m_pDropTargetHelper));
#endif
	m_bCanShowDescription = (m_pDropTargetHelper != NULL);
	if (m_bCanShowDescription)
	{
#if (WINVER < 0x0600)
		OSVERSIONINFOEX VerInfo;
		DWORDLONG dwlConditionMask = 0;
		::ZeroMemory(&VerInfo, sizeof(OSVERSIONINFOEX));
		VerInfo.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
		VerInfo.dwMajorVersion = 6;
		VER_SET_CONDITION(dwlConditionMask, VER_MAJORVERSION, VER_GREATER_EQUAL);
		m_bCanShowDescription = ::VerifyVersionInfo(&VerInfo, VER_MAJORVERSION, dwlConditionMask) && ::IsAppThemed();
#else
		m_bCanShowDescription = (0 != ::IsAppThemed());
#endif
	} 
}

COleDropTargetEx::~COleDropTargetEx()
{
	if (NULL != m_pDropTargetHelper)
		m_pDropTargetHelper->Release();
}

void COleDropTargetEx::ClearStateFlags()
{
	m_bDescriptionUpdated = false;
	m_bEntered = false;
	m_bHasDragImage = false;
	m_bTextAllowed = false;
	m_bHasFlag = false;
}

BOOL COleDropTargetEx::Register(CWnd* pWnd, TargetFun* pFun /* = NULL */)
{
	if (pFun == NULL) {
		pFun = dynamic_cast<TargetFun*>(pWnd);
	}
	
	if (pFun != NULL) 
	{
		m_pTargetFun = pFun;
	}

	if (pWnd != NULL) {
		return __super::Register(pWnd);
	}

	return FALSE;
}

bool COleDropTargetEx::SetDropDescriptionText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1, LPCWSTR lpszInsert /*= NULL*/)
{
	bool bRet = false;
	if (m_bCanShowDescription)
	{
		bRet = m_DropDescription.SetText(nType, lpszText, lpszText1);
		if (bRet && lpszInsert)
			m_DropDescription.SetInsert(lpszInsert);
		m_bUseDropDescription |= bRet;
	}
	return bRet;
}


bool COleDropTargetEx::SetDropInsertText(LPCWSTR lpszText)
{
	return m_bCanShowDescription && m_DropDescription.SetInsert(lpszText);
}

bool COleDropTargetEx::SetDropDescription(DROPIMAGETYPE nImageType, LPCWSTR lpszText, bool bCreate)
{
	ASSERT(m_lpDataObject);

	bool bHasDescription = false;

	if (m_bHasDragImage && m_bCanShowDescription && m_lpDataObject)
	{
		STGMEDIUM StgMedium;
		FORMATETC FormatEtc;
		if (CDragDropHelper::GetGlobalData(m_lpDataObject, CFSTR_DROPDESCRIPTION, FormatEtc, StgMedium))
		{
			bHasDescription = true;
			bool bUpdate = false;
			DROPDESCRIPTION *pDropDescription = static_cast<DROPDESCRIPTION*>(::GlobalLock(StgMedium.hGlobal));
			if (DROPIMAGE_INVALID == nImageType)
				bUpdate = m_DropDescription.ClearDescription(pDropDescription);
			else if (m_bTextAllowed)
			{
				if (NULL == lpszText)
					lpszText = m_DropDescription.GetText(nImageType, m_DropDescription.HasInsert());
				bUpdate = m_DropDescription.SetDescription(pDropDescription, lpszText);
			}
			if (pDropDescription->type != nImageType)
			{
				bUpdate = true;
				pDropDescription->type = nImageType;
			}
			::GlobalUnlock(StgMedium.hGlobal);
			if (bUpdate)
				bUpdate = SUCCEEDED(m_lpDataObject->SetData(&FormatEtc, &StgMedium, TRUE));
			if (!bUpdate)
				::ReleaseStgMedium(&StgMedium);
		}
		if (!bHasDescription && bCreate &&
			DROPIMAGE_INVALID != nImageType &&
			(m_bTextAllowed || nImageType > DROPIMAGE_LINK))
		{
			StgMedium.tymed = TYMED_HGLOBAL;
			StgMedium.hGlobal = ::GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, sizeof(DROPDESCRIPTION));
			StgMedium.pUnkForRelease = NULL;
			if (StgMedium.hGlobal)
			{
				DROPDESCRIPTION *pDropDescription = static_cast<DROPDESCRIPTION*>(::GlobalLock(StgMedium.hGlobal));
				pDropDescription->type = nImageType;
				if (m_bTextAllowed)
				{
					if (NULL == lpszText)
						lpszText = m_DropDescription.GetText(nImageType, m_DropDescription.HasInsert());
					m_DropDescription.SetDescription(pDropDescription, lpszText);
				}
				::GlobalUnlock(StgMedium.hGlobal);
				bHasDescription = SUCCEEDED(m_lpDataObject->SetData(&FormatEtc, &StgMedium, TRUE));
				if (!bHasDescription)
					::GlobalFree(StgMedium.hGlobal);
			}
		}
	}
	m_bDescriptionUpdated = true;
	return bHasDescription;
}

bool COleDropTargetEx::SetDropDescription(DROPEFFECT dwEffect)
{
	bool bHasDescription = false;
	m_bHasDragImage = (0 != CDragDropHelper::GetGlobalDataDWord(m_lpDataObject, _T("DragWindow")));
	if (m_bHasDragImage && m_bCanShowDescription)
		m_bTextAllowed = (DSH_ALLOWDROPDESCRIPTIONTEXT == CDragDropHelper::GetGlobalDataDWord(m_lpDataObject, _T("DragSourceHelperFlags")));

	if (m_bTextAllowed)								
	{
		DROPIMAGETYPE nType = static_cast<DROPIMAGETYPE>(dwEffect);
		LPCWSTR lpszText = m_DropDescription.GetText(nType, m_DropDescription.HasInsert());
		if (lpszText)
			bHasDescription = SetDropDescription(nType, lpszText, true);	
		else
			bHasDescription = ClearDropDescription();
	}
	m_bDescriptionUpdated = true;
	return bHasDescription;
}

DROPEFFECT COleDropTargetEx::OnDragEnter(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point)
{
	DROPEFFECT dwRet = DROPEFFECT_NONE;
	ClearStateFlags();

	if (m_pTargetFun)
	{
		dwRet = m_pTargetFun->OnDragEnterFunc(pWnd, pDataObject, dwKeyState, point);
	}
	else
	{
		dwRet = COleDropTarget::OnDragEnter(pWnd, pDataObject, dwKeyState, point);
	}

	if (m_bAcceptExplorer || CDragDropHelper::IsHaveFlag(pDataObject))
	{
		m_bHasFlag = true;

		dwRet |= m_nAllowEffect;
	}

	if (m_bUseDropDescription && !m_bDescriptionUpdated)
		SetDropDescription(dwRet);

	if (m_pDropTargetHelper)
		m_pDropTargetHelper->DragEnter(pWnd->GetSafeHwnd(), m_lpDataObject, &point, dwRet);
	m_bEntered = true;
	return dwRet;
}

void COleDropTargetEx::OnDragLeave(CWnd* pWnd)
{
	if (m_pTargetFun)
		m_pTargetFun->OnDragLeaveFunc(pWnd);
	else
		COleDropTarget::OnDragLeave(pWnd);

	ClearDropDescription();

	if (m_pDropTargetHelper)
		m_pDropTargetHelper->DragLeave();

	ClearStateFlags();
}


DROPEFFECT COleDropTargetEx::OnDragOver(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point)
{
	DROPEFFECT dwRet = DROPEFFECT_NONE;	
	m_bDescriptionUpdated = false;

	if (m_pTargetFun)
	{
		dwRet = m_pTargetFun->OnDragOverFunc(pWnd, pDataObject, dwKeyState, point);
	}
	else
	{
		dwRet = COleDropTarget::OnDragOver(pWnd, pDataObject, dwKeyState, point);
	}

	if (m_bEntered && m_bHasFlag) {
		dwRet |= m_nAllowEffect;
	}
	
	if (m_bUseDropDescription && !m_bDescriptionUpdated)
		SetDropDescription(dwRet);

	if (m_pDropTargetHelper)
		m_pDropTargetHelper->DragOver(&point, dwRet);
	return dwRet;
}


DROPEFFECT COleDropTargetEx::OnDropEx(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT dropList, CPoint point)
{           
	return __super::OnDropEx(pWnd, pDataObject, dropDefault, dropList, point);

	DROPEFFECT dwRet = (DROPEFFECT)-1;
	if (m_pTargetFun)
	{
		dwRet = m_pTargetFun->OnDropFunc(pWnd, pDataObject, dropDefault, point);
	}
	else
	{
		dwRet = COleDropTarget::OnDropEx(pWnd, pDataObject, dropDefault, dropList, point);
	}

	if (m_bEntered && m_bHasFlag) {
		dwRet |= m_nAllowEffect;
	}
	
	if ((DROPEFFECT)-1 != dwRet)
		//OnPostDrop(pDataObject, dwRet, point);

	return dwRet;
}


BOOL COleDropTargetEx::OnDrop(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropEffect, CPoint point)
{           
	BOOL bRet = FALSE;

	if (m_pTargetFun)
	{
		bRet = m_pTargetFun->OnDropFunc(pWnd, pDataObject, dropEffect, point);
	}
	else
	{
		bRet = COleDropTarget::OnDrop(pWnd, pDataObject, dropEffect, point);
	}

	if (m_bEntered && m_bHasFlag) {
		dropEffect |= m_nAllowEffect;
	}
	//OnPostDrop(pDataObject, dropEffect, point);

	if (m_pDropTargetHelper)
		m_pDropTargetHelper->Drop(m_lpDataObject, &point, dropEffect);
	
	return bRet;
}

void COleDropTargetEx::OnPostDrop(COleDataObject* pDataObject, DROPEFFECT dropEffect, CPoint point)
{
	if (m_bEntered && m_bHasFlag)
	{
		CDragDropHelper::SetGlobalDataDWord(m_lpDataObject, CFSTR_PERFORMEDDROPEFFECT, dropEffect);
	}
}

INT COleDropTargetEx::IsDoScroll(CWnd* pWnd, DWORD dwKeyState, CPoint point)
{
	CRect rect;

	pWnd->GetClientRect(rect);
	

	if (m_bHasFlag && pWnd->IsKindOf(RUNTIME_CLASS(CListCtrl))) 
	{

		CListCtrl* pList = dynamic_cast<CListCtrl*>(pWnd);

		CRect header;
		if (pList && pList->GetHeaderCtrl()) {
			pList->GetHeaderCtrl()->GetWindowRect(header);

			rect.DeflateRect(0, header.Height(), 0, 0);
			rect.DeflateRect(10, 10);
		}
		else
			return FALSE;
		

		if (point.y < rect.top) 
		{
			return SCROLL_UP;
		}
		else if (point.x < rect.left) 
		{
			return SCROLL_LEFT;
		}
		else if (point.y > rect.bottom) 
		{
			return SCROLL_DOWN;
		}
		else if (point.x > rect.right) 
		{
			return SCROLL_RIGHT;
		}
	}

	return FALSE;
}

//这个函数在 OnDragEnter OnDragOver之前被调用，
//如果设置了 dwEffect DROPEFFECT_SCROLL属性，就不会调用那俩个函数
DROPEFFECT COleDropTargetEx::OnDragScroll(CWnd* pWnd, DWORD dwKeyState, CPoint point)
{
	DROPEFFECT dwEffect = DROPEFFECT_NONE;
	
	INT ndirect = IsDoScroll(pWnd, dwKeyState, point);

	if (ndirect != FALSE) {
		dwEffect = DROPEFFECT_SCROLL;
	}

	if (m_pTargetFun && (dwEffect & DROPEFFECT_SCROLL))
	{
		dwEffect = m_pTargetFun->OnDragScrollFunc(pWnd, dwKeyState, point, ndirect);
	}
	else
	{
		dwEffect = COleDropTarget::OnDragScroll(pWnd, dwKeyState, point);
	}
	
	if ((dwEffect & DROPEFFECT_SCROLL) && m_pDropTargetHelper)
	{
		if (m_bEntered)
			m_pDropTargetHelper->DragOver(&point, dwEffect);
		else
			m_pDropTargetHelper->DragEnter(pWnd->GetSafeHwnd(), m_lpDataObject, &point, dwEffect);
		dwEffect |= DROPEFFECT_SCROLL;
	}
	return dwEffect;
}

BEGIN_INTERFACE_MAP(COleDropTargetEx, COleDropTarget)
	INTERFACE_PART(COleDropTargetEx, IID_IDropTarget, DropTargetEx)
END_INTERFACE_MAP()

STDMETHODIMP_(ULONG) COleDropTargetEx::XDropTargetEx::AddRef()
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) COleDropTargetEx::XDropTargetEx::Release()
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	return pThis->ExternalRelease();
}

STDMETHODIMP COleDropTargetEx::XDropTargetEx::QueryInterface(REFIID iid, LPVOID* ppvObj)
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	return (HRESULT)pThis->ExternalQueryInterface(&iid, ppvObj);
}

STDMETHODIMP COleDropTargetEx::XDropTargetEx::DragEnter(THIS_ LPDATAOBJECT lpDataObject,
	DWORD dwKeyState, POINTL pt, LPDWORD pdwEffect)
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	return pThis->m_xDropTarget.DragEnter(lpDataObject, dwKeyState, pt, pdwEffect);
}

STDMETHODIMP COleDropTargetEx::XDropTargetEx::DragOver(THIS_ DWORD dwKeyState,
	POINTL pt, LPDWORD pdwEffect)
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	return pThis->m_xDropTarget.DragOver(dwKeyState, pt, pdwEffect);
}

STDMETHODIMP COleDropTargetEx::XDropTargetEx::DragLeave(THIS)
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	return pThis->m_xDropTarget.DragLeave();
}

STDMETHODIMP COleDropTargetEx::XDropTargetEx::Drop(THIS_ LPDATAOBJECT lpDataObject,
	DWORD dwKeyState, POINTL pt, LPDWORD pdwEffect)
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	return pThis->m_xDropTarget.Drop(lpDataObject, dwKeyState, pt, pdwEffect);
}
 