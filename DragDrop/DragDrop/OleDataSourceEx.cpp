#include "StdAfx.h"
#include "OleDataSourceEx.h"

#define Delete_Object(p) if (p) {\
	::DeleteObject(p);\
	p = NULL;\
}

//拖拽全局标签
UINT g_uCustomClipbrdFormats = RegisterClipboardFormat(_T("MsvList_3BCFE9D1_6D61_4cb6_9D0B_3BB3F643CA82"));

COleDataSourceEx::COleDataSourceEx()
	: m_nDragResult(0)
	, m_bSetDescriptionText(false)
	, m_pDragSourceHelper(NULL)
	, m_pDragSourceHelper2(NULL)
	, m_hDragBitmap(NULL)
{
	
#if defined(IID_PPV_ARGS) 
	::CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pDragSourceHelper));
#else
	::CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER,
		IID_IDragSourceHelper, reinterpret_cast<LPVOID*>(&m_pDragSourceHelper));
#endif

	if (m_pDragSourceHelper) //获取IID_IDragSourceHelper2接口，只有这个接口才支持拖拽提示
	{
#if defined(IID_PPV_ARGS) && defined(__IDragSourceHelper2_INTERFACE_DEFINED__)
		m_pDragSourceHelper->QueryInterface(IID_PPV_ARGS(&m_pDragSourceHelper2));
#else
		m_pDragSourceHelper->QueryInterface(IID_IDragSourceHelper2, reinterpret_cast<LPVOID*>(&m_pDragSourceHelper2));
#endif
		if (NULL != m_pDragSourceHelper2) 
		{
			m_pDragSourceHelper->Release();
			m_pDragSourceHelper = static_cast<IDragSourceHelper*>(m_pDragSourceHelper2);
		}
	}
}

COleDataSourceEx::~COleDataSourceEx()
{
	if (NULL != m_pDragSourceHelper)
		m_pDragSourceHelper->Release();
	Delete_Object(m_hDragBitmap);
}

VOID COleDataSourceEx::ClearGlobalData()
{
	int sz = m_VCacheData.size();
	for (int i = 0; i < sz; i++) {
		if (m_VCacheData[i] != NULL) {
			GlobalFree(m_VCacheData[i]);
			m_VCacheData[i] = NULL;
		}
	}
	m_VCacheData.clear();
}

DROPEFFECT COleDataSourceEx::DoDragDropEx(DROPEFFECT dwEffect, LPCRECT lpRectStartDrag /*= NULL*/)
{
	bool bUseDescription = (m_pDragSourceHelper2 != NULL) && ::IsAppThemed();
	
	if (bUseDescription && m_bSetDescriptionText)
	{												
		CLIPFORMAT cfDS = CDragDropHelper::RegisterFormat(CFSTR_DROPDESCRIPTION); 
		if (cfDS)
		{
			HGLOBAL hGlobal = ::GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, sizeof(DROPDESCRIPTION));
			if (hGlobal)
			{
				DROPDESCRIPTION *pDropDescription = static_cast<DROPDESCRIPTION*>(::GlobalLock(hGlobal));
				pDropDescription->type = DROPIMAGE_INVALID;
				::GlobalUnlock(hGlobal);
				CacheGlobalData(cfDS, hGlobal);
			}
		}
	}

	SetCacheGlobalDropFlagData();

	COleDropSourceEx dropSource;
	if (bUseDescription)
	{
		dropSource.m_pIDataObj = static_cast<LPDATAOBJECT>(GetInterface(&IID_IDataObject));

		if (m_bSetDescriptionText)
			dropSource.m_pDropDescription = &m_DropDescription;
	}
	
	dwEffect = DoDragDrop(dwEffect, lpRectStartDrag, static_cast<COleDropSource*>(&dropSource));
	m_nDragResult = dropSource.m_nResult;

	if (dwEffect & ~DROPEFFECT_SCROLL)
		m_nDragResult |= DRAG_RES_DROPPED;

	Delete_Object(m_hDragBitmap);
	ClearGlobalData();

	return dwEffect;
}


bool COleDataSourceEx::AllowDropDescriptionText()
{
	return m_pDragSourceHelper2 ? SUCCEEDED(m_pDragSourceHelper2->SetFlags(DSH_ALLOWDROPDESCRIPTIONTEXT)) : false;
}

bool COleDataSourceEx::SetDropDescriptionText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1, LPCWSTR lpszInsert /*= NULL*/)
{
	bool bRet = false;
	if (m_pDragSourceHelper2)
	{
		bRet = m_DropDescription.SetText(nType, lpszText, lpszText1);
		if (bRet && lpszInsert)
			m_DropDescription.SetInsert(lpszInsert);
		m_bSetDescriptionText |= bRet;
	}
	return bRet;
}

bool COleDataSourceEx::SetDragImageWindow(HWND hWnd, POINT* pPoint)
{
	HRESULT hRes = E_NOINTERFACE;
	if (m_pDragSourceHelper)
	{
		POINT pt = { 0, 0 };
		if (NULL == pPoint)
			pPoint = &pt;
		hRes = m_pDragSourceHelper->InitializeFromWindow(hWnd, pPoint, 
			static_cast<LPDATAOBJECT>(GetInterface(&IID_IDataObject)));
#ifdef _DEBUG
		if (FAILED(hRes))
			TRACE1("COleDataSourceEx::SetDragImageWindow: InitializeFromWindow failed with code %#X\n", hRes);
#endif
	}
	return SUCCEEDED(hRes);
}


HBITMAP COleDataSourceEx::GeneraDIBBitmap(HDC hsrcdc, HBITMAP hsrcbitmap)
{
	BITMAP bmap = { 0 };
	GetObject(hsrcbitmap, sizeof(BITMAP), &bmap);
	int cx = bmap.bmWidth;
	int cy = bmap.bmHeight;

	HDC hPaintDC = ::CreateCompatibleDC(hsrcdc);
	ASSERT(hPaintDC);
	ASSERT(hsrcbitmap);
	HBITMAP hOldPaintBitmap = (HBITMAP) ::SelectObject(hPaintDC, hsrcbitmap);

	BITMAPINFO bmi = { 0 };
	bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bmi.bmiHeader.biWidth = cx;
	bmi.bmiHeader.biHeight = cy;
	bmi.bmiHeader.biPlanes = 1;
	bmi.bmiHeader.biBitCount = 32;
	bmi.bmiHeader.biCompression = BI_RGB;
	bmi.bmiHeader.biSizeImage = cx * cy * sizeof(DWORD);
	LPDWORD pDest = NULL;
	HDC hCloneDC = ::CreateCompatibleDC(hsrcdc);
	HBITMAP hBitmap = ::CreateDIBSection(hsrcdc, &bmi, DIB_RGB_COLORS, (LPVOID*)&pDest, NULL, 0);
	ASSERT(hCloneDC);
	ASSERT(hBitmap);
	if (hBitmap != NULL)
	{
		HBITMAP hOldBitmap = (HBITMAP) ::SelectObject(hCloneDC, hBitmap);
		::BitBlt(hCloneDC, 0, 0, cx, cy, hPaintDC, 0, 0, SRCCOPY);
		::SelectObject(hCloneDC, hOldBitmap);
		::DeleteDC(hCloneDC);
		::GdiFlush();
	}

	// Cleanup
	::SelectObject(hPaintDC, hOldPaintBitmap);
	::DeleteDC(hPaintDC);

	Delete_Object(m_hDragBitmap);
	m_hDragBitmap = hBitmap;

	return hBitmap;
}


bool COleDataSourceEx::SetDragImage(HDC hDc, HBITMAP hBitmap, const CPoint* pOffset, COLORREF clr)
{
	ASSERT(hBitmap);

	HRESULT hRes = E_NOINTERFACE;
	hBitmap = GeneraDIBBitmap(hDc, hBitmap);

	if (hBitmap && m_pDragSourceHelper)
	{
		BITMAP bm = {0};
		SHDRAGIMAGE di = {0};

		VERIFY(::GetObject(hBitmap, sizeof(bm), &bm));
		di.sizeDragImage.cx = bm.bmWidth;
		di.sizeDragImage.cy = bm.bmHeight;

		if (pOffset)
		{
			di.ptOffset = *pOffset;
			if (di.ptOffset.x < 0)
				di.ptOffset.x = di.sizeDragImage.cx / 2;
			else if (di.ptOffset.x > di.sizeDragImage.cx)
				di.ptOffset.x = di.sizeDragImage.cx;

			if (di.ptOffset.y < 0)
				di.ptOffset.y = di.sizeDragImage.cy / 2;
			else if (di.ptOffset.y > di.sizeDragImage.cy)
				di.ptOffset.y = di.sizeDragImage.cy;
		}
		else
		{
			di.ptOffset.x = di.sizeDragImage.cx / 2;
			di.ptOffset.y = di.sizeDragImage.cy / 2;
		}
		di.hbmpDragImage = hBitmap;
		di.crColorKey = (CLR_INVALID == clr) ? ::GetSysColor(COLOR_WINDOW) : clr;

		hRes = m_pDragSourceHelper->InitializeFromBitmap(&di,
			static_cast<LPDATAOBJECT>(GetInterface(&IID_IDataObject)));

#ifdef _DEBUG
		if (FAILED(hRes))
			TRACE1("COleDataSourceEx::SetDragImage: InitializeFromBitmap failed with code %#X\n", hRes);
#endif
	}

	if (FAILED(hRes) && hBitmap)
		::DeleteObject(hBitmap);
	return SUCCEEDED(hRes);
}


bool COleDataSourceEx::SetCacheGloablDropFileListData(CStringList& list, UINT sz)
{
	if (sz <= 0)
		return false;

	FORMATETC etc = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	TCHAR*         pszBuff = NULL;
	DROPFILES*     pDrop = NULL;
	HGLOBAL hdrop = GlobalAlloc(GHND | GMEM_SHARE, sizeof(DROPFILES) + sizeof(TCHAR) * (sz + 1));
	if (hdrop == NULL)
		return false;

	pDrop = (DROPFILES*)GlobalLock(hdrop);
	pDrop->pFiles = sizeof(DROPFILES);

#ifdef _UNICODE
	pDrop->fWide = TRUE;
#endif;

	pszBuff = (TCHAR*)(LPBYTE(pDrop) + sizeof(DROPFILES));
	POSITION npos = list.GetHeadPosition();
	while (npos != NULL)
	{
		lstrcpy(pszBuff, (LPCTSTR)list.GetNext(npos));
		pszBuff = 1 + _tcschr(pszBuff, '\0');
	}
	GlobalUnlock(hdrop);

	CacheGlobalData(CF_HDROP, hdrop, &etc);

	m_VCacheData.push_back(hdrop);
	return true;
}

bool COleDataSourceEx::SetCacheGlobalClipboardData(LPCTSTR clipboardformat, DWORD data)
{
	if (clipboardformat == NULL)
		return false;
	UINT uformat = ::RegisterClipboardFormatW(clipboardformat);
	FORMATETC etc = {uformat , NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	HGLOBAL hdrop = GlobalAlloc(GHND | GMEM_SHARE, sizeof(DWORD));
	
	DWORD* pDropEffect = (DWORD*)GlobalLock(hdrop);
	*pDropEffect = data;
	GlobalUnlock(hdrop);

	CacheGlobalData(uformat, hdrop, &etc);

	m_VCacheData.push_back(hdrop);
	return true;
}

bool COleDataSourceEx::SetCacheGlobalDropFlagData()
{
	if (g_uCustomClipbrdFormats) {
		HGLOBAL hdrop = GlobalAlloc(GHND | GMEM_SHARE, sizeof(bool));//初始化0
		FORMATETC etc = { g_uCustomClipbrdFormats, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };

		CacheGlobalData(g_uCustomClipbrdFormats, hdrop, &etc);
		
		m_VCacheData.push_back(hdrop);
		return true;
	}
	return false;
}

BEGIN_INTERFACE_MAP(COleDataSourceEx, COleDataSource)
	INTERFACE_PART(COleDataSourceEx, IID_IDataObject, DataObjectEx)
END_INTERFACE_MAP()

STDMETHODIMP_(ULONG) COleDataSourceEx::XDataObjectEx::AddRef()
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) COleDataSourceEx::XDataObjectEx::Release()
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->ExternalRelease();
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::QueryInterface(REFIID iid, LPVOID* ppvObj)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return (HRESULT)pThis->ExternalQueryInterface(&iid, ppvObj);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::GetData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.GetData(lpFormatEtc, lpStgMedium);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::GetDataHere(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.GetDataHere(lpFormatEtc, lpStgMedium);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::QueryGetData(LPFORMATETC lpFormatEtc)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.QueryGetData(lpFormatEtc);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::GetCanonicalFormatEtc(LPFORMATETC lpFormatEtcIn, LPFORMATETC lpFormatEtcOut)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.GetCanonicalFormatEtc(lpFormatEtcIn, lpFormatEtcOut);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::SetData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium, BOOL bRelease)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	HRESULT hr = pThis->m_xDataObject.SetData(lpFormatEtc, lpStgMedium, bRelease);
	if (DATA_E_FORMATETC == hr &&			
		(lpFormatEtc->tymed & (TYMED_HGLOBAL | TYMED_ISTREAM)) &&
		lpFormatEtc->cfFormat >= 0xC000)
	{
		pThis->CacheData(lpFormatEtc->cfFormat, lpStgMedium, lpFormatEtc);
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::EnumFormatEtc(DWORD dwDirection, LPENUMFORMATETC* ppenumFormatetc)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.EnumFormatEtc(dwDirection, ppenumFormatetc);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::DAdvise(LPFORMATETC lpFormatEtc, DWORD advf, LPADVISESINK pAdvSink, LPDWORD pdwConnection)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.DAdvise(lpFormatEtc, advf, pAdvSink, pdwConnection);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::DUnadvise(DWORD dwConnection)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.DUnadvise(dwConnection);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::EnumDAdvise(LPENUMSTATDATA* ppenumAdvise)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.EnumDAdvise(ppenumAdvise);
}

