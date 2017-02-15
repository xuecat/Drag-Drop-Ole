#pragma once
/*!
 * @file OleDataSourceEx.h
 *
 * @author XUECAT
 * @date 五月 2016
 *
 * @note 调用时必须指针开辟到堆（即必须用new）
 */
#include "OleDropSourceEx.h"
#include <vector>

class COleDataSourceEx : public COleDataSource
{
public:
	COleDataSourceEx();
	virtual ~COleDataSourceEx();
	
public:
	void ClearGlobalData();
	inline bool WasDragStarted() const { return m_nDragResult & DRAG_RES_STARTED; }
	inline int GetDragResult() const { return m_nDragResult; }

public:

	/** 
	 *  @return
	 *  @note		拖拽提示字符项目
	 */
	bool AllowDropDescriptionText();
	bool SetDropDescriptionText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1, LPCWSTR lpszInsert = NULL);
	
	//void IsDragOnSpecially(bool bdrag);

	/** 
	 *  @param [in] hDc     控件DC
	 *  @param [in] hBitmap 要拖拽的位图
	 *  @param [in] pOffset 拖拽鼠标的偏移量（因为拖拽的大小默认用图片的大小，偏移量默认用中间偏移）
	 *  @param [in] clr     拖拽图片的背景颜色
	 *  @return				是否成功
	 *  @note				会把DDB位图拷贝转换成DIB位图，在放入clipboard进行拖拽
	 */
	bool SetDragImage(HDC hDc, HBITMAP hBitmap, const CPoint* pOffset = NULL, COLORREF clr = RGB(255, 255, 255));
	inline bool SetDragImage(HDC hDc, CBitmap& Bitmap, const CPoint* pOffset, COLORREF clr)
	{
		return SetDragImage(hDc, static_cast<HBITMAP>(Bitmap.Detach()), pOffset, clr);
	}

	bool SetDragImageWindow(HWND hWnd, POINT* pPoint);

	/** 
	 *  @return
	 *  @note		设置拖拽数据
	 */
	bool SetCacheGloablDropFileListData(CStringList& list, UINT sz);
	bool SetCacheGlobalClipboardData(LPCTSTR clipboardformat, DWORD data);

	/**
	*  @param [in] dwEffect 拖拽类型 DROPEFFECT_COPY  DROPEFFECT_MOVE
	*  @param [in] lpRectStartDrag 拖拽的起始矩形。即当鼠标离开这个矩形后拖拽事件才正式调用（有延迟时间设定，具体看msdn）
	*  @return
	*  @note
	*/
	DROPEFFECT DoDragDropEx(DROPEFFECT dwEffect, LPCRECT lpRectStartDrag = NULL);
protected:

	/** 
	 *  @note 记住用::DeleteObject释放
	 */
	HBITMAP GeneraDIBBitmap(HDC hsrcdc, HBITMAP hsrcbitmap);

	bool SetCacheGlobalDropFlagData();

protected:
	bool				m_bSetDescriptionText;

	int					m_nDragResult;

	IDragSourceHelper*	m_pDragSourceHelper;
	IDragSourceHelper2*	m_pDragSourceHelper2;

	HBITMAP				m_hDragBitmap;

	CDropDescription	m_DropDescription;
	
	 
	 /**< @brief 为了自控制内存，所有的拖拽数据全部用的是自释放标志,  */
	/**< @brief 临时存放数据, 拖拽结束后全部释放 */
	std::vector<HGLOBAL> m_VCacheData;
public:
	
	BEGIN_INTERFACE_PART(DataObjectEx, IDataObject)
		INIT_INTERFACE_PART(COleDataSourceEx, DataObjectEx)
		STDMETHOD(GetData)(LPFORMATETC, LPSTGMEDIUM);
		STDMETHOD(GetDataHere)(LPFORMATETC, LPSTGMEDIUM);
		STDMETHOD(QueryGetData)(LPFORMATETC);
		STDMETHOD(GetCanonicalFormatEtc)(LPFORMATETC, LPFORMATETC);
		STDMETHOD(SetData)(LPFORMATETC, LPSTGMEDIUM, BOOL);
		STDMETHOD(EnumFormatEtc)(DWORD, LPENUMFORMATETC*);
		STDMETHOD(DAdvise)(LPFORMATETC, DWORD, LPADVISESINK, LPDWORD);
		STDMETHOD(DUnadvise)(DWORD);
		STDMETHOD(EnumDAdvise)(LPENUMSTATDATA*);
	END_INTERFACE_PART(DataObjectEx)

	DECLARE_INTERFACE_MAP()
};
