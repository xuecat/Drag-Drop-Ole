#pragma once
/*!
 * @file OleDataSourceEx.h
 *
 * @author XUECAT
 * @date ���� 2016
 *
 * @note ����ʱ����ָ�뿪�ٵ��ѣ���������new��
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
	 *  @note		��ק��ʾ�ַ���Ŀ
	 */
	bool AllowDropDescriptionText();
	bool SetDropDescriptionText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1, LPCWSTR lpszInsert = NULL);
	
	//void IsDragOnSpecially(bool bdrag);

	/** 
	 *  @param [in] hDc     �ؼ�DC
	 *  @param [in] hBitmap Ҫ��ק��λͼ
	 *  @param [in] pOffset ��ק����ƫ��������Ϊ��ק�Ĵ�СĬ����ͼƬ�Ĵ�С��ƫ����Ĭ�����м�ƫ�ƣ�
	 *  @param [in] clr     ��קͼƬ�ı�����ɫ
	 *  @return				�Ƿ�ɹ�
	 *  @note				���DDBλͼ����ת����DIBλͼ���ڷ���clipboard������ק
	 */
	bool SetDragImage(HDC hDc, HBITMAP hBitmap, const CPoint* pOffset = NULL, COLORREF clr = RGB(255, 255, 255));
	inline bool SetDragImage(HDC hDc, CBitmap& Bitmap, const CPoint* pOffset, COLORREF clr)
	{
		return SetDragImage(hDc, static_cast<HBITMAP>(Bitmap.Detach()), pOffset, clr);
	}

	bool SetDragImageWindow(HWND hWnd, POINT* pPoint);

	/** 
	 *  @return
	 *  @note		������ק����
	 */
	bool SetCacheGloablDropFileListData(CStringList& list, UINT sz);
	bool SetCacheGlobalClipboardData(LPCTSTR clipboardformat, DWORD data);

	/**
	*  @param [in] dwEffect ��ק���� DROPEFFECT_COPY  DROPEFFECT_MOVE
	*  @param [in] lpRectStartDrag ��ק����ʼ���Ρ���������뿪������κ���ק�¼�����ʽ���ã����ӳ�ʱ���趨�����忴msdn��
	*  @return
	*  @note
	*/
	DROPEFFECT DoDragDropEx(DROPEFFECT dwEffect, LPCRECT lpRectStartDrag = NULL);
protected:

	/** 
	 *  @note ��ס��::DeleteObject�ͷ�
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
	
	 
	 /**< @brief Ϊ���Կ����ڴ棬���е���ק����ȫ���õ������ͷű�־,  */
	/**< @brief ��ʱ�������, ��ק������ȫ���ͷ� */
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
