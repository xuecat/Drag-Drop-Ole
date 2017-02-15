#pragma once

/*!
 * @file OleDropTargetEx.h
 *
 * @author XUECAT
 * @date 五月 2016
 *
 * @note TargetFun外部继承使用
 */

#include "DragDropHelper.h"

#define SCROLL_LEFT  1
#define SCROLL_RIGHT 2
#define SCROLL_UP	 3
#define SCROLL_DOWN  4

class TargetFun;
class COleDropTargetEx : public COleDropTarget
{

public:
	COleDropTargetEx();

	virtual DROPEFFECT OnDragEnter(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
	virtual DROPEFFECT OnDragOver(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
	virtual DROPEFFECT OnDropEx(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT dropList, CPoint point);
	virtual BOOL OnDrop(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropEffect, CPoint point);
	virtual void OnDragLeave(CWnd* pWnd);    
	virtual DROPEFFECT OnDragScroll(CWnd* pWnd, DWORD dwKeyState, CPoint point);

public:
	virtual ~COleDropTargetEx();

	bool SetDropDescriptionText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1, LPCWSTR lpszInsert = NULL);
	bool SetDropInsertText(LPCWSTR lpszDropInsert);
	bool SetDropDescription(DROPIMAGETYPE nImageType, LPCWSTR lpszText, bool bCreate);
	bool SetDropDescription(DROPEFFECT dwEffect);
	inline bool ClearDropDescription()
	{
		return SetDropDescription(DROPIMAGE_INVALID, NULL, false);
	}
	
	/** 
	 *  @param [in] pWnd 注册窗口
	 *  @param [in] pFun 回调函数类（方便使用TargetFun的回调函数）
	 *  @return
	 *  @note
	 */
	BOOL Register(CWnd* pWnd, TargetFun* pFun = NULL);

	inline void IsAcceptDropFromExplorer(bool Is) { m_bAcceptExplorer = Is; }
	inline void SetAcceptEffect(DROPEFFECT effect) { m_nAllowEffect = effect; }

protected:
	void		ClearStateFlags();
	void		OnPostDrop(COleDataObject* pDataObject, DROPEFFECT dropEffect, CPoint point);
	
	/** 
	 *  @return  是否加了DROPEFFECT_SCROLL
	 *  @note    预留，用来判断是否滑动
	 */
	INT  IsDoScroll(CWnd* pWnd, DWORD dwKeyState, CPoint point);
protected:

	TargetFun*			m_pTargetFun;

	IDropTargetHelper*	m_pDropTargetHelper;

	CDropDescription	m_DropDescription;
private:
	bool				m_bAcceptExplorer;

	bool				m_bCanShowDescription;
	bool				m_bUseDropDescription;
	bool				m_bDescriptionUpdated;

	bool				m_bHasFlag;

	bool				m_bEntered;
	bool				m_bHasDragImage;
	bool				m_bTextAllowed;

	DROPEFFECT			m_nAllowEffect;
public:
	BEGIN_INTERFACE_PART(DropTargetEx, IDropTarget)
		INIT_INTERFACE_PART(COleDropTargetEx, DropTargetEx)
		STDMETHOD(DragEnter)(LPDATAOBJECT, DWORD, POINTL, LPDWORD);
		STDMETHOD(DragOver)(DWORD, POINTL, LPDWORD);
		STDMETHOD(DragLeave)();
		STDMETHOD(Drop)(LPDATAOBJECT, DWORD, POINTL, LPDWORD);
	END_INTERFACE_PART(DropTargetEx)

	DECLARE_INTERFACE_MAP()

};

class __declspec(novtable) TargetFun
{
	/*!
	 * @class TargetFun
	 *
	 * @brief 

	 * @node 如果设置了回调函数
	 * @node OnDragEnterFunc、OnDragOverFunc 若不使用，请返回DROPEFFECT_NONE.
	 * @node OnDragScrollFunc 若不使用请返回DROPEFFECT_NONE、使用就带DROPEFFECT_SCROLL
	 * @node OnDropFunc 的true和false与 最后拖拽返回值有关（成功true，失败false）

	 * @author XUECAT
	 * @date 五月 2016
	 */
public:
	virtual DROPEFFECT	OnDragEnterFunc(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point) = 0;
	virtual DROPEFFECT	OnDragOverFunc(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point) = 0;
	virtual BOOL		OnDropFunc(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropeffect, CPoint point) = 0;
	virtual VOID		OnDragLeaveFunc(CWnd* pWnd) = 0;
	virtual DROPEFFECT  OnDragScrollFunc(CWnd* pWnd, DWORD dwKeyState, CPoint point, INT ndirect) = 0;

	COleDropTargetEx	m_DropTarget;
};
