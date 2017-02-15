
#pragma once
/*!
* @file OleDropSourceEx.h
*
* @author XUECAT
* @date 六月 2016
*
* @node 若想动态修改拖拽过程中的鼠标样式(discription的提示图片无法修改已尝试,可以通过改变cursor来达到改变discription的错觉)
* @node 可参照https://wpf.2000things.com/tag/drag-and-drop/ 在GiveFeedback动态修改
*
*/


#include "DragDropHelper.h"


#define DRAG_RES_NOT_STARTED		0		// 拖拽未开始
#define DRAG_RES_STARTED			1		// 拖拽已开始 (timeout expired or mouse left the start rect)
#define DRAG_RES_CANCELLED			2		// 取消 (may be started or not)
#define DRAG_RES_RELEASED			4		// 左鼠标释放
#define DRAG_RES_DROPPED			8		//  (return effect not DROPEEFECT_NONE)


class COleDropSourceEx : public COleDropSource
{
public:
	COleDropSourceEx();

protected:
	bool			m_bSetCursor;
	int				m_nResult;
	LPDATAOBJECT	m_pIDataObj;
	const CDropDescription *m_pDropDescription;

	bool			SetDragImageCursor(DROPEFFECT dwEffect);

	virtual BOOL	OnBeginDrag(CWnd* pWnd);
	virtual SCODE	QueryContinueDrag(BOOL bEscapePressed, DWORD dwKeyState);

	virtual SCODE	GiveFeedback(DROPEFFECT dropEffect);

public:
	friend class COleDataSourceEx;
};

