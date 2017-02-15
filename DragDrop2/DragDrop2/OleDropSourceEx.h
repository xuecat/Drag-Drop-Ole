
#pragma once

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

