
#pragma once
/*!
* @file OleDropSourceEx.h
*
* @author XUECAT
* @date ���� 2016
*
* @node ���붯̬�޸���ק�����е������ʽ(discription����ʾͼƬ�޷��޸��ѳ���,����ͨ���ı�cursor���ﵽ�ı�discription�Ĵ��)
* @node �ɲ���https://wpf.2000things.com/tag/drag-and-drop/ ��GiveFeedback��̬�޸�
*
*/


#include "DragDropHelper.h"


#define DRAG_RES_NOT_STARTED		0		// ��קδ��ʼ
#define DRAG_RES_STARTED			1		// ��ק�ѿ�ʼ (timeout expired or mouse left the start rect)
#define DRAG_RES_CANCELLED			2		// ȡ�� (may be started or not)
#define DRAG_RES_RELEASED			4		// ������ͷ�
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

