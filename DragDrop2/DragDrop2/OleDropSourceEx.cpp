
#include "StdAfx.h"
#include "OleDropSourceEx.h"

#include <basetsd.h>
#include <WinUser.h>

#ifndef DDWM_UPDATEWINDOW
#define DDWM_UPDATEWINDOW	(WM_USER + 3)
#endif

#define DDWM_SETCURSOR		(WM_USER + 2)

COleDropSourceEx::COleDropSourceEx()
	: m_bSetCursor(true)
	, m_nResult(DRAG_RES_NOT_STARTED)
	, m_pIDataObj(NULL)
	, m_pDropDescription(NULL)
{
}

bool COleDropSourceEx::SetDragImageCursor(DROPEFFECT dwEffect)
{
	HWND hWnd = (HWND)ULongToHandle(CDragDropHelper::GetGlobalDataDWord(m_pIDataObj, _T("DragWindow")));
	if (hWnd)
	{
		WPARAM wParam = 0;
		switch (dwEffect & ~DROPEFFECT_SCROLL)
		{
		case DROPEFFECT_NONE : wParam = 1; break;
		case DROPEFFECT_COPY : wParam = 3; break;
		case DROPEFFECT_MOVE : wParam = 2; break;
		case DROPEFFECT_LINK : wParam = 4; break;
		}
		::SendMessage(hWnd, DDWM_SETCURSOR, wParam, 0);
	}
	return NULL != hWnd;
}


BOOL COleDropSourceEx::OnBeginDrag(CWnd* pWnd)
{
	
	BOOL bRet = COleDropSource::OnBeginDrag(pWnd);
	if (m_bDragStarted)
		m_nResult = DRAG_RES_STARTED;
	else if (GetKeyState((m_dwButtonDrop & MK_LBUTTON) ? VK_LBUTTON : VK_RBUTTON) >= 0)
		m_nResult = DRAG_RES_RELEASED;
	else
		m_nResult = DRAG_RES_CANCELLED;
	return bRet;
}


SCODE COleDropSourceEx::QueryContinueDrag(BOOL bEscapePressed, DWORD dwKeyState)
{
	SCODE scRet = COleDropSource::QueryContinueDrag(bEscapePressed, dwKeyState);
	if (0 == (dwKeyState & m_dwButtonDrop))
		m_nResult |= DRAG_RES_RELEASED;
	if (DRAGDROP_S_CANCEL == scRet)	
		m_nResult |= DRAG_RES_CANCELLED;
	return scRet;
}


SCODE COleDropSourceEx::GiveFeedback(DROPEFFECT dropEffect)
{
	SCODE sc = COleDropSource::GiveFeedback(dropEffect);
	
    if (m_bDragStarted && m_pIDataObj)
    {
		bool bOldStyle = (0 == CDragDropHelper::GetGlobalDataDWord(m_pIDataObj, _T("IsShowingLayered")));
	
		
		if ((bOldStyle && !m_bSetCursor) || (m_pDropDescription && !bOldStyle))
		{
			FORMATETC FormatEtc;
			STGMEDIUM StgMedium;
			if (CDragDropHelper::GetGlobalData(m_pIDataObj, CFSTR_DROPDESCRIPTION, FormatEtc, StgMedium))
			{
				bool bChangeDescription = false;
				DROPDESCRIPTION *pDropDescription = static_cast<DROPDESCRIPTION*>(::GlobalLock(StgMedium.hGlobal));
				if (bOldStyle)
				{
					bChangeDescription = CDragDropHelper::ClearDescription(pDropDescription);
				}
				
				else if (pDropDescription->type <= DROPIMAGE_LINK)
				{
					DROPIMAGETYPE nImageType = CDragDropHelper::DropEffectToDropImage(dropEffect);
					
					if (DROPIMAGE_INVALID != nImageType &&
						pDropDescription->type != nImageType)
					{
						if (m_pDropDescription->HasText(nImageType))
						{
							bChangeDescription = true;
							pDropDescription->type = nImageType;
							m_pDropDescription->SetDescription(pDropDescription, nImageType);
						}
					
						else
							bChangeDescription = CDragDropHelper::ClearDescription(pDropDescription);
					}
				}	
				::GlobalUnlock(StgMedium.hGlobal);
				if (bChangeDescription)			
				{								
					if (FAILED(m_pIDataObj->SetData(&FormatEtc, &StgMedium, TRUE)))
						bChangeDescription = false;
				}
				if (!bChangeDescription)		
					::ReleaseStgMedium(&StgMedium);
			}
		}
		if (!bOldStyle)								
		{
			if (m_bSetCursor)
			{
				
				HCURSOR hCursor = (HCURSOR)LoadImage(
					NULL,							
					MAKEINTRESOURCE(32512),	
					IMAGE_CURSOR,					
					0, 0,						
					LR_DEFAULTSIZE | LR_SHARED);
				::SetCursor(hCursor);
			}
		
			SetDragImageCursor(dropEffect);
			sc = S_OK;						
        }
		m_bSetCursor = bOldStyle;
    }
	return sc;
}
