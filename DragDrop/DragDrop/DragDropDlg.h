
// DragDropDlg.h : header file
//

#pragma once
#include "OleDropTargetEx.h"
#include "OleDataSourceEx.h"
#include "afxcmn.h"

// CDragDropDlg dialog
class CDragDropDlg : public CDialogEx , public TargetFun
{
// Construction
public:
	CDragDropDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_DRAGDROP_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


	virtual DROPEFFECT	OnDragEnterFunc(CWnd*, COleDataObject*, DWORD, CPoint) { return DROPEFFECT_NONE; }
	virtual DROPEFFECT	OnDragOverFunc(CWnd*, COleDataObject*, DWORD, CPoint) { return DROPEFFECT_NONE; }
	virtual BOOL		OnDropFunc(CWnd*, COleDataObject* P, DROPEFFECT, CPoint);
	virtual VOID		OnDragLeaveFunc(CWnd*) {}
	virtual DROPEFFECT  OnDragScrollFunc(CWnd*, DWORD, CPoint, INT ndirect) 
	{ 
		switch (ndirect)
		{
		case SCROLL_UP:
		{
			m_List.PostMessage(WM_VSCROLL, SB_LINEUP, NULL);
		} break;
		case  SCROLL_DOWN:
			m_List.PostMessage(WM_VSCROLL, SB_LINEDOWN, NULL);
			break;
		case SCROLL_LEFT:
			m_List.PostMessage(WM_HSCROLL, SB_LINELEFT, NULL);
			break;
		case SCROLL_RIGHT:
			m_List.PostMessage(WM_HSCROLL, SB_LINERIGHT, NULL);
			break;
		default:
			break;
		}
		return DROPEFFECT_SCROLL;
	}

// Implementation
protected:
	HICON m_hIcon;

	
	

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg HCURSOR OnQueryDragIcon();

	afx_msg void OnBegindragFilelist(NMHDR* pNMHDR, LRESULT* pResult);
	DECLARE_MESSAGE_MAP()
	
public:
	CListCtrl m_List;
};
