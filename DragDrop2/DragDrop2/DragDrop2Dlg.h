
// DragDrop2Dlg.h : header file
//

#pragma once

#include "OleDropTargetEx.h"
#include "afxcmn.h"
#include "XTabCtrl.h"
#include "DialogInfo.h"

// CDragDrop2Dlg dialog
class CDragDrop2Dlg : public CDialogEx, public TargetFun
{
// Construction
public:
	CDragDrop2Dlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_DRAGDROP2_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	virtual DROPEFFECT	OnDragEnterFunc(CWnd*, COleDataObject*, DWORD, CPoint) { TRACE(_T("ENTER")); return DROPEFFECT_NONE; }
	virtual DROPEFFECT	OnDragOverFunc(CWnd*, COleDataObject*, DWORD, CPoint) { return DROPEFFECT_NONE; }
	virtual BOOL		OnDropFunc(CWnd*, COleDataObject*, DROPEFFECT, CPoint);
	virtual VOID		OnDragLeaveFunc(CWnd*) { TRACE(_T("LEAVE")); }
	virtual DROPEFFECT  OnDragScrollFunc(CWnd* pWnd, DWORD dwKeyState, CPoint point, INT ndirect) { return DROPEFFECT_NONE; }

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	CDialogInfo m_DialogInfo;
	CXTabCtrl m_tabctrl;
};
