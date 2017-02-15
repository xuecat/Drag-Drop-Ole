
// DragDropDlg.cpp : implementation file
//

#include "stdafx.h"
#include "DragDrop.h"
#include "DragDropDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CDragDropDlg dialog



CDragDropDlg::CDragDropDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CDragDropDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CDragDropDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_List);
}

BEGIN_MESSAGE_MAP(CDragDropDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_LBUTTONDOWN()
	ON_WM_QUERYDRAGICON()
	ON_NOTIFY(LVN_BEGINDRAG, IDC_LIST1, OnBegindragFilelist)
END_MESSAGE_MAP()


// CDragDropDlg message handlers

BOOL CDragDropDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	m_List.InsertColumn(0, L"ID", LVCFMT_LEFT, 80);//插入列
	m_List.InsertColumn(1, L"NAME", LVCFMT_LEFT, 120);
	m_List.InsertColumn(2, L"HEHE", LVCFMT_LEFT, 120);
	m_List.InsertColumn(3, L"DEDEE", LVCFMT_LEFT, 120);

	
	for (int i = 0; i < 40; i++)
	{
		TCHAR buffer[10] = { 0 };
		_itot_s(i, buffer, 10, 10);
		CString str("data ");

		int nRow = m_List.InsertItem(i, str + buffer);//插入行
		m_List.SetItemText(nRow, 1, str + _T("wangqiu") + buffer);//设置数据
		m_List.SetItemText(nRow, 2, str +  _T("wangqiuHE") + buffer);//设置数据
		m_List.SetItemText(nRow, 3, str + _T("wangqiuHEDE") + buffer);//设置数据

	}

	// TODO: Add extra initialization here
	m_DropTarget.Register(&m_List, this);
	m_DropTarget.SetAcceptEffect(DROPEFFECT_MOVE);

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CDragDropDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CDragDropDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CDragDropDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

#include "resource.h"

VOID CDragDropDlg::OnLButtonDown(UINT nFlags, CPoint point)
{
	
}

BOOL CDragDropDlg::OnDropFunc(CWnd*, COleDataObject* P, DROPEFFECT, CPoint)
{
	TRACE(_T("local move \n"));

	CStringList list;
	if (CDragDropHelper::GetGloablDropFileListData(P, list))
	{
		TCHAR buffer[10] = { 0 };
		int count = list.GetCount();
		_itot_s(count, buffer, 10, 10);

		TRACE(_T("list count:"));
		TRACE(buffer);
		TRACE(_T("\n"));

		TRACE(_T("list head:"));
		CString str = list.GetHead();
		TRACE(str);
	}
	return TRUE;
}

VOID CDragDropDlg::OnBegindragFilelist(NMHDR* pNMHDR, LRESULT* pResult)
{
	*pResult = 0;

	UINT			uBuffSize = 0;
	CStringList    lsDraggedFiles;
	POSITION       pos;
	
	pos = m_List.GetFirstSelectedItemPosition();

	while (pos != NULL)
	{
		CString sFile = m_List.GetItemText(m_List.GetNextSelectedItem(pos), 0);
		lsDraggedFiles.AddTail(sFile);
		uBuffSize += lstrlen(sFile) + 1;
	}
	
	
	CDC* pDC = GetDC();
	CBitmap bit;
	bit.LoadBitmap(IDB_BITMAP1);


	//必须使用指针
	COleDataSourceEx* pDataSource = new COleDataSourceEx;
	pDataSource->AllowDropDescriptionText();
	pDataSource->SetDropDescriptionText(DROPIMAGE_MOVE, _T("移动到"), _T("移动到 %1"), _T("hehe窗口"));

	pDataSource->SetCacheGloablDropFileListData(lsDraggedFiles, uBuffSize);
	pDataSource->SetDragImage(pDC->GetSafeHdc(), (HBITMAP)bit.m_hObject, &CPoint(0, 0));
	int result = pDataSource->DoDragDropEx(DROPEFFECT_MOVE | DROPEFFECT_COPY);
	pDataSource->InternalRelease();

	ReleaseDC(pDC);
}