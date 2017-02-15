
// DragDrop2Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "DragDrop2.h"
#include "DragDrop2Dlg.h"
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


// CDragDrop2Dlg dialog



CDragDrop2Dlg::CDragDrop2Dlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CDragDrop2Dlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CDragDrop2Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TAB1, m_tabctrl);
}

BEGIN_MESSAGE_MAP(CDragDrop2Dlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
END_MESSAGE_MAP()

// CDragDrop2Dlg message handlers

BOOL CDragDrop2Dlg::OnInitDialog()
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

	// TODO: Add extra initialization here
	m_DialogInfo.Create(CDialogInfo::IDD, &m_tabctrl);

	m_tabctrl.AddTab(&m_DialogInfo, 
		L"More Information", 0);
	m_tabctrl.AddTab(&m_DialogInfo,
		L"Test Information", 1);

	//** customizing the tab control --------
	//	m_tabctrl.SetDisabledColor(RGB(255, 0, 0));
	m_tabctrl.SetSelectedColor(RGB(0, 0, 255));
	m_tabctrl.SetMouseOverColor(RGB(255, 255, 255));
	m_tabctrl.EnableTab(1, FALSE);
	m_tabctrl.SetItemSize(CSize(50, 50));

	m_DropTarget.Register(this);
	m_DropTarget.SetAcceptEffect(DROPEFFECT_COPY);
	m_DropTarget.SetDropDescriptionText(DROPIMAGE_COPY, _T("复制到"), _T("复制到 %1"), _T("WQ窗口"));

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CDragDrop2Dlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CDragDrop2Dlg::OnPaint()
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
HCURSOR CDragDrop2Dlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

BOOL CDragDrop2Dlg::OnDropFunc(CWnd*, COleDataObject* p, DROPEFFECT, CPoint)
{
	TRACE(_T("copy move \n"));

	CStringList list;
	if (CDragDropHelper::GetGloablDropFileListData(p, list))
	{
		TCHAR buffer[10] = { 0 };
		int count = list.GetCount();
		_itot_s(count, buffer, 10, 10);

		TRACE(_T("list count:"));
		TRACE(buffer);
		TRACE(_T("\n"));

		TRACE(_T("list head:"));
		TRACE(list.GetHead().GetString());
	}
	return TRUE;
}