// DialogInfo.cpp : implementation file
//

#include "stdafx.h"
#include "DragDrop2.h"
#include "DialogInfo.h"
#include "afxdialogex.h"


// CDialogInfo dialog

IMPLEMENT_DYNAMIC(CDialogInfo, CDialogEx)

CDialogInfo::CDialogInfo(CWnd* pParent /*=NULL*/)
	: CDialogEx(CDialogInfo::IDD, pParent)
{

}

CDialogInfo::~CDialogInfo()
{
}

void CDialogInfo::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CDialogInfo, CDialogEx)
END_MESSAGE_MAP()


// CDialogInfo message handlers
