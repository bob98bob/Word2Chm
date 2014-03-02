// AdvanceDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Word2Chm.h"
#include "AdvanceDlg.h"
//#include "afxdialogex.h"
#include "OfficeWord.h"

// CAdvanceDlg dialog

IMPLEMENT_DYNAMIC(CAdvanceDlg, CDialog)

CAdvanceDlg::CAdvanceDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAdvanceDlg::IDD, pParent)

	, m_bHeader(FALSE)
	, m_bFooter(FALSE)
	, m_bPNG(TRUE)
	, m_bRelated(TRUE)
	, m_bListNumber(FALSE)
{

}

CAdvanceDlg::~CAdvanceDlg()
{
}

void CAdvanceDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_COPYRIGHT, m_copyrightEdit);
	DDX_Check(pDX, IDC_CHECK_HEADER, m_bHeader);
	DDX_Check(pDX, IDC_CHECK_FOOTER, m_bFooter);
	DDX_Check(pDX, IDC_CHECK_PNG, m_bPNG);
	DDX_Check(pDX, IDC_CHECK_RELATED, m_bRelated);
	DDX_Check(pDX, IDC_CHECK_LISTNUMBER, m_bListNumber);
}


BEGIN_MESSAGE_MAP(CAdvanceDlg, CDialog)
	ON_BN_CLICKED(IDOK, &CAdvanceDlg::OnBnClickedOk)

END_MESSAGE_MAP()


// CAdvanceDlg message handlers


void CAdvanceDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	m_copyrightEdit.GetWindowText(CChmConfig::GetInstance()->m_strCopyright);

	CChmConfig::GetInstance()->m_bHeader = m_bHeader;
	CChmConfig::GetInstance()->m_bFooter = m_bFooter;
	CChmConfig::GetInstance()->m_bPNG = m_bPNG;
	CChmConfig::GetInstance()->m_bRelatedTopics = m_bRelated;
	CChmConfig::GetInstance()->m_bListNumber = m_bListNumber;
	CDialog::OnOK();
}




BOOL CAdvanceDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// TODO:  Add extra initialization here
	m_copyrightEdit.SetWindowText(CChmConfig::GetInstance()->m_strCopyright);
	m_bHeader = CChmConfig::GetInstance()->m_bHeader;
	m_bFooter = CChmConfig::GetInstance()->m_bFooter;
	m_bPNG = CChmConfig::GetInstance()->m_bPNG;
	m_bRelated = CChmConfig::GetInstance()->m_bRelatedTopics;
	m_bListNumber = CChmConfig::GetInstance()->m_bListNumber;
	UpdateData(FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}
