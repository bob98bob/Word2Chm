// RegisterDialog.cpp : 实现文件
//

#include "stdafx.h"
#include "Word2Chm.h"
#include "RegisterDialog.h"
#include "SoftRegisterManager.h"

// CRegisterDialog 对话框

IMPLEMENT_DYNCREATE(CRegisterDialog, CDialog)

CRegisterDialog::CRegisterDialog(CWnd* pParent /*=NULL*/)
	: CDialog(CRegisterDialog::IDD, pParent)
	, m_strSerialNumber(_T(""))
	, m_strMachineCode(_T(""))
{

}

CRegisterDialog::~CRegisterDialog()
{
}

void CRegisterDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_SERIAL, m_strSerialNumber);
	DDX_Text(pDX, IDC_EDIT_MACHINE, m_strMachineCode);
}

BOOL CRegisterDialog::OnInitDialog()
{
	CDialog::OnInitDialog();
	CSoftRegisterManager regManager;
	m_strMachineCode = regManager.GenerateMachineCode();
	UpdateData(FALSE);
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

BEGIN_MESSAGE_MAP(CRegisterDialog, CDialog)
	ON_NOTIFY(NM_CLICK, IDC_SYSLINK1, &CRegisterDialog::OnNMClickSyslink1)
	ON_BN_CLICKED(IDOK, &CRegisterDialog::OnBnClickedOk)
END_MESSAGE_MAP()


// CRegisterDialog 消息处理程序


void CRegisterDialog::OnNMClickSyslink1(NMHDR *pNMHDR, LRESULT *pResult)
{
	ShellExecute(NULL, "open", "http://www.crsky.com/soft/24851.html", NULL, NULL, SW_SHOW);
	*pResult = 0;
}

CString CRegisterDialog::GetSerialNumber()
{
	return m_strSerialNumber;
}
void CRegisterDialog::OnBnClickedOk()
{
	UpdateData(TRUE);
	
	
	OnOK();
}
