#pragma once

#ifdef _WIN32_WCE
#error "Windows CE ��֧�� CDialog��"
#endif 

// CRegisterDialog �Ի���

class CRegisterDialog : public CDialog
{
	DECLARE_DYNCREATE(CRegisterDialog)

public:
	CRegisterDialog(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CRegisterDialog();

	CString GetSerialNumber();
// ��д

// �Ի�������
	enum { IDD = IDD_DIALOG_REGISTER};

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��
	virtual BOOL OnInitDialog();

	DECLARE_MESSAGE_MAP()
	
public:
	afx_msg void OnNMClickSyslink1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedOk();
	CString m_strSerialNumber;
	CString m_strMachineCode;
};
