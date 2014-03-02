#pragma once

#ifdef _WIN32_WCE
#error "Windows CE 不支持 CDialog。"
#endif 

// CRegisterDialog 对话框

class CRegisterDialog : public CDialog
{
	DECLARE_DYNCREATE(CRegisterDialog)

public:
	CRegisterDialog(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CRegisterDialog();

	CString GetSerialNumber();
// 重写

// 对话框数据
	enum { IDD = IDD_DIALOG_REGISTER};

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持
	virtual BOOL OnInitDialog();

	DECLARE_MESSAGE_MAP()
	
public:
	afx_msg void OnNMClickSyslink1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedOk();
	CString m_strSerialNumber;
	CString m_strMachineCode;
};
