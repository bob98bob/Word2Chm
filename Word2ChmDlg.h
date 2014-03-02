// Word2ChmDlg.h : 头文件
//

#pragma once
#include "afxwin.h"


// CWord2ChmDlg 对话框
class CWord2ChmDlg : public CDialog
{
// 构造
public:
	CWord2ChmDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_WORD2CHM_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedConvert();
	afx_msg void OnBnClickedOk();

protected:

	static unsigned __stdcall WordToChmProcess(LPVOID param);
	
public:
	CString m_strTempDir;
	CString m_strWord;
	CString m_strChm;
	CString m_strChmTitle;
	afx_msg void OnBnClickedButtonViewChm();
	afx_msg void OnBnClickedButtonWordBrowse();
	afx_msg void OnBnClickedButtonChmBrowse();
	CButton m_BtnGenChm;
	CButton m_BtnViewChm;

	BOOL	m_bRegistered;
	BOOL    m_bFinished;
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg void OnBnClickedButtonAbout();
	afx_msg void OnBnClickedButtonRegister();
	CButton m_registerBtn;
	afx_msg void OnBnClickedButtonHelp();
	afx_msg void OnClose();
	afx_msg void OnBnClickedButtonAdvance();
};
