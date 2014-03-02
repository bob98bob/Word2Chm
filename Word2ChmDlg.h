// Word2ChmDlg.h : ͷ�ļ�
//

#pragma once
#include "afxwin.h"


// CWord2ChmDlg �Ի���
class CWord2ChmDlg : public CDialog
{
// ����
public:
	CWord2ChmDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_WORD2CHM_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
