#pragma once
#include "afxwin.h"


// CAdvanceDlg dialog

class CAdvanceDlg : public CDialog
{
	DECLARE_DYNAMIC(CAdvanceDlg)

public:
	CAdvanceDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CAdvanceDlg();

// Dialog Data
	enum { IDD = IDD_DIALOG_ADVANCE };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CEdit m_copyrightEdit;
	afx_msg void OnBnClickedOk();


	BOOL m_bHeader;
	BOOL m_bFooter;
	virtual BOOL OnInitDialog();
	BOOL m_bPNG;
	BOOL m_bRelated;
	BOOL m_bListNumber;
};
