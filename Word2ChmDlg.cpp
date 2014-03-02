// Word2ChmDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Word2Chm.h"
#include "Word2ChmDlg.h"
#include "RegisterDialog.h"
#include "OfficeWord.h"
#include "process.h"
#include "stdlib.h"
#include "SoftRegisterManager.h"
#include "AdvanceDlg.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define CONVERT_ENENT            1

// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

	// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
public:
	
	afx_msg void OnNMClickSyslink1(NMHDR *pNMHDR, LRESULT *pResult);
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	ON_NOTIFY(NM_CLICK, IDC_SYSLINK1, &CAboutDlg::OnNMClickSyslink1)
END_MESSAGE_MAP()


// CWord2ChmDlg �Ի���


const int maxUnRegisteredFileSize = 1024*1024;	//1M
const CString validSerialNumber = "w2c";

CWord2ChmDlg::CWord2ChmDlg(CWnd* pParent /*=NULL*/)
: CDialog(CWord2ChmDlg::IDD, pParent)
, m_strWord(_T(""))
, m_strChm(_T(""))
, m_strChmTitle(_T(""))
{
	m_bRegistered = FALSE;
	m_bFinished = TRUE;
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CWord2ChmDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_WORD, m_strWord);
	DDX_Text(pDX, IDC_EDIT_CHM, m_strChm);
	DDX_Text(pDX, IDC_EDIT_TITLE, m_strChmTitle);
	DDX_Control(pDX, IDC_BUTTON_CREATE_CHM, m_BtnGenChm);
	DDX_Control(pDX, IDC_BUTTON_VIEW_CHM, m_BtnViewChm);
	DDX_Control(pDX, IDC_BUTTON_REGISTER, m_registerBtn);
}

BEGIN_MESSAGE_MAP(CWord2ChmDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON_CREATE_CHM, &CWord2ChmDlg::OnBnClickedConvert)
	ON_BN_CLICKED(IDOK, &CWord2ChmDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDC_BUTTON_VIEW_CHM, &CWord2ChmDlg::OnBnClickedButtonViewChm)
	ON_BN_CLICKED(IDC_BUTTON_WORD_BROWSE, &CWord2ChmDlg::OnBnClickedButtonWordBrowse)
	ON_BN_CLICKED(IDC_BUTTON_CHM_BROWSE, &CWord2ChmDlg::OnBnClickedButtonChmBrowse)
	ON_WM_TIMER()
	ON_BN_CLICKED(IDC_BUTTON_ABOUT, &CWord2ChmDlg::OnBnClickedButtonAbout)
	ON_BN_CLICKED(IDC_BUTTON_REGISTER, &CWord2ChmDlg::OnBnClickedButtonRegister)
	ON_BN_CLICKED(IDC_BUTTON_HELP, &CWord2ChmDlg::OnBnClickedButtonHelp)
	ON_WM_CLOSE()
	ON_BN_CLICKED(IDC_BUTTON_ADVANCE, &CWord2ChmDlg::OnBnClickedButtonAdvance)
END_MESSAGE_MAP()


// CWord2ChmDlg ��Ϣ�������

BOOL CWord2ChmDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	CString appPath = AfxGetApp()->m_pszHelpFilePath;
	appPath = appPath.Left(appPath.ReverseFind('\\'));

	m_strTempDir.Format(_T("%s\\Temp"), appPath);
	
	CreateDirectory(m_strTempDir, NULL);
	SetCurrentDirectory(m_strTempDir);

	//ע����

	HKEY hKEY;
	long ret = (::RegOpenKeyEx(HKEY_LOCAL_MACHINE,"Software\\Word2Chm\\",0,KEY_READ,&hKEY));
	if(ret != ERROR_SUCCESS)
	{
	}

	//ȡ��ע�����к�
	BYTE serial[128] = {0};
	DWORD type = REG_SZ;//������������
	DWORD cbData = 128; //�������ݳ���

	ret = ::RegQueryValueEx(hKEY,"SerialNumber",NULL,&type,serial,&cbData);
	if(ret != ERROR_SUCCESS)
	{
	}

	::RegCloseKey(hKEY);

	CString serialNumber(serial);
	CSoftRegisterManager regManager;

	m_bRegistered = regManager.IsValid(serialNumber);

	if(m_bRegistered)
	{	
		m_registerBtn.ShowWindow(SW_HIDE);
	}

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CWord2ChmDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CWord2ChmDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CWord2ChmDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CWord2ChmDlg::OnBnClickedOk()
{
	if(!m_bFinished)
	{
		int code = MessageBox("���ڽ���ת������ȷʵҪ�˳���", "����", MB_YESNO);
		if(IDYES != code) return;
	}

	OnOK();
}


void CWord2ChmDlg::OnBnClickedButtonViewChm()
{
	if(m_strChm.IsEmpty())
	{
		MessageBox("��ѡ��Word�ļ�ת��!");
		return;
	}
	//WinExec(m_strChm, SW_SHOWNORMAL);
 
    ShellExecute(NULL,_T("OPEN"),m_strChm,NULL,NULL,SW_SHOWNORMAL);
}

void CWord2ChmDlg::OnBnClickedButtonWordBrowse()
{

	static TCHAR BASED_CODE szFilter[] = _T("Word Files (*.doc;*.docx)|*.doc;*docx|");
	//CFileDialog dlgFile(TRUE,NULL,NULL,NULL,szFilter,NULL);

	CFileDialog dlgFile(TRUE, NULL, NULL,OFN_FILEMUSTEXIST| OFN_HIDEREADONLY, szFilter);

	char drive[32];
	char dir[1024];
	char fileName[1024];
	if(dlgFile.DoModal() == IDOK)
	{
		m_strWord = dlgFile.GetPathName();
		_splitpath_s(LPCTSTR(m_strWord),drive,sizeof(drive)-1,dir, sizeof(dir)-1, fileName, sizeof(fileName)-1, NULL,0);
		m_strChm.Format(_T("%s\%s\%s.chm"), drive, dir, fileName);
		m_strChmTitle = fileName;
		UpdateData(FALSE);
	}

}

void CWord2ChmDlg::OnBnClickedButtonChmBrowse()
{
	static TCHAR BASED_CODE szFilter[] = _T("CHM Files (*.chm)|*.chm|");
	CFileDialog dlgFile(FALSE,NULL,NULL,NULL,szFilter,NULL);

	CString ext;
	if(dlgFile.DoModal() == IDOK)
	{
		m_strChm = dlgFile.GetPathName();
		ext = dlgFile.GetFileExt();
		ext.MakeLower();
		if(ext != _T("chm"))
			m_strChm += ".chm";
		UpdateData(FALSE);
	}
}

void CWord2ChmDlg::OnBnClickedConvert()
{

	if(m_strWord.IsEmpty())
	{
		MessageBox("��ѡ��Word�ļ�!");
		return;
	}
	
	//if(!m_bRegistered)
	//{
	//	HANDLE hFile=CreateFile(m_strWord,GENERIC_READ,0,NULL,OPEN_EXISTING,0,NULL);
	//	long nFileSize=GetFileSize(hFile,NULL);
	//	CloseHandle(hFile);

	//	if(nFileSize > maxUnRegisteredFileSize)
	//	{
	//		MessageBox("δע��汾ֻ�ܴ���С��1M���ļ�!");
	//		return;
	//	}

	//}
	
	UpdateData(TRUE);
	_DeleteFile(m_strChm);
	_DeleteFile(m_strTempDir + "\\*.*");

	m_BtnGenChm.EnableWindow(FALSE);
	m_BtnViewChm.EnableWindow(FALSE);
	SetTimer(1, 1000,NULL);
	unsigned threadID;
	_beginthreadex(NULL, 0, &CWord2ChmDlg::WordToChmProcess, this, 0, &threadID);

}

unsigned __stdcall CWord2ChmDlg::WordToChmProcess(LPVOID param)
{
	CWord2ChmDlg* pWord2ChmDlg = (CWord2ChmDlg*)param;
	ASSERT(pWord2ChmDlg);
	pWord2ChmDlg->m_bFinished = FALSE;
	CoInitialize(NULL);
	
	BOOL b = FALSE;

	try
	{

		COfficeWord wordParser(pWord2ChmDlg->m_strWord, pWord2ChmDlg->m_strTempDir);
		wordParser.SetRegistered(pWord2ChmDlg->m_bRegistered);
		b = wordParser.StartWord();
		if(!b)
		{
			AfxMessageBox("����wordʧ��!");
			goto end;
		}

		b = wordParser.GenerateChmHelp(pWord2ChmDlg->m_strChmTitle,pWord2ChmDlg->m_strChm);
		if(!b)
		{
			AfxMessageBox("word�ĵ���û�д����ʽ!");
			goto end;
		}

	}
	catch(COleDispatchException* pe)
	{
		AfxMessageBox(pe->m_strDescription);
		pe->Delete();
	}

end:
	pWord2ChmDlg->m_BtnGenChm.EnableWindow(TRUE);
	pWord2ChmDlg->m_BtnViewChm.EnableWindow(TRUE);
	pWord2ChmDlg->KillTimer(1);
	CStatic* pStatic = (CStatic*)pWord2ChmDlg->GetDlgItem(IDC_STATIC_INFO);
	pStatic->SetWindowTextA("");
	pWord2ChmDlg->m_bFinished = TRUE;
	if(b)
	{
		pWord2ChmDlg->OnBnClickedButtonViewChm();
	}
	return 0;
}
void CWord2ChmDlg::OnTimer(UINT_PTR nIDEvent)
{
	static int i = 0;
	if(nIDEvent == 1)
	{
		CString strInfo;
		strInfo = _T("����ת��");
		CStatic* pStatic = (CStatic*)GetDlgItem(IDC_STATIC_INFO);
		switch(i)
		{
		case 0:
			break;
		case 1:
			strInfo += ".";
			break;
		case 2:
			strInfo += ". .";
			break;
		case 3:
			strInfo += ". . .";
			break;
		case 4:
			strInfo += ". . . .";
			break;
		case 5:
			strInfo += ". . . . .";
			break;
		case 6:
			strInfo += ". . . . . .";
			break;
		}

		pStatic->SetWindowTextA(strInfo);
		i++;
		if(i==7)
			i = 1;
		
	}

	CDialog::OnTimer(nIDEvent);
}

void CWord2ChmDlg::OnBnClickedButtonAbout()
{
	CAboutDlg dlg;
	dlg.DoModal();
}


void CAboutDlg::OnNMClickSyslink1(NMHDR *pNMHDR, LRESULT *pResult)
{

	ShellExecute(NULL, "open", "http://blog.csdn.net/bob98", NULL, NULL, SW_SHOW);
	*pResult = 0;
}

void CWord2ChmDlg::OnBnClickedButtonRegister()
{
	CRegisterDialog  dlg;
	if(dlg.DoModal() == IDOK)
	{
		CString strSerial = dlg.GetSerialNumber();

		CSoftRegisterManager regManager;
		
		m_bRegistered = regManager.IsValid(strSerial);

		if(m_bRegistered)
		{	
			m_registerBtn.ShowWindow(SW_HIDE);
		}
	
		if(m_bRegistered)
		{

			HKEY hKEY;
			DWORD dw;

			long ret = (::RegCreateKeyEx(HKEY_LOCAL_MACHINE,"Software\\Word2Chm\\",0,
				NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&hKEY,&dw));
			if(ret != ERROR_SUCCESS)
			{
			}

			//long ret = (::RegCreateKeyEx(HKEY_LOCAL_MACHINE,"Software\\",0,
				//NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&hKEY,&dw));
			//ret = ::RegSetKeyValue(hKEY,"Word2Chm", "SerialNumber", REG_SZ, strSerial, strSerial.GetLength());
			ret = RegSetValueEx(
						  hKEY,
						  "SerialNumber",
						  0,
						  REG_SZ,
						  (BYTE*)strSerial.GetBuffer(),
						  strSerial.GetLength()
						);

			strSerial.ReleaseBuffer();
			if(ret != ERROR_SUCCESS)
			{
			}

			::RegCloseKey(hKEY);

			MessageBox("ע��ɹ�!", "", MB_OK);

		}
		else
		{
			MessageBox("ע��ʧ�ܣ��빺��!", "", MB_ICONEXCLAMATION);
		}
	}
}

void CWord2ChmDlg::OnBnClickedButtonHelp()
{
	CString appPath = AfxGetApp()->m_pszHelpFilePath;
	appPath = appPath.Left(appPath.ReverseFind('\\') + 1);
	CString strCmdHelp;

	//strCmdHelp.Format(_T("hh.exe %c%s\\word2chm.chm%c"), '"', appPath, '"');//notepad.exe
	//WinExec(strCmdHelp, SW_SHOW);

	strCmdHelp.Format(_T("%c%s\\word2chm.chm%c"), '"', appPath, '"');

	//SHELLEXECUTEINFO       ShExecInfo       =       {0};       
	//ShExecInfo.cbSize       =       sizeof(SHELLEXECUTEINFO);       
	//ShExecInfo.fMask       =       SEE_MASK_NOCLOSEPROCESS;       
	//ShExecInfo.hwnd       =       NULL;       
	//ShExecInfo.lpVerb       =       NULL;       
	//ShExecInfo.lpFile       =       strCmdHelp;       //�ļ�·�� 
	//ShExecInfo.lpParameters       =       " ";       
	//ShExecInfo.lpDirectory       =       NULL;       
	//ShExecInfo.nShow       =       SW_SHOW;       
	//ShExecInfo.hInstApp       =       NULL;       
	//ShellExecuteEx(&ShExecInfo);  

	ShellExecute(NULL,_T("OPEN"),strCmdHelp,NULL,NULL,SW_SHOWNORMAL);
}


void CWord2ChmDlg::OnClose()
{
	// TODO: Add your message handler code here and/or call default
	if(!m_bFinished)
	{
		int code = MessageBox("���ڽ���ת������ȷʵҪ�˳���", "����", MB_YESNO);
		if(IDYES != code) return;
	}

	CDialog::OnClose();
}


void CWord2ChmDlg::OnBnClickedButtonAdvance()
{
	// TODO: Add your control notification handler code here
	CAdvanceDlg advanceDlg;
	advanceDlg.DoModal();
}
