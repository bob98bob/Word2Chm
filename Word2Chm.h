// Word2Chm.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CWord2ChmApp:
// �йش����ʵ�֣������ Word2Chm.cpp
//

class CWord2ChmApp : public CWinApp
{
public:
	CWord2ChmApp();

// ��д
	public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CWord2ChmApp theApp;