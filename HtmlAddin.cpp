#include "StdAfx.h"
#include "HtmlAddin.h"


CHtmlAddin::CHtmlAddin(void)
{
}


CHtmlAddin::~CHtmlAddin(void)
{
}

CHtmlAddinsManager::CHtmlAddinsManager()
{

}

CHtmlAddinsManager::~CHtmlAddinsManager()
{

}

void CHtmlAddinsManager::Process(CString strHtmlFile)
{
	
	CString strHtml;

	for(int i=0; i<m_addins.size(); i++)
	{
		m_addins[i]->Process(strHtml);
	}
}


void CTopAddin::Process(CString & strHtml)
{

}

void CBottomAddin::Process(CString & strHtml)
{

}