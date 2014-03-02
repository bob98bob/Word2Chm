#pragma once
#include <vector>
using namespace std;

class CHtmlAddin
{
public:
	CHtmlAddin(void);
	~CHtmlAddin(void);
	virtual void Process(CString & strHtml) = 0;
};

class CHtmlAddinsManager
{
public:
	CHtmlAddinsManager();
	~CHtmlAddinsManager();
	void Process(CString strHtmlFile);

protected:
	vector<CHtmlAddin*> m_addins;
};


class CTopAddin : public CHtmlAddin
{
public:
	CTopAddin():CHtmlAddin(){}
	void Process(CString & strHtml);
protected:
	CString m_strContent;
};

class CBottomAddin : public CHtmlAddin
{
public:
	CBottomAddin():CHtmlAddin(){}
	void Process(CString & strHtml);
protected:
	CString m_strContent;
};
