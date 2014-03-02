/*
*  word2chm
*  word 转 chm 工具
*	
*  目的  : 编写chm格式的文件很是麻烦，但我们都习惯于写word文档，所以做一个转换工具会方便很多。
*  目的2 : 了解word api、chm原理。
*
*  日期  : 2009-11-06
*  作者  : 鲍龙洋 
*
*/
#pragma once

#include "CApplication.h"
#include "CDocuments.h"
#include "CWordDocument.h"
#include "CParagraphs.h"
#include "CParagraph.h"
#include "CRange.h"
#include "CListFormat.h"
#include "CHyperlinks.h"
#include "CHyperlink.h"
#include <vector>
#include "HtmlAddin.h"

using namespace std;
BOOL _DeleteFile(CString szFileOrFolder);

class COfficeWord;
class COutlineTreeItem
{
public:
	friend COfficeWord;
	COutlineTreeItem()
	{
		_parentItem = NULL;
		_firstChildItem = NULL;
		_nextItem = NULL;
		_paragraph = NULL;
	}
	~COutlineTreeItem()
	{

	}
	
private:
	COutlineTreeItem* _parentItem;
	COutlineTreeItem* _firstChildItem;
	COutlineTreeItem* _nextItem;
	CParagraph		  _paragraph;
	CString			  _htmlFile;
	UINT			  _pageIndex; // generate html page index

};

typedef vector<COutlineTreeItem*> ItemArray;
typedef vector<CString> Tokens;
typedef vector<CString> Files;

class COfficeWord
{
public:
	COfficeWord(CString strDoc, CString htmlDir);
	~COfficeWord(void);
	BOOL	StartWord();
	BOOL	GenerateChmHelp(CString strChmTitle, CString strChmFile);
	void	SetRegistered(BOOL bRegistered);
protected:
	void    Release();
	void	GenerateOutlineTree();
	BOOL	GenerateHtmlFiles();
	void    GenerateItemArray(COutlineTreeItem* item, ItemArray& itemArr);
	void    GenerateHHC_UL_LI_Tokens(COutlineTreeItem* item, Tokens& tokens);
	void    GenerateHHK_UI_LI_Tokens(COutlineTreeItem* item, Tokens& tokens);
	void	GenerateLI_Tokens(COutlineTreeItem* item, Tokens& tokens);
	BOOL	GenerateHHC(CString strhhc);
	BOOL    GenerateHHK(CString strhhk);
	BOOL	GenerateHHP(CString strhhp, CString strhhc, CString strhhk, CString strchm);

	//      Hyperlinks 
	void	GenerateHyperlinks(CRange range);
	CString ConvertInternalHyperlinkToExternalHyperlink(CString strHyperlink);
	CRange	GetBookmarkRange(CString strBookmark);
	int     GetHtmlPageFromRange(CRange range);
	void    RemoveUnderlineOfHyperlinks(CWordDocument doc);
	void	GenerateRelatedTopics(CWordDocument doc, COutlineTreeItem* pItem);
	
	//		Paragraph space
	void	SetParagraphSpaceAfterAndBefore(CParagraph paragraph, float after, float before);
	void	SetParagraphLineSpace(CParagraph paragraph, int iLineSpacingRule, float space);
	void    SetDocumentSingleLineSpace(CWordDocument doc);

	//      List Number
	void	RemoveListNumber(CWordDocument doc);
	void	GenerateUnRegisteredFootnotes(CWordDocument doc);

	void    GenerateCopyright(CWordDocument doc);
private:
	CString			_htmlDirectory;
	CString			_strDoc;
	CApplication	_wordApp;
	CWordDocument	_wordDoc;
	COutlineTreeItem* _outlineTree;
	BOOL			_bRemvoeList;
	BOOL			_bWordExist;
	CString			_strTitle;
	Files           _files;    //html files
	BOOL			_bHyperlinkUnderline;
	BOOL			_bRelatedTopics;
	CString			_strRelatedTopicsTitle;
	BOOL			_bRegistered;

	CHtmlAddinsManager  _htmlAddinsManager;
};

typedef enum
{
	wdLineSpaceSingle,
	wdLineSpace1pt5,
	wdLineSpaceDouble,
	wdLineSpaceAtLeast,
	wdLineSpaceExactly,
	wdLineSpaceMultiple

}LineSpacingRule;

class CChmConfig
{
protected:
	CChmConfig()
	{
		m_bHeader = FALSE;
		m_bFooter = FALSE;
		m_bPNG = TRUE;
		m_bRelatedTopics = TRUE;
		m_bListNumber = FALSE;
	}

public:
	static CChmConfig* GetInstance()
	{
		static CChmConfig chmConfig;
		return &chmConfig;
	}

	CString m_strCopyright;
	BOOL	m_bHeader;
	BOOL	m_bFooter;
	BOOL    m_bPNG;
	BOOL    m_bRelatedTopics;
	BOOL	m_bListNumber;
};


