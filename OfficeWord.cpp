/*
*  word2chm
*  word 转 chm 工具
*	
*  目的  : 编写chm格式的文件很是麻烦，但我们都习惯于写word文档，所以做一个转换工具会方便很多。
*  目的2 : 了解word api、chm原理。
*
*
*  日期  : 2009-11-06
*  作者  : 鲍龙洋 
*
*/

/*  log
   
   2011-02-10   增加文档内部超链接、锚点支持
   2011-04-24	增加页眉、页脚支持
   2011-11-13	增加PNG高质量图片、解决图片不显示问题

*/

#include "StdAfx.h"
#include "OfficeWord.h"
#include "CBookmark0.h"
#include "CBookmarks.h"
#include "CStyles.h"
#include "CStyle.h"
#include "CFont0.h"
#include "CParagraphFormat.h"
#include "CSelection.h"
#include "CSection.h"
#include "CSections.h"
#include "CHeaderFooter.h"
#include "CHeadersFooters.h"
#include "CWebOptions.h"
#include "CnlineShapes.h"
#include "CnlineShape.h"
#include "CHorizontalLineFormat.h"
#include "CList0.h"

#define  UNREGISTER_FOOTNOTES		_T("这是未注册版本,请购买注册！")

char Separator = 12; //分节符标记
BOOL _DeleteFile(CString szFileOrFolder);
void _SearchDirFiles(CString strDir, Files& files);
/*
HHA_CompileHHP(const char*, LPVOID, LPVOID, int)
第一个参数为你要编译的hhp文件名
第二个参数是编译日志回调函数
第三个参数是编译进度回调函数
第四个参数不知道什么用，我把他置为0
*/

typedef BOOL (WINAPI *HHA_CompileHHP)(const char*, LPVOID, LPVOID, int);
BOOL CALLBACK FunLog(char* pstr);
BOOL CALLBACK FunProc(char* pstr);

COfficeWord::COfficeWord(CString strDoc, CString htmlDir)
{
	_strDoc = strDoc;
	_strTitle = _T("doc");
	_wordApp  = NULL;
	_wordDoc  = NULL;
	_bRemvoeList = TRUE;
	_bWordExist = FALSE;
	_htmlDirectory = htmlDir;
	_outlineTree = new COutlineTreeItem();
	_bHyperlinkUnderline = FALSE;
	_bRelatedTopics = TRUE;
	_bRegistered = FALSE;
	//CoInitialize(NULL);
}

COfficeWord::~COfficeWord(void)
{
	COleVariant covTrue((short)TRUE);
	COleVariant covFalse((short)FALSE);
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	Release();
	_wordDoc.Close(covFalse, covOptional, covOptional);
	//_wordDoc.ReleaseDispatch();
	//if(!_bWordExist)
	{
		_wordApp.Quit(covFalse, covTrue, covFalse);
	}
	//_wordApp.ReleaseDispatch();
	delete _outlineTree;
	//CoUninitialize();
}

void COfficeWord::Release()
{
	ItemArray	itemArr;
	GenerateItemArray(_outlineTree->_firstChildItem, itemArr);
	for(int i=0; i<(int)itemArr.size(); i++)
	{
		itemArr[i]->_paragraph.DetachDispatch();
		delete itemArr[i];
	}

}

BOOL COfficeWord::StartWord()
{
	//CLSID   clsid;
	//HRESULT		hr;
	//IUnknown    *pUnk;
	//IDispatch   *pDisp;

	//CLSIDFromProgID( L"Word.Application",   &clsid );

	//hr = GetActiveObject( clsid,   NULL,   (IUnknown**)&pUnk );
	//if ( !FAILED(hr) ){
	//	hr = pUnk->QueryInterface( IID_IDispatch,  (void   **)&pDisp );
	//	ASSERT( !FAILED(hr) );
	//	_wordApp.AttachDispatch( pDisp, TRUE );
	//	_bWordExist = TRUE;
	//	pUnk -> Release();

	//}
	//else
	{
		COleException pe;
		if(!_wordApp.CreateDispatch(_T("Word.Application"), &pe))
		{
			pe.ReportError();
			return FALSE;
		}

	}
	_wordApp.put_Visible(FALSE);

	CDocuments documents = _wordApp.get_Documents();

	COleVariant covTrue((short)TRUE);
	COleVariant covFalse((short)FALSE);
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant covFile(_strDoc, VT_BSTR);

	_wordDoc = documents.Open2002(covFile, covFalse, covTrue,
		covFalse, covOptional, covOptional,
		covFalse, covOptional, covOptional,
		covOptional, covOptional,covTrue,
		covOptional,covOptional,covOptional);

	return TRUE;
}

/*
 生成大纲目录树
*/
void COfficeWord::GenerateOutlineTree()
{
	int iPageNumber = 0;
	CParagraphs paragraphs = _wordDoc.get_Paragraphs();
	int count = paragraphs.get_Count();

	CParagraph paragraph;
	COutlineTreeItem* prevItem = _outlineTree;
	for(int i=1; i<=count; i++)
	{
		paragraph = paragraphs.Item(i);
		if(paragraph.get_OutlineLevel() >= 10)
		{
			continue;
		}

		CString strSeparator(Separator);
		CString strText = CRange(paragraph.get_Range()).get_Text();
		if(strText == strSeparator)
		{
			continue;
		}

		COutlineTreeItem* pNewItem = new COutlineTreeItem();
		pNewItem->_paragraph = paragraph;
		pNewItem->_pageIndex = iPageNumber++;

		if( prevItem == _outlineTree )
		{
			prevItem->_firstChildItem = pNewItem;
			pNewItem->_parentItem = prevItem;
		}
		else
		{
			int prevLevel = prevItem->_paragraph.get_OutlineLevel();
			int currLevel = pNewItem->_paragraph.get_OutlineLevel();

			if(currLevel == prevLevel)
			{
				prevItem->_nextItem = pNewItem;
				pNewItem->_parentItem = prevItem->_parentItem;
			}

			if(currLevel > prevLevel)
			{
				prevItem->_firstChildItem = pNewItem;
				pNewItem->_parentItem = prevItem;
			}

			if(currLevel < prevLevel)
			{
				COutlineTreeItem* pPrevItem = prevItem;
				COutlineTreeItem* pItem = prevItem->_parentItem;
				while(pItem)
				{
					if(pItem == _outlineTree)
					{
						break;
					}
					if(pItem->_paragraph.get_OutlineLevel() < currLevel)
					{
						break;
					}
					if(pItem->_paragraph.get_OutlineLevel() == currLevel)
					{
						pPrevItem = pItem;
						break;
					}	
					pPrevItem = pItem;
					pItem = pItem->_parentItem;
				}
				pPrevItem->_nextItem = pNewItem;
				pNewItem->_parentItem = pPrevItem->_parentItem;
			}
		}
		prevItem = pNewItem;
	}
}

void  COfficeWord::GenerateItemArray(COutlineTreeItem* item, ItemArray& itemArr)
{
	if(!item)
	{
		ASSERT(item);
		return;
	}

	itemArr.push_back(item);
	if(item->_firstChildItem)
	{
		GenerateItemArray(item->_firstChildItem, itemArr);
	}
	if(item->_nextItem)
	{
		GenerateItemArray(item->_nextItem, itemArr);
	}
}

BOOL COfficeWord::GenerateHtmlFiles()
{
	ItemArray	itemArr;
	GenerateItemArray(_outlineTree->_firstChildItem, itemArr);

	_files.clear();
	CDocuments documents = _wordApp.get_Documents();

	COleVariant covTrue((short)TRUE);
	COleVariant covFalse((short)FALSE);
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	CWordDocument doc = documents.Add(covOptional, covFalse,  covOptional, covTrue);
	// remove hyperlinks underlines
	if(!_bHyperlinkUnderline)
		RemoveUnderlineOfHyperlinks(doc);

	COutlineTreeItem* pItem;
	CParagraph  paragraph;
	CRange range;

	int count = itemArr.size();
	for(int i=0; i<(int)itemArr.size(); i++)
	{
		pItem = itemArr[i];
		range = pItem->_paragraph.get_Range();
		paragraph = pItem->_paragraph.Next(covOptional);

		CParagraphs docParagraphs = doc.get_Paragraphs();
		int count = docParagraphs.get_Count();
		COleVariant covDocStart((long)CRange(CParagraph(docParagraphs.Item(1)).get_Range()).get_Start());
		COleVariant covDocEnd((long)CRange(CParagraph(docParagraphs.Item(count)).get_Range()).get_End());
		CRange docRange = doc.Range(covDocStart, covDocEnd);
		docRange.Cut();
		docRange.ReleaseDispatch();

		if(!paragraph.m_lpDispatch)
		{

		}
		else if(paragraph.get_OutlineLevel() == 10)
		{			
			if(i<(int)itemArr.size() - 1)
			{
				CParagraph endParagraph = itemArr[i+1]->_paragraph.Previous(covOptional);
				CRange endRange = endParagraph.get_Range();
				COleVariant covStart((long)range.get_Start());
				COleVariant covEnd((long)endRange.get_End());
				range = _wordDoc.Range(covStart, covEnd);
				range.Copy();
				endParagraph.ReleaseDispatch();
				endRange.ReleaseDispatch();
			}
			else
			{
				CParagraphs  paragraphs(_wordDoc.get_Paragraphs());
				count = paragraphs.get_Count();
				CParagraph  par(paragraphs.Item(count));
				CRange endRange = par.get_Range();
				COleVariant covStart((long)range.get_Start());
				COleVariant covEnd((long)endRange.get_End());
				range = _wordDoc.Range(covStart, covEnd);
				range.Copy();
				paragraphs.ReleaseDispatch();
				par.ReleaseDispatch();
				endRange.ReleaseDispatch();
			}	
		}
		else
		{
			range.Copy();
		}


		CRange insertRange = CRange(CParagraph(docParagraphs.Item(1)).get_Range());
		insertRange.SetRange(0, 0);

		//header process
		if(CChmConfig::GetInstance()->m_bHeader)
		{
			CHeadersFooters Headers(CSection(CSections(range.get_Sections()).Item(1)).get_Headers());
			CHeaderFooter Header(Headers.Item(1));
			CRange(Header.get_Range()).Copy();
			insertRange.Paste();
		
			docParagraphs.Add(covOptional);
			docParagraphs.Add(covOptional);
			CRange rg(CParagraph(docParagraphs.Add(covOptional)).get_Range());
			insertRange.SetRange(rg.get_Start(), rg.get_End());
		}

		//paste contents
		range.Copy();
		insertRange.Paste();

		//adjust processing range
		insertRange.SetRange(0,
			CRange(CParagraph(docParagraphs.Item(docParagraphs.get_Count())).get_Range()).get_End());

		//Hyperlinks process
		GenerateHyperlinks(insertRange);

		//Related topics
		if(CChmConfig::GetInstance()->m_bRelatedTopics)
		{
			GenerateRelatedTopics(doc, pItem);
		}

		//Remove List Number
		if(!CChmConfig::GetInstance()->m_bListNumber)
		{
			RemoveListNumber(doc);
		}
		GenerateUnRegisteredFootnotes(doc);

		GenerateCopyright(doc);

		//footer process
		if(CChmConfig::GetInstance()->m_bFooter)
		{
			CRange rg(CParagraph(docParagraphs.Add(covOptional)).get_Range());
			insertRange.SetRange(rg.get_Start(), rg.get_End());

			CHeadersFooters Footers(CSection(CSections(range.get_Sections()).Item(1)).get_Footers());
			CHeaderFooter Footer(Footers.Item(1));
			CRange(Footer.get_Range()).Copy();
			insertRange.Paste();
		}

		//Line space
		SetDocumentSingleLineSpace(doc);

		CString htmlFile;
		htmlFile.Format(_T("%s\\%d.html"), _htmlDirectory, i);
		pItem->_htmlFile.Format(_T("%d.html"), i);

		_files.push_back(pItem->_htmlFile);
		
		//AllowPNG
		if(CChmConfig::GetInstance()->m_bPNG)
		{
			CWebOptions webOptions(doc.get_WebOptions());
			webOptions.put_AllowPNG(true);
		}

		COleVariant covFile(htmlFile, VT_BSTR);
		COleVariant fileFormat(long(10));
		doc.SaveAs(covFile, fileFormat, covOptional, covOptional, covFalse, covOptional,
			covOptional, covOptional, covOptional, covOptional, covOptional, covOptional,
			covOptional, covOptional, covOptional, covOptional);

		range.ReleaseDispatch();
		insertRange.ReleaseDispatch();
		paragraph.ReleaseDispatch();

		docParagraphs.ReleaseDispatch();
	}

	doc.Close(covFalse, covOptional, covOptional);
	doc.ReleaseDispatch();
	documents.ReleaseDispatch();

	return TRUE;
}

void COfficeWord::GenerateHHC_UL_LI_Tokens(COutlineTreeItem* item, Tokens& tokens)
{
	ASSERT(item);
	GenerateLI_Tokens(item,tokens);
	if(item->_firstChildItem)
	{
		tokens.push_back(_T("<UL>"));
		GenerateHHC_UL_LI_Tokens(item->_firstChildItem, tokens);
		tokens.push_back(_T("</UL>"));
	}

	COutlineTreeItem* next = item->_nextItem;
	while(next)
	{
		GenerateLI_Tokens(next,tokens);
		if(next->_firstChildItem)
		{
			tokens.push_back(_T("<UL>"));
			GenerateHHC_UL_LI_Tokens(next->_firstChildItem, tokens);
			tokens.push_back(_T("</UL>"));
		}
		next = next->_nextItem;
	}
}

void  COfficeWord::GenerateHHK_UI_LI_Tokens(COutlineTreeItem* item, Tokens& tokens)
{
	ASSERT(item);
	GenerateLI_Tokens(item,tokens);
	if(item->_firstChildItem)
	{
		GenerateHHK_UI_LI_Tokens(item->_firstChildItem, tokens);
	}

	COutlineTreeItem* next = item->_nextItem;
	while(next)
	{
		GenerateLI_Tokens(next,tokens);
		if(next->_firstChildItem)
		{
			GenerateHHK_UI_LI_Tokens(next->_firstChildItem, tokens);
		}
		next = next->_nextItem;
	}
}

void COfficeWord::GenerateLI_Tokens(COutlineTreeItem* item, Tokens& tokens)
{
	ASSERT(item);
	CString strLine;
	strLine.Format(_T("<LI> <OBJECT type=%ctext/sitemap%c>"), '"', '"');
	tokens.push_back(strLine);
	CRange range;
	range = item->_paragraph.get_Range();
	CString strText = range.get_Text();

	if(CChmConfig::GetInstance()->m_bListNumber)
	{
		CListFormat listFormat = range.get_ListFormat();
		strText = listFormat.get_ListString()+ " " + strText;
	}

	strLine.Format(_T("<param name=%cName%c value=%c%s%c>"), '"', '"', '"', strText, '"');
	tokens.push_back(strLine);
	strLine.Format(_T("<param name=%cLocal%c value=%c%s%c>"), '"', '"', '"', item->_htmlFile, '"');
	tokens.push_back(strLine);

	/* leaf node */
	if(!item->_firstChildItem) 
	{
		/* image index, 9 represent '?', 11 represent '-' */
		strLine.Format(_T("<param name=%cImageNumber%c value=%c%d%c>"), '"', '"', '"', 11, '"');
		tokens.push_back(strLine);
	}

	tokens.push_back(_T("</OBJECT>"));
	range.ReleaseDispatch();
}

BOOL COfficeWord::GenerateHHC(CString strhhc)
{
	CString strLine;
	Tokens tokens;
	tokens.push_back(_T("<HTML>"));
	tokens.push_back(_T("<HEAD>"));
	strLine.Format(_T("<meta http-equiv=Content-Type content=%ctext/html; charset=gb2312%c>"), '"', '"');
	tokens.push_back(strLine);
	strLine.Format(_T("  <meta name=%cGENERATOR%c content=%cWord-2-CHM%c>"), '"', '"', '"', '"');
	tokens.push_back(strLine);
	tokens.push_back(_T("<!-- Sitemap 1.0 -->"));
	tokens.push_back(_T("</HEAD>"));
	tokens.push_back(_T("<BODY>"));
	strLine.Format(_T("<OBJECT type=%ctext/site properties%c>"), '"', '"');
	tokens.push_back(strLine);
	strLine.Format(_T("	<param name=%cWindow Styles%c value=%c0x800235%c>"), '"', '"', '"', '"');
	tokens.push_back(strLine);
	tokens.push_back(_T("</OBJECT>"));
	tokens.push_back(_T("<UL>"));

	GenerateHHC_UL_LI_Tokens(_outlineTree->_firstChildItem, tokens);

	tokens.push_back(_T("</UL>"));
	tokens.push_back(_T("</BODY>"));
	tokens.push_back(_T("</HTML>"));

	CStdioFile file;
	CFileException ex;

	if (!file.Open(strhhc, CFile::modeCreate | CFile::modeWrite | CFile::typeText , &ex))
	{
		return FALSE;
	}

	for(int i=0; i<(int)tokens.size(); i++)
	{
		strLine = tokens[i];
		strLine.Remove('\n');
		strLine.Remove('\r');
		file.WriteString(strLine);
		file.WriteString(_T("\n"));
	}
	file.Close();

	return TRUE;
}
BOOL COfficeWord::GenerateHHK(CString strhhk)
{
	CString strLine;
	Tokens tokens;
	tokens.push_back(_T("<HTML>"));
	tokens.push_back(_T("<HEAD>"));
	strLine.Format(_T("<meta http-equiv=Content-Type content=%ctext/html; charset=gb2312%c>"), '"', '"');
	tokens.push_back(strLine);
	strLine.Format(_T("  <meta name=%cGENERATOR%c content=%cWord-2-CHM%c>"), '"', '"', '"', '"');
	tokens.push_back(strLine);
	tokens.push_back(_T("<!-- Sitemap 1.0 -->"));
	tokens.push_back(_T("</HEAD>"));
	tokens.push_back(_T("<BODY>"));
	tokens.push_back(_T("<UL>"));

	GenerateHHK_UI_LI_Tokens(_outlineTree->_firstChildItem, tokens);

	tokens.push_back(_T("</UL>"));
	tokens.push_back(_T("</BODY>"));
	tokens.push_back(_T("</HTML>"));

	CStdioFile file;
	CFileException ex;

	if (!file.Open(strhhk, CFile::modeCreate | CFile::modeWrite | CFile::typeText , &ex))
	{
		return FALSE;
	}

	for(int i=0; i<(int)tokens.size(); i++)
	{
		strLine = tokens[i];
		strLine.Remove('\n');
		strLine.Remove('\r');
		file.WriteString(strLine);
		file.WriteString(_T("\n"));
	}
	file.Close();

	return TRUE;
}
BOOL COfficeWord::GenerateHHP(CString strhhp, CString strhhc, CString strhhk, CString strchm)
{
	CStdioFile file;
	CFileException ex;
	CString strLine;

	if (!file.Open(strhhp, CFile::modeCreate | CFile::modeWrite | CFile::typeText , &ex))
	{
		return FALSE;
	}

	file.WriteString(_T("[OPTIONS]\n"));
	file.WriteString(_T("Compatibility=1.1 Or later\n"));
	strLine.Format(_T("Compiled file=%s\n"), strchm);
	file.WriteString(strLine);
	file.WriteString(_T("Default Window=Main\n"));

	file.WriteString(_T("Default topic=0.html\n"));
	file.WriteString(_T("Display compile progress=No\n"));
	file.WriteString(_T("Enhanced decompilation=Yes\n"));
	file.WriteString(_T("Full-text search=Yes\n"));
	strLine.Format(_T("Title=%s\n"), _strTitle);
	file.WriteString(strLine);

	strLine.Format(_T("Contents file=%s\n"), strhhc);
	file.WriteString(strLine);
	strLine.Format(_T("Index file=%s\n"), strhhk);
	file.WriteString(strLine);
	file.WriteString(_T("Default font=宋体,9,1\n"));
	file.WriteString(_T("Language=0x804 中文(中国)\n"));

	file.WriteString(_T("[WINDOWS]\n"));

	strLine.Format(_T("Main=%c%s%c,%c%s%c,%c%s%c,%c0.html%c,,,,,,0x62520,,0x300e,[135,50,1015,730],0x1030000,,,,,,0\n"),
		'"', _strTitle, '"', '"', strhhc, '"', '"', strhhk, '"', '"', '"');

	file.WriteString(strLine);
	file.WriteString(_T("[FILES]\n"));

	//html files
	for(int i=0; i<_files.size(); i++)
	{
		file.WriteString(_files[i]);
		file.WriteString("\n");
	}

	//image files
	CString imageFile;
	CString searchDir;
	Files files;
	for(int i=0; i<_files.size(); i++)
	{
		files.clear();
		searchDir.Format(_T("%s\\%d.files\\*.*"),_htmlDirectory,i);
		_SearchDirFiles(searchDir, files);
		for(int j=0; j<files.size(); j++)
		{
			imageFile.Format(_T("%d.files\\%s"), i,files[j]);
			file.WriteString(imageFile);
			file.WriteString("\n");
		}
	}

	file.Close();

	return TRUE;
}

BOOL COfficeWord::GenerateChmHelp(CString strChmTitle, CString strChmFile)
{
	if(_outlineTree->_firstChildItem)
	{
		Release();
	}

	GenerateOutlineTree();

	if(!_outlineTree->_firstChildItem)
	{
		return FALSE;
	}

	GenerateHtmlFiles();

	_strTitle = strChmTitle;
	CString strhhc, strhhk, strhhp;
	CString _strhhc, _strhhk, _strhhp;

	strhhc = strChmTitle + _T(".hhc");
	strhhk = strChmTitle + _T(".hhk");
	strhhp = strChmTitle + _T(".hhp");

	_strhhc.Format(_T("%s\\%s"), _htmlDirectory, strhhc);
	_strhhk.Format(_T("%s\\%s"), _htmlDirectory, strhhk);
	_strhhp.Format(_T("%s\\%s"), _htmlDirectory, strhhp);

	GenerateHHC(_strhhc);
	GenerateHHK(_strhhk);
	GenerateHHP(_strhhp, strhhc, strhhk, strChmFile);

	CString appPath = AfxGetApp()->m_pszHelpFilePath;
	appPath = appPath.Left(appPath.ReverseFind('\\') + 1);

	/*
	CString cmd;
	cmd.Format(_T("%s\\hhc.exe %c%s%c"), appPath, '"', _strhhp, '"'); 

	int retCode = WinExec(cmd, SW_SHOWMINNOACTIVE);
	if(retCode <= 31)
		return FALSE;
	*/

	HINSTANCE hinstLib; 
	BOOL fFreeResult, fRunTimeLinkSuccess = FALSE; 
	HHA_CompileHHP CompileFunc = NULL;
	
	hinstLib = LoadLibrary("hha.dll"); 
	
	if (hinstLib != NULL)
	{
		CompileFunc = (HHA_CompileHHP) GetProcAddress(hinstLib, "HHA_CompileHHP"); 
		
		LPCSTR pzFileNmae = _strhhp.GetBuffer(_strhhp.GetLength());
		if (fRunTimeLinkSuccess = (CompileFunc != NULL)) 
		{
			if(CompileFunc(pzFileNmae, FunLog, FunProc, 0))
			{
			}
		}

		fFreeResult = FreeLibrary(hinstLib); 
	}

	//Sleep(1000);
	//CString strTemp = _htmlDirectory + "\\*.*";
	//_DeleteFile(strTemp);
	return TRUE;
}

//////////////////////////////////////////////////////////////////
// 文    件 : 
// 函 数 名 : _DeleteFile
// 功能描述 : 删除目录或文件及子目录
// 调 用 者 : 
// 参    数 ：szFileOrFolder 支持带*.*
// 返 回 值 : BOOL 
//          
//----------------------------------------------------------------
// 作    者 : 鲍龙样
// 电子邮件 : baoly@neusoft.com
// 创建日期 : 2006-01-20
//----------------------------------------------------------------

BOOL _DeleteFile(CString szFileOrFolder)
{
	BOOL bRes = TRUE;
	WIN32_FIND_DATA wfData;
	HANDLE h = ::FindFirstFile(szFileOrFolder, &wfData);
	if(h == (HANDLE)0xffffffff)
	{
		return bRes = FALSE;
	}
	do {
		CString ingore = wfData.cFileName;
		CString srcFolder = szFileOrFolder;
		if(wfData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
		{
			if(ingore == "." || ingore == ".." ) continue;
			if(srcFolder.Find("*.*") != -1)
			{
				int index = srcFolder.Find("*.*");
				srcFolder = srcFolder.Left(index);
				srcFolder += wfData.cFileName;	
			}
			CString szFolder(srcFolder);
			srcFolder += "\\*.*";
			if(bRes = _DeleteFile(srcFolder))
			{	
				SetFileAttributes(srcFolder, FILE_ATTRIBUTE_ARCHIVE);
				if(!RemoveDirectory(szFolder))
				{
					CString error;
					error += "cannot delete directory: ";
					error += szFolder;
					AfxMessageBox(error);
					bRes = FALSE;
				}
			}	
		}
		else
		{
			CString szFile = szFileOrFolder;
			if(szFile.Find("*.*") != -1)
			{	
				int index = szFile.Find("*.*");
				szFile = szFile.Left(index);
				szFile += wfData.cFileName;
			}
			SetFileAttributes(szFile, FILE_ATTRIBUTE_ARCHIVE);
			if(!DeleteFile(szFile))
			{
				CString error;
				error += "cannot delete file: ";
				error += wfData.cFileName;
				AfxMessageBox(error);
				bRes = FALSE;
			}
		}
		
	} while(::FindNextFile(h, &wfData));
	FindClose(h);
	return bRes;

}

void _SearchDirFiles(CString strDir, Files& files)
{
	BOOL bRes = TRUE;
	WIN32_FIND_DATA wfData;
	HANDLE h = ::FindFirstFile(strDir, &wfData);
	if(h == (HANDLE)0xffffffff)
	{
		return;
	}
	do 
	{
		if(wfData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
		{
			continue;
		}
		files.push_back(wfData.cFileName);

	} while(::FindNextFile(h, &wfData));
	FindClose(h);
}

BOOL CALLBACK FunLog(char* pstr)
{
	ASSERT(pstr);
	CString strMsg;
	strMsg.Format("%s", pstr);
	
	return true;
}

BOOL CALLBACK FunProc(char* pstr)
{
	ASSERT(pstr);
	CString strMsg;
	strMsg.Format("%s", pstr);
	
	return true;
}

/*-----------------------------------------------------------------------------------
 范围内的超链接处理
------------------------------------------------------------------------------------*/
void COfficeWord::GenerateHyperlinks(CRange range)
{
	CHyperlinks hyperLinks(range.get_Hyperlinks());

	VARIANT vr;
	vr.vt = VT_I4;
	for(int i=1; i<=hyperLinks.get_Count(); i++)
	{
		vr.lVal = i;
		CHyperlink hyperLink(hyperLinks.Item(&vr));

		try
		{
			CString strAddress = hyperLink.get_Address();
			CString strName = hyperLink.get_Name(); //书签名
			CString strSubAddress = hyperLink.get_SubAddress();

			/* 地址为空代表引用内部超链接，不为空代表外部地址(保持原样不处理) */
			if(strAddress.IsEmpty())
			{
				CString strExternalAddress =
					ConvertInternalHyperlinkToExternalHyperlink(strName);
				hyperLink.put_Address(strExternalAddress);

			}
		}
		catch(...)
		{

		}

	}

}

/*-----------------------------------------------------------------------------------
 把引用文档内部标题的超链接转换为外部文件的超链接串

 word文档内部的超链接是word的交叉引用，类似html中的锚点，word内部超链接是通过书签(BookMark)
 
 来定位,在建立超链接的时候，word会自动为被引用的大纲标题建立隐藏书签(可在书签管理中选中包含“

 隐藏书签”查看word自动建立的书签）。

------------------------------------------------------------------------------------*/
CString COfficeWord::ConvertInternalHyperlinkToExternalHyperlink(CString strHyperlink)
{
	CString   strExternalHyperlink;

	//ItemArray	itemArr;
	//GenerateItemArray(_outlineTree->_firstChildItem, itemArr);

	//COutlineTreeItem* pItem;
	//CRange range;
	//for(int i=0; i<itemArr.size(); i++)
	//{
	//	pItem = itemArr[i];
	//	range = pItem->_paragraph.get_Range();
	//	CBookmarks bookmarks(range.get_Bookmarks());
	//	bookmarks.put_ShowHidden(TRUE);

	//	VARIANT vr;
	//	vr.vt = VT_I4;
	//	for(int j=1; j<=bookmarks.get_Count(); j++)
	//	{
	//		vr.lVal = j;
	//		CBookmark0 bookmark(bookmarks.Item(&vr));
	//		CString strName = bookmark.get_Name();

	//		if(strName.Compare(strHyperlink) == 0)
	//		{
	//			strExternalHyperlink.Format(_T("%d.html"), i);
	//			return strExternalHyperlink;
	//		}
	//	}

	//}

	CRange	range = GetBookmarkRange(strHyperlink);
	int	iHtmlPage = GetHtmlPageFromRange(range);
	strExternalHyperlink.Format(_T("%d.html"), iHtmlPage);
	return strExternalHyperlink;
}

CRange	COfficeWord::GetBookmarkRange(CString strBookmark)
{
	VARIANT vr;
	vr.vt = VT_I4;

	CBookmarks bookmarks(_wordDoc.get_Bookmarks());
	bookmarks.put_ShowHidden(TRUE);
	for(int i=1; i<=bookmarks.get_Count(); i++)
	{
		vr.lVal = i;
		CBookmark0 bookmark(bookmarks.Item(&vr));
		if(strBookmark.Compare(bookmark.get_Name()) == 0)
			return CRange(bookmark.get_Range());

	}

	//ASSERT(0);

	return CRange(0);
}

int COfficeWord::GetHtmlPageFromRange(CRange range)
{
	ItemArray	itemArr;
	GenerateItemArray(_outlineTree->_firstChildItem, itemArr);

	COutlineTreeItem* pItem;
	CRange rg;
	int i;
	for(i=0; i<itemArr.size(); i++)
	{
		pItem = itemArr[i];
		rg = pItem->_paragraph.get_Range();
		if(range.get_Start() < rg.get_Start())
			break;
 	}
	if(i == 0) return 0;
	return  i-1;
}

void  COfficeWord::RemoveUnderlineOfHyperlinks(CWordDocument doc)
{
	CStyles styles(doc.get_Styles());

	VARIANT vr;
	vr.vt = VT_I4;

	for(int i=1; i<=styles.get_Count(); i++)
	{
		vr.lVal = i;
		CStyle style(styles.Item(&vr));

		CString strName = style.get_NameLocal();
		strName.MakeLower();
		if(strName.Find("超链接") != -1 || strName.Find("hyperlink") != -1)
		{
			CFont0 font(style.get_Font());
			font.put_Underline(0);
		}
	}
}

/*
	产生相关联的主题，把子标题追加到父标题页的末尾
*/
void  COfficeWord::GenerateRelatedTopics(CWordDocument doc, COutlineTreeItem* pItem)
{
	ASSERT(pItem);

	if(!_bRelatedTopics) return;
	if(!pItem->_firstChildItem) return;

	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COutlineTreeItem* pSubItem = pItem->_firstChildItem;
	
	CParagraphs paragraphs(doc.get_Paragraphs());
	CHyperlinks hperlinks(doc.get_Hyperlinks());

	CParagraph spaceParagraph(paragraphs.Add(covOptional));
	//SetParagraphLineSpace(spaceParagraph, LineSpacingRule::wdLineSpaceSingle, 1);
	//SetParagraphSpaceAfterAndBefore(spaceParagraph, 3, 3);

	while(pSubItem)
	{
		CParagraph paragraph(paragraphs.Add(covOptional));
		CRange range(pSubItem->_paragraph.get_Range());
		CRange rg(paragraph.get_Range());
		CString strText = range.get_Text();
		strText.Remove(13);
		rg.put_Text(strText);
		
		//SetParagraphLineSpace(paragraph, LineSpacingRule::wdLineSpaceSingle, 1);
		//SetParagraphSpaceAfterAndBefore(paragraph, 3, 3);

		CListFormat listFormat = rg.get_ListFormat();
		listFormat.RemoveNumbers(covOptional);

		//CFont0 font(rg.get_Font());
		//font.put_Size(12);

		// hyperlink
		CString strAddress;
		strAddress.Format("%d.html", pSubItem->_pageIndex);
		COleVariant vAddress(strAddress);
		VARIANT vr = vAddress.Detach();
		hperlinks.Add(rg.m_lpDispatch, &vr, covOptional, covOptional, covOptional, covOptional);

		CParagraph nextPargraph(paragraphs.Add(covOptional));
		pSubItem = pSubItem->_nextItem;
	}
}

void COfficeWord::SetParagraphSpaceAfterAndBefore(CParagraph paragraph, float after, float before)
{
	paragraph.put_SpaceAfterAuto(0);
	paragraph.put_SpaceBeforeAuto(0);
	paragraph.put_SpaceAfter(after);
	paragraph.put_SpaceBefore(before);
}

/*
iLineSpacingRule:
	wdLineSpace1pt5			1		1.5 倍行距。该行距相当于当前字号加 6 磅。 
	wdLineSpaceAtLeast		3		行距至少为一个指定值。该值需要单独指定。 
	wdLineSpaceDouble		2		双倍行距。 
	wdLineSpaceExactly		4		行距只能是所需的最大行距。此设置所使用的行距通常小于单倍行距。 
	wdLineSpaceMultiple		5		由指定的行数确定的行距。 
	wdLineSpaceSingle		0		单倍行距，默认值。 
*/
void COfficeWord::SetParagraphLineSpace(CParagraph paragraph, int iLineSpacingRule, float space)
{
	//paragraph.put_LineSpacing(space);
	paragraph.put_LineSpacingRule(0);
}

void COfficeWord::SetDocumentSingleLineSpace(CWordDocument doc)
{
	CParagraphs paragraphs(doc.get_Paragraphs());
	
	for(int i=1; i<=paragraphs.get_Count(); i++)
	{
		CParagraph paragraph(paragraphs.Item(i));
		SetParagraphLineSpace(paragraph, 0, 0);
		SetParagraphSpaceAfterAndBefore(paragraph, 3, 3);
	}
}

void  COfficeWord::RemoveListNumber(CWordDocument doc)
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	CParagraphs listParagraphs(doc.get_ListParagraphs());
	for(int i=1; i<=listParagraphs.get_Count(); i++)
	{
		CParagraph paragraph(listParagraphs.Item(i));
		if(paragraph.get_OutlineLevel() >= 10)
			continue;
		CRange range(paragraph.get_Range());
		CListFormat listFormat = range.get_ListFormat();
		listFormat.RemoveNumbers(covOptional);
	}
}

void  COfficeWord::SetRegistered(BOOL bRegistered)
{
	_bRegistered = bRegistered;
}

void  COfficeWord::GenerateUnRegisteredFootnotes(CWordDocument doc)
{
	if(_bRegistered)
		return;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	CParagraphs paragraphs(doc.get_Paragraphs());
	paragraphs.Add(covOptional);
	paragraphs.Add(covOptional);
	CParagraph paragraph(paragraphs.Add(covOptional));
	CRange rg(paragraph.get_Range());
	rg.put_Text(UNREGISTER_FOOTNOTES);	
	CFont0 font(rg.get_Font());
	font.put_Size(10);
	font.put_Color(RGB(0,0,255));

}

void  COfficeWord::GenerateCopyright(CWordDocument doc)
{
	if(CChmConfig::GetInstance()->m_strCopyright.IsEmpty())
		return;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	CParagraphs paragraphs(doc.get_Paragraphs());
	paragraphs.Add(covOptional);
	paragraphs.Add(covOptional);
	CParagraph paragraph(paragraphs.Add(covOptional));
	CRange rg(paragraph.get_Range());
	CnlineShapes inlineShapes(rg.get_InlineShapes());
	CnlineShape inlineShape(inlineShapes.AddHorizontalLineStandard(covOptional));
	CHorizontalLineFormat horzLineFormat(inlineShape.get_HorizontalLineFormat());
	horzLineFormat.put_PercentWidth(100.0);

	CParagraph para(paragraphs.Add(covOptional));
	CRange range(para.get_Range());
	range.put_Text(CChmConfig::GetInstance()->m_strCopyright);	
	CFont0 font(rg.get_Font());
	font.put_Size(10);
	font.put_Color(RGB(168,168,168));
}
