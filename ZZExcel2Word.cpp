#include "StdAfx.h"
#include "ZZExcel2Word.h"
#include "..\CExcelApplication.h"
#include "..\CExcelWorkbook.h"
#include "..\CExcelWorkbooks.h"
#include "..\CExcelWorksheet.h"
#include "..\CExcelWorksheets.h"
#include "..\CExcelRange.h"
#include <sstream>
CZZExcel2Word::CZZExcel2Word()
{
	m_stringWordDocKey =  _T("管道序号");
	m_iValueNameRow = 2;
	InitExportSettings();
	m_ioutputOnlyoneFile = 0;
}
#define INI_DEBUG_APPNAME _T("调试规则")
#define INI_DEBUG_KEYNAME_ONLYOUTPUTONEFILE _T("仅输出一个文件")

#define INI_EXCEL_APPNAME _T("读取Excel规则")
#define INI_EXCEL_KEYNAME_VALUENAMEROWINDEX _T("表头行")

#define INI_REPORT_APPNAME _T("报告属性")
#define INI_REPORT_KEYNAME _T("报告关键字")
#define INI_DATAMAP_APPNAME _T("数据匹配")
#define MAX_LENGTH 1024
#define INI_FILE_NAME _T("ZZ.ini")
#define Template_FILE_NAME _T("ZZTemplate.dot")
std::wstring CZZExcel2Word::GetCurrentDir()
{
	TCHAR szBuf[MAX_PATH];
	ZeroMemory(szBuf,MAX_PATH);
	std::wstring initFilePath;
	if (GetCurrentDirectory(MAX_PATH,szBuf) > 0)	
	{
		initFilePath = szBuf;
		initFilePath += _T("\\");
	}
	else
	{
		initFilePath = _T("C:\\");
	}
	return initFilePath;
}
HRESULT CZZExcel2Word::InitExportSettings()
{
	std::wstring curdir = GetCurrentDir();
	std::wstring initFilePath;
	m_stringWordTemplatePath = curdir + Template_FILE_NAME;
	initFilePath = curdir+INI_FILE_NAME;
	return InitExportSettings(initFilePath.c_str());
}
//遍历ini文件
HRESULT CZZExcel2Word::InitExportSettings(LPCTSTR pFilePath)  
{  
	if (!PathFileExists(pFilePath))
	{
		WritePrivateProfileString(INI_EXCEL_APPNAME,INI_EXCEL_KEYNAME_VALUENAMEROWINDEX,_T("2"),pFilePath);
		WritePrivateProfileString(INI_REPORT_APPNAME,INI_REPORT_KEYNAME,_T("管道序号"),pFilePath);
		WritePrivateProfileString(INI_DATAMAP_APPNAME,_T("管道名称"),_T("管道名称"),pFilePath);
		WritePrivateProfileString(INI_DEBUG_APPNAME,INI_DEBUG_KEYNAME_ONLYOUTPUTONEFILE,_T("1"),pFilePath);
	}
	// TODO: Add your control notification handler code here  
	TCHAR strAppNameTemp[1024];//所有AppName的返回值  
	TCHAR strKeyNameTemp[1024];//对应每个AppName的所有KeyName的返回值  
	TCHAR strReturnTemp[1024];//返回值  
	DWORD dwKeyNameSize;//对应每个AppName的所有KeyName的总长度  
	//所有AppName的总长度  
	DWORD dwAppNameSize = GetPrivateProfileString(NULL,NULL,NULL,strAppNameTemp,MAX_LENGTH,pFilePath);  
	if(dwAppNameSize>0)  
	{  
		TCHAR *pAppName = new TCHAR[dwAppNameSize];  
		int nAppNameLen=0;  //每个AppName的长度  
		for(DWORD i = 0;i<dwAppNameSize;i++)  
		{  
			pAppName[nAppNameLen++]=strAppNameTemp[i];  
			if(strAppNameTemp[i]==0)  
			{  
				dwKeyNameSize = GetPrivateProfileString(pAppName,NULL,NULL,strKeyNameTemp,1024,pFilePath);  
				if(dwAppNameSize>0)  
				{  
					TCHAR *pKeyName = new TCHAR[dwKeyNameSize];  
					int nKeyNameLen=0;    //每个KeyName的长度  
					for(DWORD j = 0;j<dwKeyNameSize;j++)  
					{  
						pKeyName[nKeyNameLen++]=strKeyNameTemp[j];  
						if(strKeyNameTemp[j]==0)  
						{  
							if(GetPrivateProfileString(pAppName,pKeyName,NULL,strReturnTemp,1024,pFilePath))  
							{
								NULL;
							}
							//my code here. szg
							std::wstring tempappstr = pAppName;
							if (tempappstr == INI_REPORT_APPNAME)
							{
								m_stringWordDocKey = strReturnTemp;
							}
							if (tempappstr == INI_DATAMAP_APPNAME)
							{
								std::wstring tempbookmarkstr = pKeyName;
								std::wstring tempReturnstr = strReturnTemp;
								AddBookMarkDataPair(tempReturnstr,tempbookmarkstr);
							}
							if (tempappstr == INI_EXCEL_APPNAME)
							{
								std::wstring tempkeyname = pKeyName;
								if (tempkeyname == INI_EXCEL_KEYNAME_VALUENAMEROWINDEX)
								{
									m_iValueNameRow = _ttoi(strReturnTemp);
								} 
							}
							if (tempappstr == INI_DEBUG_APPNAME)
							{
								std::wstring tempkeyname = pKeyName;
								if (tempkeyname == INI_DEBUG_KEYNAME_ONLYOUTPUTONEFILE)
								{
									m_ioutputOnlyoneFile = _ttoi(strReturnTemp);
								} 
							}
							memset(pKeyName,0,dwKeyNameSize);  
							nKeyNameLen=0;  
						}  
					}  
					delete[]pKeyName;  
				}  
				memset(pAppName,0,dwAppNameSize);  
				nAppNameLen=0;  
			}  
		}  
		delete[]pAppName;  
	}  
	return S_OK;
}  
CZZExcel2Word::~CZZExcel2Word(void)
{
	ClearWordDoc();
}


HRESULT CZZExcel2Word::ExportDataToWordDoc(std::wstring LocationFolder)
{
	std::vector<PZZWordDoc>::iterator it;

	for (it = m_vecWordDoc.begin(); it != m_vecWordDoc.end(); ++it)
	{
		PZZWordDoc temp = *it;
		temp->GenerateWordDoc(LocationFolder);
	}
	return S_OK;
}
HRESULT CZZExcel2Word::AddBookMarkDataPair(std::wstring DataName,std::wstring BookMarkName)
{
	m_mapDataItem2BookMark.insert(std::pair<std::wstring, std::wstring>(BookMarkName, DataName));
	return S_OK;
}
HRESULT CZZExcel2Word::TransferExcelFiles2Word(std::vector<std::wstring> vecExcelFiles)
{
	std::vector<std::wstring>::iterator it;

	for (it = vecExcelFiles.begin(); it != vecExcelFiles.end(); ++it)
	{
		std::wstring temp = *it;
		BuildDataFromExcelFile(temp,m_stringWordDocKey);
	}
	return S_OK;
}
static std::wstring GetStringFromExcelCell(CExcelRange& useRange)
{
	COleVariant keyValue = useRange.get_Value2();	
	std::wstring itemString;
	if (keyValue.vt != VT_BSTR)
	{
		if (keyValue.vt == VT_R8)
		{
			std::wostringstream ostr;
			ostr<<keyValue.dblVal;
			itemString = ostr.str();
		}
	}
	else
	{
		if (keyValue.bstrVal != NULL)
		{
			itemString = keyValue.bstrVal;
		}	
	}
	return itemString;
}
HRESULT CZZExcel2Word::BuildDataFromExcelFile(std::wstring ExcelFile,std::wstring stringDocKey)
{
	CExcelApplication ExcelApp; 
	CExcelWorkbooks wbsMyBooks;  
	CExcelWorkbook wbMyBook;  
	CExcelWorksheets wssMysheets; 
	CExcelWorksheet wsMysheet;  
	//创建Excel 2000服务器(启动Excel) 

//	IID clsid;
//	HRESULT hr = IIDFromString(_T("Excel.Application"), &clsid);

	if (!ExcelApp.CreateDispatch(_T("Excel.Application"),NULL))  
	{   
		AfxMessageBox(_T("创建Excel服务失败!"));  
		exit(1);   
	}  

	//利用模板文件建立新文档  
	wbsMyBooks = ExcelApp.get_Workbooks(); 
	 COleVariant  avar((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
	wbMyBook = wbsMyBooks.Open(ExcelFile.c_str(),avar,avar,avar,avar,avar,avar,avar,avar,avar,avar,avar,avar,avar,avar); 
	//得到Worksheets   
	wssMysheets = wbMyBook.get_Sheets(); 
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,   VT_ERROR);
	CExcelRange useRange;
	for (int iSheetIdx = 1; iSheetIdx <= wssMysheets.get_Count() ; iSheetIdx++)
	{
		//得到sheet 
		wsMysheet=wssMysheets.get_Item(COleVariant((short)iSheetIdx)); 
		
#ifdef _DEBUG
		CString sheetname = wsMysheet.get_Name();
#endif // _DEBUG
		//首先分析表的第一行，得到数据关键字
		int iKeyColumnIdx = -1;
		CExcelRange useRange = wsMysheet.get_UsedRange();

		long iRowNum = useRange.get_Count();
		useRange = useRange.get_Columns();

		long iStartRow = useRange.get_Row();
		int nColumn = useRange.get_Count();
		long iStartCol = useRange.get_Column();
		long iValueNameRow = m_iValueNameRow;
		//find keyName Column index
		for (int iColIdx = iStartCol; iColIdx <= nColumn ; iColIdx++)
		{
			useRange = wsMysheet.get_Cells();
			COleVariant keyValue=useRange.get_Item(_variant_t(iValueNameRow),_variant_t(iColIdx));
			useRange = keyValue.pdispVal;
			std::wstring itemString = GetStringFromExcelCell(useRange);
			if (stringDocKey == itemString )
			{
				iKeyColumnIdx = iColIdx;
				break;
			}
		}
		if (iKeyColumnIdx < 0)
		{
			continue;//no valid data with doc key string
		}
		//根据关键字列创建doc

		for (int iRodIdx = 3; iRodIdx <= iRowNum ; iRodIdx++)
		{
			useRange = wsMysheet.get_Cells();
			COleVariant keyValue=useRange.get_Item(_variant_t(iRodIdx),_variant_t(iKeyColumnIdx));
			useRange = keyValue.pdispVal;
			std::wstring keyitemString = GetStringFromExcelCell(useRange);
			if (keyitemString.empty())
			{
				continue;
			}
			PZZWordDoc pDoc = GetDocFromKeyString(keyitemString);
			if (pDoc == NULL)
			{
				pDoc = CreateDoc(keyitemString);
				if (pDoc == NULL)
				{
					continue;
				}
				SetExportReportSetting(pDoc);
				m_vecWordDoc.push_back(pDoc);
			}

			//逐个读取相关doc的数据
			for (int iColIdx = iStartCol; iColIdx <= nColumn ; iColIdx++)
			{
				if (iColIdx == iKeyColumnIdx)
				{
					continue;
				}
				useRange = wsMysheet.get_Cells();
				keyValue=useRange.get_Item(_variant_t(iRodIdx),_variant_t(iColIdx));
				useRange = keyValue.pdispVal;
				std::wstring valueitemString  = GetStringFromExcelCell(useRange);
				if (valueitemString.empty())
				{
					//无效数据
					continue;
				}
				//查找这个值对应的名字
				useRange = wsMysheet.get_Cells();
				keyValue=useRange.get_Item(_variant_t(iValueNameRow),_variant_t(iColIdx));
				useRange = keyValue.pdispVal;	
				std::wstring valuenameitemString = GetStringFromExcelCell(useRange);

				if (valuenameitemString.empty())
				{
					//无效数据名字
					continue;
				}
				pDoc->AddDataItem(valuenameitemString,valueitemString);
			}
		}
		if (m_ioutputOnlyoneFile ==1)
		{
			break;
		}
		wsMysheet.ReleaseDispatch();  
	}//sheet scan
	wssMysheets.ReleaseDispatch();  
	wbsMyBooks.Close();
	wbMyBook.ReleaseDispatch();  
	wbsMyBooks.ReleaseDispatch();  

	ExcelApp.Quit();

	ExcelApp.ReleaseDispatch();
	return S_OK;
}


PZZWordDoc CZZExcel2Word::CreateDoc(std::wstring key)
{
	PZZWordDoc newDoc = new CZZWordDoc();
	newDoc->SetName(key);
	return newDoc;
}

PZZWordDoc CZZExcel2Word::GetDocFromKeyString(std::wstring key)
{
	std::vector<PZZWordDoc>::iterator it;

	for (it = m_vecWordDoc.begin(); it != m_vecWordDoc.end(); ++it)
	{
		PZZWordDoc temp = *it;
		if (temp->GetName() == key)
		{
			return temp;
		}
	}
	return NULL;
}

void CZZExcel2Word::ClearWordDoc()
{
	std::vector<PZZWordDoc>::iterator it;

	for (it = m_vecWordDoc.begin(); it != m_vecWordDoc.end(); ++it)
	{
		PZZWordDoc temp = *it;
		delete temp;
	}
	m_vecWordDoc.clear();
}

void CZZExcel2Word::SetExportReportSetting( PZZWordDoc pDoc )
{
	pDoc->SetMapDataItem2BookMark(m_mapDataItem2BookMark);
	pDoc->SetStringWordTemplatePath(m_stringWordTemplatePath);
}
