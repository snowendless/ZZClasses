#pragma once
#include <string>
#include <vector>

#include "ZZWordDoc.h"
class CZZExcel2Word
{
	std::map<std::wstring,std::wstring> m_mapDataItem2BookMark;

	std::wstring m_stringWordTemplatePath;
	std::wstring m_stringWordDocKey;
	int m_iValueNameRow;
	int m_ioutputOnlyoneFile;
	std::vector<PZZWordDoc> m_vecWordDoc;
	PZZWordDoc GetDocFromKeyString(std::wstring key);
	PZZWordDoc CreateDoc(std::wstring key);
	void ClearWordDoc();
public:
	static std::wstring GetCurrentDir();
	void SetStringWordTemplatePath(std::wstring val) { m_stringWordTemplatePath = val; }
	HRESULT InitExportSettings();
	HRESULT AddBookMarkDataPair(std::wstring BookMarkName,std::wstring DataName);
	HRESULT ExportDataToWordDoc(std::wstring LocationFolder);
	HRESULT TransferExcelFiles2Word(std::vector<std::wstring> vecExcelFiles);
	HRESULT BuildDataFromExcelFile(std::wstring ExcelFile,std::wstring stringDocKey);

	void SetExportReportSetting( PZZWordDoc pDoc );

	HRESULT InitExportSettings(LPCTSTR pFilePath);
	CZZExcel2Word();
	~CZZExcel2Word(void);


};

