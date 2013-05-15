#pragma once
#include <string>
#include <vector>
#include <map>
#include "ZZDataItem.h"
class CZZWordDoc
{
	std::map<std::wstring,std::wstring> m_mapDataItem2BookMark;
	std::wstring m_stringWordTemplatePath;
	std::wstring m_stringName;
	std::vector<PZZDataItem> m_vecDataItems;
	HRESULT GetBookMarkFromDataname(std::wstring dataName,std::vector<std::wstring>& vecBookMarkStrings);
public:
	void SetMapDataItem2BookMark(std::map<std::wstring,std::wstring> val) { m_mapDataItem2BookMark = val; }
	HRESULT AddBookMarkDataPair(std::wstring BookMarkName,std::wstring DataName);
	std::wstring GetStringWordTemplatePath() const { return m_stringWordTemplatePath; }
	void SetStringWordTemplatePath(std::wstring val) { m_stringWordTemplatePath = val; }
	HRESULT GenerateWordDoc(std::wstring LocationFolder);
	HRESULT AddDataItem(std::wstring DataName,std::wstring dataString);
	std::wstring GetName() const { return m_stringName; }
	void SetName(std::wstring val) { m_stringName = val; }
	CZZWordDoc(void);
	~CZZWordDoc(void);

	void ClearDataItem();

};
typedef CZZWordDoc* PZZWordDoc;
