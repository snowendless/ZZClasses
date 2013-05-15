#pragma once
#include <string>
class CZZDataItem
{
	std::wstring m_stringName;
public:
	std::wstring GetName() const { return m_stringName; }
	void SetName(std::wstring val) { m_stringName = val; }
	virtual std::wstring GetValueString() = 0;
	CZZDataItem(void);
	~CZZDataItem(void);
};
typedef CZZDataItem* PZZDataItem;

