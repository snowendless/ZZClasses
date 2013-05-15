#pragma once
#include "zzdataitem.h"
class CZZStringDataItem :
	public CZZDataItem
{
	std::wstring m_stringValue;
public:
	std::wstring GetStringValue() const { return m_stringValue; }
	void SetStringValue(std::wstring val) { m_stringValue = val; }
	virtual std::wstring GetValueString(){ return m_stringValue; }
	CZZStringDataItem(void);
	~CZZStringDataItem(void);
};
typedef CZZStringDataItem* PZZStringDataItem;
