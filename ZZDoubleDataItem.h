#pragma once
#include "zzdataitem.h"
class CZZDoubleDataItem :
	public CZZDataItem
{
	double	m_dValue;
public:
	double GetValue() const { return m_dValue; }
	void SetValue(double val) { m_dValue = val; }
	virtual std::wstring GetValueString();
	CZZDoubleDataItem(void);
	~CZZDoubleDataItem(void);
};

