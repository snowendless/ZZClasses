#include "StdAfx.h"
#include "ZZDoubleDataItem.h"
#include <sstream>

CZZDoubleDataItem::CZZDoubleDataItem(void)
{
	m_dValue = 0;
}


CZZDoubleDataItem::~CZZDoubleDataItem(void)
{
}
 std::wstring CZZDoubleDataItem::GetValueString()
 {
	 std::wostringstream   ostr;

	 return ostr.str();
 }