// SzgswHijackSink.h: interface for the CSzgswHijackSink class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(SZGSWHIJACKSINK__INCLUDED_)
#define SZGSWHIJACKSINK__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"
#include "IESink.h"

class CSzgswHijackSink : public CIESink
{
protected:
	virtual void OnDocumentComplete(IDispatch *pDisp, VARIANT *pvarURL);
};

void HijackLogin(IDispatch *pDisp);

#endif // !defined(SZGSWHIJACKSINK__INCLUDED_)