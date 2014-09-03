// IEHelper.h: define ie helper class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(IEHELPER__INCLUDED_)
#define IEHELPER__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"

namespace iehelper
{
CComPtr<IWebBrowser2> GetWebBrowser2(IHTMLDocument2*);
CComPtr<IHTMLDocument2> GetHTMLDocument2(IWebBrowser2*);
CComPtr<IHTMLWindow2> GetHTMLWindow2(IHTMLDocument2*);
CComPtr<IHTMLWindow2> GetHTMLWindow2(IWebBrowser2*);
CComPtr<IHTMLFormElement> GetHTMLForm(IHTMLDocument2*, LPCTSTR);
CComPtr<IHTMLElement> GetFormInput(IHTMLFormElement*, LPCTSTR);
HRESULT ExecuteScript(IHTMLWindow2*, LPCTSTR, LPCTSTR = "JavaScript");
}

#endif // !defined(IEHELPER__INCLUDED_)