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
char* GetIIDName(IID iid);
IWebBrowser2* GetWebBrowser(IHTMLDocument2 *pIHTMLDocument2);
HWND GetWindowHandle(IWebBrowser2 *pWebBrowser2);
HWND GetWindowHandle(IHTMLDocument2 *pIHTMLDocument2);
IHTMLFormElement* GetForm(IHTMLDocument2 *pIHTMLDocument2, LPCSTR lpsz);
IHTMLElement* GetField(IHTMLFormElement *pIHTMLFormElement, LPCSTR lpsz);
void EnumDocument(IHTMLDocument2 *pIHTMLDocument2, HWND hWnd);
}

#endif // !defined(IEHELPER__INCLUDED_)