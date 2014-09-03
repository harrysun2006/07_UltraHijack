#if !defined(IESPY__INCLUDED_)
#define IESPY__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"
#define WM_SHELL WM_USER + 1

void HijackLogin(IDispatch *pDisp);

#endif // !defined(IESPY__INCLUDED_)