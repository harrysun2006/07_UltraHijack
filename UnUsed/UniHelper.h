// UniHelper.h: define unique-instance helper class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(UNIHELPER__INCLUDED_)
#define UNIHELPER__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

namespace unihelper
{
BOOL OpenUniqueInstance(HINSTANCE, LPCTSTR, WNDPROC);
void CloseUniqueInstance();
}

#endif // !defined(UNIHELPER__INCLUDED_)