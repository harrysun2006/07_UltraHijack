// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__A9DB83DB_A9FD_11D0_BFD1_444553540000__INCLUDED_)
#define AFX_STDAFX_H__A9DB83DB_A9FD_11D0_BFD1_444553540000__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers
#define STRICT

/*
#ifndef WINVER				// 允许使用 Windows 95 和 Windows NT 4 或更高版本的特定功能。
#define WINVER 0x0500       // 为 Windows98 和 Windows 2000 及更新版本改变为适当的值。
#endif

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0500
#endif                           

#ifndef _WIN32_WINDOWS
#define _WIN32_WINDOWS 0x0500
#endif

#ifndef _WIN32_IE           // 允许使用 IE 4.0 或更高版本的特定功能。
#define _WIN32_IE 0x0500    // 为 IE 5.0 及更新版本改变为适当的值。
#endif
*/

#define _ATL_APARTMENT_THREADED

#include <afxdtctl.h>		// MFC support for Internet Explorer 4 Common Controls
#include <atlbase.h>
extern CComModule _Module;
#include <atlcom.h>
#include <atlwin.h>
#include <atlhost.h>

#include <exdisp.h>
#include <windows.h>
#include <iostream>
#include <fstream>

#if WINVER < 0x0500
#define     GA_MIC          1
#define     GA_PARENT       1
#define     GA_ROOT         2
#define     GA_ROOTOWNER    3
#define     GA_MAC          4
HWND GetAncestor(HWND hwnd, UINT gaFlags);
#endif

// TODO: reference additional headers your program requires here

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__A9DB83DB_A9FD_11D0_BFD1_444553540000__INCLUDED_)
