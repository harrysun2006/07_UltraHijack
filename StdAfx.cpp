// stdafx.cpp : source file that includes just the standard includes
//	IESpy.pch will be the pre-compiled header
//	stdafx.obj will contain the pre-compiled type information

#include "stdafx.h"

#if WINVER < 0x0500
HWND GetAncestor(HWND hwnd, UINT gaFlags)
{
	if (!IsWindow(hwnd)) return NULL;
	HWND hParentWnd = NULL;
	while(hwnd) {
		hParentWnd = hwnd;
		hwnd = GetParent(hwnd);
	}
	return hParentWnd;
}
#endif

// TODO: reference any additional headers you need in STDAFX.H
// and not in this file
