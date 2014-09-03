// UniHelper.cpp : unique-instance helper class
//

#include "stdafx.h"
#include "UniHelper.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

using namespace std;

namespace unihelper
{
WNDCLASS	uniWndClass = {0, &DefWindowProc, 0, 0, 0, 0, 0, 0, _T("CUniqueWindow")};
DWORD		dwBSMRecipients = BSM_APPLICATIONS, dwMessageId;
HWND		hUniqueWnd = NULL;
HANDLE		hMutex = NULL;
WNDPROC		procCallback;

HWND AllocateHWND(HINSTANCE hInstance, WNDPROC proc)
{
	HWND		hWnd;
	WNDCLASS	tWndClass;
	BOOL		isRegistered;
	uniWndClass.hInstance = hInstance;
	uniWndClass.lpfnWndProc = &DefWindowProc;
	isRegistered = GetClassInfo(hInstance, uniWndClass.lpszClassName, &tWndClass);
	if (!isRegistered || tWndClass.lpfnWndProc != &DefWindowProc)
	{
		if (isRegistered) UnregisterClass(uniWndClass.lpszClassName, hInstance);
		RegisterClass(&uniWndClass);
	}
	hWnd = CreateWindowEx(WS_EX_TOOLWINDOW, uniWndClass.lpszClassName, _T(""), WS_POPUP, 0, 0, 0, 0, 0, 0, hInstance, NULL);
	if (proc) SetWindowLong(hWnd, GWL_WNDPROC, (long)&proc);
	return hWnd;
}

void DeallocateHWnd(HWND hWnd)
{
	DestroyWindow(hWnd);
}

BOOL WINAPI UniqueInstanceWndProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	if (Msg == dwMessageId)
	{
		procCallback(hWnd, Msg, wParam, lParam);
		return TRUE;
	}
	else return DefWindowProc(hWnd, Msg, wParam, lParam);
}

BOOL OpenUniqueInstance(HINSTANCE hInstance, LPCTSTR id, WNDPROC proc)
{
	BOOL	bRet;
	dwMessageId = RegisterWindowMessage(id);
	procCallback = proc;
	hUniqueWnd = AllocateHWND(hInstance, (WNDPROC)UniqueInstanceWndProc);
	hMutex = OpenMutex(MUTEX_ALL_ACCESS, FALSE, id);
	if (hMutex == NULL)
	{
		hMutex = CreateMutex(NULL, FALSE, id);
		if (hMutex == NULL || GetLastError() == ERROR_ALREADY_EXISTS) bRet = FALSE;
		else bRet = TRUE;
	} else {
		bRet = FALSE;
	}
	CString str;
	str.Format("hMutex = %d, bRet = %d", hMutex, bRet);
	MessageBox(0, str, "Debug", MB_ICONINFORMATION);

	CloseHandle(hMutex);
	if (!bRet) BroadcastSystemMessage(BSF_IGNORECURRENTTASK || BSF_POSTMESSAGE, &dwBSMRecipients, dwMessageId, 0, 0);
	return bRet;
}

void CloseUniqueInstance()
{
	DeallocateHWnd(hUniqueWnd);
}

}