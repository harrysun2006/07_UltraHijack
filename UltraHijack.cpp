// IESpy.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "resource.h"
#include "UltraHijack.h"
#include "SystemTrayClass.h"
#include "Spyer.h"
#include "SzgswHijackSink.h"
#include "IEHelper.h"

#define	ID_UNIQUE_APP	_T("UltraHijack")

#pragma comment (lib, "ShellHook.lib")
extern BOOL WINAPI InstallShellHookEx(HWND hWnd, UINT uMsg);
extern BOOL WINAPI UnInstallShellHookEx(HWND hWnd);

CComModule _Module;

using namespace std;
using namespace spyer;
using namespace iehelper;

#define NORMAL			1
#define BOLD			2
#define ITALIC			3
#define ULINE			4
#define	VK_ALT			0x12

typedef struct _MENUITEM{ 
	HFONT hfont; 
	char  psz[1]; 
} MENUITEM;             // structure for item font and string  
typedef struct _MENUITEM * LPMENUITEM;

int AddToolTip(HWND, LPSTR, HWND, HWND, UINT);
BOOL WINAPI OnExecuteMenuItem(DWORD);
VOID WINAPI OnTrayNotify(HWND hWnd,LPARAM lParam);
BOOL WINAPI OnInitDialogNoop(HWND hWnd);
VOID WINAPI CenterWindow(HWND hWnd);
VOID WINAPI OnExitApp();
VOID WINAPI OnMeasureItem(HWND hWnd, LPMEASUREITEMSTRUCT lpmis);
VOID WINAPI OnDrawItem(HWND hWnd, LPDRAWITEMSTRUCT lpdis);

HINSTANCE		hAppInstance;
HWND			hNoopWnd = NULL, hAboutWnd = NULL, ToolTip_Handle = NULL;
HMENU			hMainMenu, hPopMenu, hSysMenu;
clsSysTray		SystemTrayEx;

DWORD			dwBSMRecipients = BSM_APPLICATIONS, dwMessageId;
HANDLE			hMutex = NULL;
WNDCLASS		uniWndClass = {0, DefWindowProc, 0, 0, 0, 0, 0, 0, NULL, _T("CUniqueWindow")};

void Error(LPSTR lpszFunction) 
{ 
	DWORD dw = GetLastError(); 

	CString str;
	str.Format("%s failed: GetLastError returned %u!\n", lpszFunction, dw); 

	MessageBox(NULL, str, "Error", MB_ICONERROR); 
	//ExitProcess(dw); 
} 

void Debug(LPCTSTR psz)
{
	MessageBox(NULL, psz, "Debug", MB_ICONINFORMATION); 
}

void SinkIE(HWND hWnd)
{
	CSzgswHijackSink* pSink = new CSzgswHijackSink();
	HRESULT hr = CSpyer::Instance()->AddIESink(hWnd, pSink);
}

LRESULT OnShell(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (wParam)
	{
	case HSHELL_WINDOWCREATED:
		SinkIE(HWND(lParam));
		break;
	case HSHELL_WINDOWDESTROYED:
		break;
	case HSHELL_WINDOWACTIVATED:
		break;
	default:
		break;
	}
	return S_OK;
}

LRESULT DoScan()
{
	CSzgswHijackSink* pSink = new CSzgswHijackSink();
	CSpyer::Instance()->AddIESink(HWND(NULL), pSink);
	return S_OK;
}

LRESULT DoHook()
{
	InstallShellHookEx(hNoopWnd, WM_SHELL);
	return S_OK;
}

LRESULT DoUnhook()
{
	UnInstallShellHookEx(hNoopWnd);
	return S_OK;
}

BOOL CALLBACK AboutDialogProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam)
{ 
	PAINTSTRUCT				ps;	
	HDC						hDC;
	switch (Msg)
	{ 
		case WM_INITDIALOG:
			break; 
		case WM_COMMAND:
			break;
		case WM_PAINT: 
			hDC = BeginPaint(hWnd, &ps);	 
			EndPaint(hWnd, &ps);
   			break;
		case WM_SYSCOMMAND:
			if(LOWORD(wParam) == SC_MINIMIZE || LOWORD(wParam) ==  SC_CLOSE)
			{
				ShowWindow(hWnd, SW_HIDE);
				return(0);
			}
			break;
		case WM_LBUTTONDOWN:
			break;
		case WM_DESTROY: 
			ShowWindow(hWnd, SW_HIDE);
			break;
		default: break;
	} 
	return FALSE;//DefWindowProc(hWnd, Msg, wParam, lParam);
} 

BOOL CALLBACK NoopDialogProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam)
{ 
	PAINTSTRUCT				ps;	
	HDC						hDC;
	// if a second instance is running, then the first will get a broadcast message here
	if (Msg == dwMessageId) return OnExecuteMenuItem(ID_ABOUT);
	switch (Msg)
	{ 
		case WM_TRAYNOTIFY:
			OnTrayNotify(hWnd, lParam);
			break;
		case WM_INITDIALOG: 
			OnInitDialogNoop(hWnd);
			return FALSE;
			break;
		case WM_COMMAND:
			OnExecuteMenuItem(LOWORD(wParam));
			break;
		case WM_MEASUREITEM: 
			OnMeasureItem(hWnd, (LPMEASUREITEMSTRUCT) lParam); 
			return TRUE;
			break; 
		case WM_DRAWITEM: 
			OnDrawItem(hWnd, (LPDRAWITEMSTRUCT) lParam); 
			return TRUE; 
			break;
		case WM_NEXTDLGCTL:
			SetFocus(GetDlgItem(hWnd, wParam));
			break;
		case WM_MENUSELECT:
			/* Currently Unused */
			break;
		case WM_SYSCOMMAND:
			if(LOWORD(wParam) == SC_MINIMIZE || LOWORD(wParam) ==  SC_CLOSE)
			{
				ShowWindow(hWnd, SW_HIDE);
				break;
			}
			OnExecuteMenuItem(LOWORD(wParam));
			break;
		case WM_SIZE:
			if(wParam == SIZE_MINIMIZED)
				ShowWindow(hWnd, SW_HIDE);
			break;
		case WM_PAINT: 
			hDC = BeginPaint(hWnd, &ps);	 
			EndPaint(hWnd, &ps);
   			break;
		case WM_SHOWWINDOW:
			//SendMessage(hOption_Win,WM_KEYDOWN,VK_ALT,0x21000041);
			//SendMessage(hOption_Win,WM_KEYUP,VK_ALT,0x21000041);
			//SetActiveWindow(hWnd);
			break;
		case WM_DESTROY: 
			ShowWindow(hWnd, SW_HIDE);
			break;
		case WM_SHELL:
			OnShell(Msg, wParam, lParam);
			break;
		default:break;
	} 
	return FALSE;//DefWindowProc(hWnd, Msg, wParam, lParam);
} 

void ForceForegroundWindow(HWND hWnd)
{
	HWND	hCurWnd;
	ULONG	curThreadId, hThreadId;
	hCurWnd = GetForegroundWindow();
	if (hWnd != hCurWnd)
	{
		GetWindowThreadProcessId(hCurWnd, &curThreadId);
		hThreadId = GetCurrentThreadId();
		if (curThreadId != hThreadId)
		{
			AttachThreadInput(curThreadId, hThreadId, TRUE);
			SetForegroundWindow(hWnd);
			AttachThreadInput(curThreadId, hThreadId, FALSE);
		} else {
			SetForegroundWindow(hWnd);
		}
		if (IsIconic(hWnd)) ShowWindow(hWnd, SW_RESTORE);
		else ShowWindow(hWnd, SW_SHOW);
	}
}

BOOL WINAPI OnExecuteMenuItem(DWORD dwItem)
{
	switch(dwItem)
	{
		case ID_EXIT_APP:
			OnExitApp();
			break;
		case ID_ABOUT:
			// ShowWindow(hAboutWnd, SW_SHOWNORMAL);
			// SetWindowPos(hAboutWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
			ForceForegroundWindow(hAboutWnd);
			CenterWindow(hAboutWnd);
			break;
		default: return FALSE;
	}
	return TRUE;
}

int AddToolTip(HWND hTTControl, LPSTR lpszString, HWND hWnd, HWND hParent, UINT tuFlags)
{
	TOOLINFO tInfo;
	int ret;
	tInfo.cbSize = sizeof(TOOLINFO);
	tInfo.hwnd = hParent;
	tInfo.uFlags = tuFlags | TTF_IDISHWND ; 
	tInfo.hinst = hAppInstance;
	tInfo.lpszText = lpszString;
	tInfo.uId = (UINT)hWnd;//GetDlgCtrlID(hWnd);

	ret = (int)SendMessage(hTTControl, TTM_ADDTOOL, NULL, (LPARAM)(LPTOOLINFO) &tInfo);
	//if(ret==TRUE)MessageBox(NULL,"Add ToolTip Successfully!","Info",MB_OK);
	return ret;
}

HWND AllocateHWND(HINSTANCE hInstance, WNDPROC proc)
{
	HWND		hWnd;
	WNDCLASS	tWndClass;
	HRESULT		hr;
	uniWndClass.hInstance = hInstance;
	uniWndClass.lpfnWndProc = DefWindowProc;
	hr = GetClassInfo(hInstance, uniWndClass.lpszClassName, &tWndClass);
	if (hr == 0 || tWndClass.lpfnWndProc != DefWindowProc)
	{
		if (hr != 0) UnregisterClass(uniWndClass.lpszClassName, hInstance);
		RegisterClass(&uniWndClass);
	}
	hWnd = CreateWindowEx(WS_EX_TOOLWINDOW, uniWndClass.lpszClassName, _T(""), WS_POPUP, 0, 0, 0, 0, 0, 0, hInstance, NULL);
	if (proc) SetWindowLong(hWnd, GWL_WNDPROC, (long) proc);
	return hWnd;
}

void DeallocateHWND(HWND hWnd)
{
	DestroyWindow(hWnd);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	hAppInstance = hInstance;
	// use mutex to restrict only one instance can be run
	dwMessageId = RegisterWindowMessage(ID_UNIQUE_APP);
	// hUniqueWnd = AllocateHWND(hInstance, (WNDPROC) UniCallbackProc);

	hMutex = OpenMutex(MUTEX_ALL_ACCESS, FALSE, ID_UNIQUE_APP);
	if (hMutex == NULL)
	{
		hMutex = CreateMutex(NULL, TRUE, ID_UNIQUE_APP);
	} else {
		BroadcastSystemMessage(BSF_IGNORECURRENTTASK || BSF_POSTMESSAGE, &dwBSMRecipients, dwMessageId, 0, 0);
		return -1;
	}

	hNoopWnd = CreateDialog(hAppInstance, MAKEINTRESOURCE(IDD_NOOP), NULL, (DLGPROC) NoopDialogProc); 
	hAboutWnd = CreateDialog(hAppInstance, MAKEINTRESOURCE(IDD_ABOUT), NULL, (DLGPROC) AboutDialogProc); 
	SendMessage(hNoopWnd,WM_SETICON,ICON_SMALL,(LPARAM)LoadImage(hAppInstance, MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,16,16,LR_DEFAULTCOLOR));
	SendMessage(hNoopWnd,WM_SETICON,ICON_BIG,(LPARAM)LoadIcon(hAppInstance, MAKEINTRESOURCE(IDI_APPICON)));

	lpCmdLine = GetCommandLine(); //this line is necessary for _ATL_MIN_CRT

    HRESULT hRes = CoInitialize(NULL);
    _ASSERTE(SUCCEEDED(hRes));

    _Module.Init(0, hInstance, &LIBID_ATLLib);

	DoScan();
	DoHook();

	MSG		msg;
	while (GetMessage(&msg, NULL, NULL, NULL))
	{
		if(msg.message == WM_TIMER && IsWindow(msg.hwnd ))
			SendMessage(msg.hwnd , TTM_RELAYEVENT, NULL, (LPARAM)&msg);
		if (!IsDialogMessage(hNoopWnd, &msg) && !IsDialogMessage(hAboutWnd, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		if(msg.message == WM_QUIT) break;
	}
	// if PostQuitMessage somewhere, it will go here, release allocated resources before exit!!!
	DoUnhook();

    _Module.Term();
    CoUninitialize();

	if (hMutex) 
	{
		ReleaseMutex(hMutex);
		CloseHandle(hMutex);
	}
	// DeallocateHWND(hUniqueWnd);

	return(msg.wParam);
}

VOID WINAPI OnTrayNotify(HWND hWnd,LPARAM lParam)
{
	POINT						ptCurPos;
	switch(LOWORD(lParam))
	{
		case WM_LBUTTONDBLCLK:
			// ShowWindow(hWnd, SW_SHOWNORMAL | SW_RESTORE);
			// BringWindowToTop(hWnd);
			// SetFocus(hWnd);
			// SetActiveWindow(hWnd);
			OnExecuteMenuItem(ID_ABOUT);
			break;
		case WM_RBUTTONUP:
			GetCursorPos(&ptCurPos);
			TrackPopupMenu(hPopMenu, 0, ptCurPos.x, ptCurPos.y, 0, hWnd, NULL);
			break;
	}
}

BOOL WINAPI OnInitDialogNoop(HWND hWnd) 
{
	//SystemTray Icon
	SystemTrayEx.SetIcon((HICON)LoadImage(hAppInstance, MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,16,16,LR_DEFAULTCOLOR));
	SystemTrayEx.hWnd = hWnd;
	SystemTrayEx.SetTipText("Ultra Hijack 1.0");
	SystemTrayEx.AddIcon();
	
	//Load Popup Menu
	hMainMenu = LoadMenu(hAppInstance, (CHAR *)IDR_MENU_MAIN);
	hPopMenu = GetSubMenu(hMainMenu, 0);
	hSysMenu = GetSystemMenu(hWnd, false);
	//AppendMenu(hSysMenu, MF_SEPARATOR, 0, NULL);
	//AppendMenu(hSysMenu, MF_POPUP, (int)hPopMenu, "Main Menu");

	//Add SystemTray Icon's & Controls' ToolTip
	ToolTip_Handle =  CreateWindowEx(WS_EX_TOOLWINDOW | WS_EX_TOPMOST, TOOLTIPS_CLASS, NULL, WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP, CW_USEDEFAULT, CW_USEDEFAULT,CW_USEDEFAULT, CW_USEDEFAULT, hWnd, NULL, hAppInstance, NULL);
	SetWindowPos(ToolTip_Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

	//Set default Item of the Menu
	MENUITEMINFO mii; 
	mii.cbSize = sizeof(mii);
	mii.fMask = MIIM_STATE;
	mii.fType = MFT_STRING;
	mii.dwTypeData = NULL;
	mii.cch = 0;

	return TRUE;
}

VOID WINAPI CenterWindow(HWND hWnd)
{
	int x,y;
	RECT r;
	x = GetSystemMetrics(SM_CXSCREEN);
	y = GetSystemMetrics(SM_CYSCREEN);
	GetWindowRect(hWnd,&r);
	x = (x + r.left - r.right)/2;
	y = (y + r.top - r.bottom)/2;
	SetWindowPos(hWnd,HWND_TOPMOST,x,y,0,0, SWP_ASYNCWINDOWPOS | SWP_NOSIZE);
}

VOID WINAPI OnExitApp() 
{ 
	SystemTrayEx.RemoveIcon();
	DestroyWindow(hAboutWnd);
	DestroyWindow(hNoopWnd);
	hAboutWnd = NULL;
	hNoopWnd = NULL;
	PostQuitMessage(0);
} 

VOID WINAPI OnMeasureItem(HWND hWnd, LPMEASUREITEMSTRUCT lpmis) 
{
	LPMENUITEM	lpmitem = (LPMENUITEM) lpmis->itemData; 
	HDC hdc = GetDC(hWnd); 
	HFONT hfntOld = (HFONT)SelectObject(hdc, lpmitem->hfont); 
	SIZE size; 
	GetTextExtentPoint32(hdc, lpmitem->psz, strlen(lpmitem->psz), &size); 

	lpmis->itemWidth = size.cx; 
	lpmis->itemHeight = size.cy; 

	SelectObject(hdc, hfntOld); 
	ReleaseDC(hWnd, hdc); 
} 
 
VOID WINAPI OnDrawItem(HWND hWnd, LPDRAWITEMSTRUCT lpdis) 
{ 
/*
	LPDRAWITEMSTRUCT lpDIS;
	lpDIS = (LPDRAWITEMSTRUCT)lParam;
	switch(lpDIS->itemID) {
		case ID_OPTIONS:
			break;
	}
	switch(lpDIS->CtlID) {
		case IDC_AUTOSTART:
			break;
	}
*/
} 
