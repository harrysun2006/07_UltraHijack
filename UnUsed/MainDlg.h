// MainDlg.h: interface for the CMainDlg class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(MAINDLG__INCLUDED_)
#define MAINDLG__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"

#define WM_SHELL WM_USER + 1

class CMainDlg : public CAxDialogImpl<CMainDlg>  
{
public:
	enum {IDD = IDD_MAINDLG};

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_SHELL, OnShell)
		COMMAND_ID_HANDLER(IDCANCEL, OnCancel)

		COMMAND_HANDLER(IDC_ENUM, BN_CLICKED, OnEnum)
		COMMAND_HANDLER(IDC_HOOK, BN_CLICKED, OnHook)
		COMMAND_HANDLER(IDC_UNHOOK, BN_CLICKED, OnUnhook)
		COMMAND_HANDLER(IDC_TEST, BN_CLICKED, OnTest)
		COMMAND_HANDLER(IDC_CLEAR, BN_CLICKED, OnClear)
	END_MSG_MAP()

	void Log(LPCTSTR psz);
private:
	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnShell(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnCancel(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	LRESULT OnEnum(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	LRESULT OnHook(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	LRESULT OnUnhook(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	LRESULT OnTest(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	LRESULT OnClear(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	void SinkIE(HWND hWnd);
};

extern CMainDlg* g_pMainWin;

#endif // !defined(MAINDLG__INCLUDED_)
