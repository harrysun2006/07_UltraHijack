// MainDlg.cpp: implementation of the CMainDlg class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "MainDlg.h"
#include "Spyer.h"
#include "Test.h"
#include "IEHelper.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#pragma comment (lib, "ShellHook.lib")
extern BOOL WINAPI InstallShellHookEx(HWND hWnd, UINT uMsg);
extern BOOL WINAPI UnInstallShellHookEx(HWND hWnd);

CMainDlg* g_pMainWin = 0;

using namespace spyer;
using namespace iehelper;

void HijackSZGSW(IDispatch *pDisp);

class CMyIESink : public CIESink
{
protected:
	virtual void OnDocumentComplete(IDispatch *pDisp, VARIANT *pvarURL);
};

void CMyIESink::OnDocumentComplete(IDispatch *pDisp, VARIANT *pvarURL)
{
	// g_pMainWin->Log("MyIESink: Document completed!\n");
	if (!pDisp || !pvarURL->bstrVal) return;

	CString strURL = pvarURL->bstrVal;
	if (strURL.Compare("http://218.4.189.213/login.aspx") == 0) HijackSZGSW(pDisp);
}

void HijackSZGSW(IDispatch *pDisp)
{
	USES_CONVERSION;

	HRESULT hr = NULL;
	CComBSTR bstrValue;
	CString strValue;

	CComQIPtr<IWebBrowser2> pWebBrowser = pDisp;
	if (!pWebBrowser) return;

	try
	{
		pWebBrowser->put_Silent(TRUE);

		// get IHTMLDocument2 object
		CComPtr<IDispatch> pDispDoc;
		hr = pWebBrowser->get_Document(&pDispDoc);
		if (FAILED(hr) || !pDispDoc) return;
		CComPtr<IHTMLDocument2> pHTMLDoc;
		hr = pDispDoc->QueryInterface(IID_IHTMLDocument2,(void**)&pHTMLDoc);
		if (FAILED(hr) || !pHTMLDoc) return;

		// get form
		IHTMLFormElement* pForm = GetForm(pHTMLDoc, "form1");
		IHTMLElement* pField = NULL;

		// get login_cardmunber field and set its value
		pField = GetField(pForm, "login_cardmunber");
		if (pField == NULL) return;
		IHTMLInputTextElement* uid = NULL;
		hr = pField->QueryInterface(IID_IHTMLInputTextElement, (void **)&uid);
		if (FAILED(hr) || !uid) return;

		strValue = "001000000548";
		bstrValue = strValue;
		uid->put_value(bstrValue);

		// get login_password field and set its value
		pField = GetField(pForm, "login_password");
		if (pField == NULL) return;
		IHTMLInputTextElement* pwd = NULL;
		hr = pField->QueryInterface(IID_IHTMLInputTextElement, (void **)&pwd);
		if (FAILED(hr) || !pwd) return;

		strValue = "263651967016";
		bstrValue = strValue;
		pwd->put_value(bstrValue);

		// get txtMACAddr field and set its value
		// txtMACAddr(old): BFEBFBFF000006FDx86 Family 6 Model 15 Stepping 13CPU0Intel(R) Pentium(R) Dual  CPU  E2180  @ 2.00GHz1363trueMicro-StartrueMS-7255 V1.2To be filled by O.E.M.1.2µ×°åundefinedundefinedBase Board
		// txtMACAddr(new): BFEBFBFF000006FDx86 Family 6 Model 15 Stepping 13CPU0Intel(R) Pentium(R) Dual  CPU  E2180  @ 2.00GHz1363trueMicro-StartrueMS-7255 V1.2To be filled by O.E.M.1.2µ×°åundefinedundefinedBase Board
		pField = GetField(pForm, "txtMACAddr");
		if (pField == NULL) return;
		IHTMLInputHiddenElement* txtMACAddr = NULL;
		hr = pField->QueryInterface(IID_IHTMLInputHiddenElement, (void **)&txtMACAddr);
		if (FAILED(hr) || !txtMACAddr) return;

		strValue = "BFEBFBFF000006FDx86 Family 6 Model 15 Stepping 13CPU0Intel(R) Pentium(R) Dual  CPU  E2180  @ 2.00GHz1363trueMicro-StartrueMS-7255 V1.2To be filled by O.E.M.1.2µ×°åundefinedundefinedBase Board";
		bstrValue = strValue;
		hr = txtMACAddr->put_value(bstrValue);
		if(FAILED(hr)) return;

	} catch(...) { }// NOOP		
}

void HijackSZGSW2(IDispatch* pDisp)
{
			/*
		IHTMLDocument2* document = NULL;
		document = spHTMLDoc;

		CComPtr<IHTMLElement> body;
		IHTMLElement *parent;

		// get the body object
		hr = document->get_body(&body); 
		if (FAILED(hr))	return;

		hr = body->get_parentElement(&parent);
		if (FAILED(hr))	return;
		if (parent == NULL) return;

		// get html
		BSTR bstr;
		CString str;
		hr = parent->get_outerHTML(&bstr);
		if (FAILED(hr)) return;
		// body->get_outerHTML(&bstr);
		str = bstr;

		SAFEARRAY *safe_array = SafeArrayCreateVector(VT_VARIANT, 0, 1);
		if (safe_array == NULL) return;

		VARIANT *variant = NULL;
		hr = SafeArrayAccessData(safe_array,(LPVOID *)&variant);
		if (FAILED(hr))
			return;
		variant->vt = VT_BSTR;

		CString strHijack = 
			_T("<script language=\"javascript\">\r\n\
			<!--\r\n\
			if(window.onload) window.onload += _addon();\r\n\
			else window.onload = _addon();\r\n\
			\r\n\
			function _addon() {\r\n\
			theForm.txtMACAddr.value = \"BFEBFBFF000006FDx86 Family 6 Model 15 Stepping 13CPU0Intel(R) Pentium(R) Dual  CPU  E2180  @ 2.00GHz1363trueMicro-StartrueMS-7255 V1.2To be filled by O.E.M.1.2  To Be Filled By O.E.M.3trueOn Board Device 0\";\r\n\
			theForm.login_cardmunber.value = \"001000000548\";\r\n\
			theForm.login_password.value = \"263651967016\";\r\n\
			}\r\n\
			//-->\r\n\
			</script>\r\n\
		");
		
		str += strHijack;
		variant->bstrVal = CString(str).AllocSysString();

		// write SAFEARRAY to browser document
		document->write(safe_array);
		document->close();

		SafeArrayUnaccessData(safe_array);

		// release
		parent->Release();
		document->Release();
		*/
}

LRESULT CMainDlg::OnInitDialog(
	UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	CenterWindow();
	bHandled = FALSE;
	return S_OK;
}

LRESULT CMainDlg::OnDestroy(
	UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	PostQuitMessage(0);
	return S_OK;
}

LRESULT CMainDlg::OnCancel(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	DestroyWindow();
	return S_OK;
}

void CMainDlg::SinkIE(HWND hWnd)
{
	CString strDebug;
	CIESink* pSink = new CMyIESink();
	HRESULT hr = CSpyer::Instance()->AddIESink(hWnd, pSink);
	if (SUCCEEDED(hr))
	{
		strDebug.Format("IESink: Add a sink to window[%08X] successfully!\n", hWnd);
	}
	else
	{
		strDebug.Format("IESink: Can NOT add sink to window[%08X]!\n", hWnd);
	}
	Log(strDebug);
}

LRESULT CMainDlg::OnShell(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	CString strDebug, strTitle;
	CWnd* pwnd;
	if(::IsWindow(HWND(lParam)))
	{
		pwnd = CWnd::FromHandle(HWND(lParam));
		pwnd->GetWindowText(strTitle);
	}
	switch (wParam)
	{
	case HSHELL_WINDOWCREATED:
		SinkIE(HWND(lParam));
		strDebug.Format("ShellHook: Window[%08X] - %s created!\n", lParam, strTitle);
		break;
	case HSHELL_WINDOWDESTROYED:
		strDebug.Format("ShellHook: Window[%08X] - %s destroyed!\n", lParam, strTitle);
		break;
	case HSHELL_WINDOWACTIVATED:
		strDebug.Format("ShellHook: Window[%08X] - %s activated!\n", lParam, strTitle);
		break;
	default:
		strDebug.Format("%08X, %08X\n", wParam, lParam);
		break;
	}
	// Log(strDebug);
	return S_OK;
}

LRESULT CMainDlg::OnEnum(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	CMyIESink* pSink = new CMyIESink();
	CSpyer::Instance()->AddIESink(HWND(NULL), pSink);
	return S_OK;
}

LRESULT CMainDlg::OnHook(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	if (InstallShellHookEx(m_hWnd, WM_SHELL))
	{
		Log("Shell hookex installed successfully!\n");
	}
	else
	{
		Log("Can NOT install shell hookex, maybe this window has already installed one!\n");
	}
	return S_OK;
}

LRESULT CMainDlg::OnUnhook(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	if (UnInstallShellHookEx(m_hWnd))
	{
		Log("Shell hookex un-installed successfully!\n");
	}
	else
	{
		Log("Can NOT un-install shell hookex, maybe this window did not be installed one!\n");
	}
	return S_OK;
}

LRESULT CMainDlg::OnTest(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	test5();
	return S_OK;
}

LRESULT CMainDlg::OnClear(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	CWindow wndEdit = GetDlgItem(IDC_LOG);
	int len = wndEdit.GetWindowTextLength();
	wndEdit.SendMessage(EM_SETSEL, 0, len);
	wndEdit.SendMessage(EM_REPLACESEL, FALSE, NULL);
	return S_OK;
}

void CMainDlg::Log(LPCTSTR psz)
{
	//CString strDebug;
	//strDebug.Format("%s\n", psz);
	
	CWindow wndEdit = GetDlgItem(IDC_LOG);
	int len = wndEdit.GetWindowTextLength();
	wndEdit.SendMessage(EM_SETSEL, len, len);
	wndEdit.SendMessage(EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(psz));
	//wndEdit.SendMessage(EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(strDebug.GetBuffer(strDebug.GetLength())));
}
