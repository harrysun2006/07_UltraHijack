#include "StdAfx.h"
#include "MainDlg.h"
#include "Test.h"
#include "IEHelper.h"

#include <oleacc.h>
#include <afxctl.h>

#pragma comment (lib, "oleacc")

#include <mshtml.h>
#include <map>
#include <string>

using namespace std;
using namespace iehelper;

void Log(LPCTSTR psz);

void test1()
{
	IWebBrowser2 *m_pInternetExplorer; 
	CoCreateInstance(CLSID_InternetExplorer,NULL,CLSCTX_SERVER,IID_IWebBrowser2,(LPVOID *)&m_pInternetExplorer);
	m_pInternetExplorer->put_Visible(VARIANT_TRUE);

	BSTR Url,Target,PostData,Head;
	VARIANT BstrUrl,BstrTarget,IFlag,BstrPostData,BstrHead;
	V_VT(&BstrUrl)=VT_BSTR;
	V_BSTR(&BstrUrl)=Url=SysAllocString(L"http://www.google.com");
	V_VT(&BstrTarget)=VT_BSTR;
	V_BSTR(&BstrTarget)=Target=SysAllocString(L"_self");
	V_VT(&BstrPostData)=VT_BSTR;
	V_BSTR(&BstrPostData)=PostData=SysAllocString(L"");
	V_VT(&BstrHead)=VT_BSTR;
	V_BSTR(&BstrHead)=Head=SysAllocString(L"");
	V_VT(&IFlag)=VT_I4;
	V_I4(&IFlag)=navNoHistory;

	m_pInternetExplorer->Navigate2(&BstrUrl,&IFlag,&BstrTarget,&BstrPostData,&BstrHead);
	SysFreeString(Url);
	SysFreeString(Target);
	SysFreeString(PostData);
	SysFreeString(Head);
	//MonitorWebBrowser(m_pInternetExplorer);
	m_pInternetExplorer->Release();
}

void test2()
{
	IID id = DIID_DWebBrowserEvents2;
	LPOLESTR strGUID = 0;
	CComBSTR bstrId = NULL;
	// 注意此处StringFromIID的正确用法, 否则会在Debug状态下出错: user breakpoint called from code 0x7c901230
	StringFromIID(id, &strGUID);
	bstrId = strGUID;
	CoTaskMemFree(strGUID);
	USES_CONVERSION;
	Log(OLE2CT(bstrId));
}

void test3()
{
	CString strDebug;
	typedef map<string, string>maps;
	typedef pair<string, string>pr;
	maps temp;
	temp.insert(pr("aa","aaaaa"));
	temp.insert(pr("bb","bbbbbb"));
	temp.insert(pr("cc","cccc"));
	maps   tent;
	maps::iterator   it;
	for(it = temp.begin(); it != temp.end(); ++it) tent.insert(*it);
	for(it = tent.begin(); it != tent.end(); ++it)
	{
		strDebug.Format("%s ==> %s\n", (*it).first, (*it).second);
		Log(strDebug);
	}
}


BOOL CALLBACK EnumIEWindow(HWND hWnd, LPARAM lParam)
{
	TCHAR szClassName[256];
	GetClassName(hWnd,  szClassName,  sizeof(szClassName));

	if (_tcscmp(szClassName, _T("Internet Explorer_Server")) == 0)
	{
		*(HWND*)lParam = hWnd;
		return FALSE;
	}
	return TRUE;
}

void test4()
{
	HWND hWnd = NULL;
	//hWnd = FindWindowEx(NULL, NULL, "IEFrame", NULL);
	EnumChildWindows(GetDesktopWindow(), EnumIEWindow, (LPARAM)&hWnd);
	if (hWnd) 
	{
		// UINT nMsg = RegisterWindowMessage(_T("WM_HTML_GETOBJECT"));
		UINT nMsg = RegisterWindowMessage(_T("WM_HTML_GETOBJECT"));
		LRESULT lRes;
		SendMessageTimeout(hWnd, nMsg, 0L, 0L, SMTO_ABORTIFHUNG, 1000, (DWORD*)&lRes);

		HRESULT hr = S_OK;

		CComPtr<IHTMLDocument2> spDocument = NULL;
		hr = ObjectFromLresult (lRes, IID_IHTMLDocument2, 0 , (LPVOID*)&spDocument);
		if (FAILED(hr) || !spDocument) Log(_T("Can NOT found IHTMLDocument2 object!\n"));
		else Log(_T("Found IHTMLDocument2 object!\n"));

		CComPtr<IWebBrowser2> spWebBrowser = NULL;
		hr = ObjectFromLresult (lRes, IID_IWebBrowser2, 0 , (LPVOID*)&spWebBrowser);
		if (FAILED(hr) || !spWebBrowser) Log(_T("Can NOT found IWebBrowser2 object!\n"));
		else Log(_T("Found IWebBrowser2 object!\n"));

		CComPtr<::IServiceProvider> spServiceProvider = NULL;
		hr = ObjectFromLresult (lRes, ::IID_IServiceProvider, 0 , (LPVOID*)&spServiceProvider);
		if (FAILED(hr) || !spServiceProvider) Log(_T("Can NOT found IServiceProvider object!\n"));
		else Log(_T("Found IServiceProvider object!\n"));

		CComPtr<IWebBrowserApp> spWebBrowserApp = NULL;
		hr = ObjectFromLresult (lRes, IID_IWebBrowserApp, 0 , (LPVOID*)&spWebBrowserApp);
		if (FAILED(hr) || !spWebBrowserApp) Log(_T("Can NOT found IWebBrowserApp object!\n"));
		else Log(_T("Found IWebBrowserApp object!\n"));
	}
	else
	{
		Log(_T("Can NOT found IE window!\n"));
	}
}

void test5()
{
	CComPtr<IShellWindows> spShellWin;
	HRESULT hr = spShellWin.CoCreateInstance(CLSID_ShellWindows);
	if (FAILED(hr)) {
		Log(_T("Can NOT create CLSID_ShellWindows object!\n"));
		return;
	}
	USES_CONVERSION;
	CComBSTR bstrURL, bstrValue;
	HWND hWnd;
	CString strDebug, strURL, strValue;
	long nCount = 0;
	spShellWin->get_Count(&nCount);
	strDebug.Format(_T("Found %d shell windows!\n"), nCount);
	Log(strDebug);

	for(int i = 0; i < nCount; i++)
	{
		CComPtr<IDispatch> pDisp;
		hr = spShellWin->Item(CComVariant((long)i), &pDisp);
		if (FAILED (hr)) {
			Log(_T("Failed to get item from shell window!\n"));
			continue;
		}
		CComQIPtr<IWebBrowser2> pWebBrowser = pDisp;
		if (!pWebBrowser) {
			Log(_T("Failed to get IWebBrowser2 object from the item!\n"));
			continue;
		}
		hWnd = GetWindowHandle(pWebBrowser);
		pWebBrowser->get_LocationURL(&bstrURL);
		strURL = bstrURL ? OLE2CT(bstrURL) : _T("");
		
		strDebug.Format(_T("[%08X]URL: %s"), hWnd, strURL);
		strDebug.Format("hWnd = %08X; URL = %s\n", hWnd, strURL);
		Log(strDebug);

		if (strURL.Compare(_T("http://218.4.189.213/login.aspx")) != 0) continue;
		try
		{
			pWebBrowser->put_Silent(TRUE);

			// get document
			CComPtr<IDispatch> pDispDoc;
			hr = pWebBrowser->get_Document(&pDispDoc);
			if (FAILED(hr) || !pDispDoc) {
				Log(_T("Can NOT find document object from IWebBrowser2!\n"));
				return;
			}
			CComPtr<IHTMLDocument2> pHTMLDoc;
			hr = pDispDoc->QueryInterface(IID_IHTMLDocument2,(void**)&pHTMLDoc);
			if (FAILED(hr) || !pHTMLDoc) {
				Log(_T("Can NOT get IHTMLDocument2 object from the IDispatch object!\n"));
				return;
			}
			EnumDocument(pHTMLDoc, NULL);
			IHTMLFormElement* pForm = GetForm(pHTMLDoc, "form1");
			IHTMLElement* pField = NULL;

			pField = GetField(pForm, "login_cardmunber");
			if (pField == NULL) {
				Log(_T("Can NOT find login_cardmunber field!\n"));
				return;
			}
			IHTMLInputTextElement* uid = NULL;
			hr = pField->QueryInterface(IID_IHTMLInputTextElement, (void **)&uid);
			if (FAILED(hr) || !uid) {
				Log(_T("The login_cardmunber is NOT a text input field!\n"));
				return;
			}
			strValue = "001000000548";
			bstrValue = strValue;
			uid->put_value(bstrValue);

			pField = GetField(pForm, "login_password");
			if (pField == NULL) {
				Log(_T("Can NOT find login_password field!\n"));
				return;
			}
			IHTMLInputTextElement* pwd = NULL;
			hr = pField->QueryInterface(IID_IHTMLInputTextElement, (void **)&pwd);
			if (FAILED(hr) || !pwd) {
				Log(_T("The login_password is NOT a text input field!\n"));
				return;
			}
			strValue = "263651967016";
			bstrValue = strValue;
			pwd->put_value(bstrValue);

			pField = GetField(pForm, "txtMACAddr");
			if (pField == NULL) {
				Log(_T("Can NOT find txtMACAddr field!\n"));
				return;
			}
			IHTMLInputHiddenElement* txtMACAddr = NULL;
			hr = pField->QueryInterface(IID_IHTMLInputHiddenElement, (void **)&txtMACAddr);
			if (FAILED(hr) || !txtMACAddr) {
				Log(_T("The txtMACAddr is NOT a hidden input field!\n"));
				return;
			}

			hr = txtMACAddr->get_value(&bstrValue);
			if(FAILED(hr)) {
				Log(_T("Can NOT get the old value of txtMACAddr field!\n"));
				return;
			}
			strValue = bstrValue ? OLE2CT(bstrValue) : _T("");
			strDebug.Format("txtMACAddr(old): %s\n", strValue);
			Log(strDebug);

			strValue = "BFEBFBFF000006FDx86 Family 6 Model 15 Stepping 13CPU0Intel(R) Pentium(R) Dual  CPU  E2180  @ 2.00GHz1363trueMicro-StartrueMS-7255 V1.2To be filled by O.E.M.1.2底板undefinedundefinedBase Board";
			bstrValue = strValue;
			hr = txtMACAddr->put_value(bstrValue);
			if(FAILED(hr)) {
				Log(_T("Can NOT set the value of txtMACAddr field!\n"));
				return;
			}
			//txtMACAddr(old): BFEBFBFF000006FDx86 Family 6 Model 15 Stepping 13CPU0Intel(R) Pentium(R) Dual  CPU  E2180  @ 2.00GHz1363trueMicro-StartrueMS-7255 V1.2To be filled by O.E.M.1.2底板undefinedundefinedBase Board
			//txtMACAddr(new): BFEBFBFF000006FDx86 Family 6 Model 15 Stepping 13CPU0Intel(R) Pentium(R) Dual  CPU  E2180  @ 2.00GHz1363trueMicro-StartrueMS-7255 V1.2To be filled by O.E.M.1.2底板undefinedundefinedBase Board


			hr = txtMACAddr->get_value(&bstrValue);
			if(FAILED(hr)) {
				Log(_T("Can NOT get the new value of txtMACAddr field!\n"));
				return;
			}
			strValue = bstrValue ? OLE2CT(bstrValue) : _T("");
			strDebug.Format("txtMACAddr(new): %s\n", strValue);
			Log(strDebug);

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

		} catch(...) { }// NOOP		
	}
}

void Log(LPCTSTR psz)
{
	g_pMainWin->Log(psz);
}
