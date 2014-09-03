// IEHelper.cpp : IE helper class
//

#include "stdafx.h"
#include "IEHelper.h"

#include <mshtml.h>		// All definitions of IHTMLxxxx interfaces
#include <oleacc.h>
#include <afxctl.h>

#pragma comment (lib, "oleacc")

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

using namespace std;

namespace iehelper
{

CComPtr<IWebBrowser2> GetWebBrowser2(IHTMLDocument2* pHTMLDocument2)
{
	if(!pHTMLDocument2) return NULL;

	HRESULT hr = S_OK;
	IWebBrowser2 *pWebBrowser2 = NULL;
	CComPtr<::IServiceProvider> pServiceProvider = NULL;

	hr = pHTMLDocument2->QueryInterface(::IID_IServiceProvider, (void**)&pServiceProvider);
	if(SUCCEEDED(hr) && pServiceProvider)
	{
		hr = pServiceProvider->QueryService(SID_SInternetExplorer, IID_IWebBrowser2, (void**)&pWebBrowser2);
	}

	return pWebBrowser2;
}

CComPtr<IHTMLDocument2> GetHTMLDocument2(IWebBrowser2* pWebBrowser2)
{
	if (!pWebBrowser2) return NULL;

	CComPtr<IDispatch> pDispatch;
	HRESULT hr = pWebBrowser2->get_Document(&pDispatch);
	if (FAILED(hr) || !pDispatch) return NULL;

	CComPtr<IHTMLDocument2> pHTMLDocument2;
	hr = pDispatch->QueryInterface(IID_IHTMLDocument2,(void**)&pHTMLDocument2);
	if (FAILED(hr) || !pHTMLDocument2) return NULL;

	return pHTMLDocument2;
}

CComPtr<IHTMLWindow2> GetHTMLWindow2(IHTMLDocument2* pHTMLDocument2)
{
	if (!pHTMLDocument2) return NULL;

	CComPtr<IHTMLWindow2> pHTMLWindow2;
	HRESULT hr = pHTMLDocument2->get_parentWindow(&pHTMLWindow2);
	return pHTMLWindow2;
}

// GetHTMLWindow2 must return CComPtr<IHTMLWindow2>, otherwise get "Access Violation" exception!!!
CComPtr<IHTMLWindow2> GetHTMLWindow2(IWebBrowser2* pWebBrowser2)
{
	return GetHTMLWindow2(GetHTMLDocument2(pWebBrowser2));
}

CComPtr<IHTMLFormElement> GetHTMLForm(IHTMLDocument2* pHTMLDocument2, LPCTSTR lpszName)
{
	if(!pHTMLDocument2) return NULL;

	HRESULT hr;
	USES_CONVERSION;

	CComQIPtr<IHTMLElementCollection> pHTMLForms;
	hr = pHTMLDocument2->get_forms(&pHTMLForms);
	if (FAILED(hr)) return NULL;

	long nCount = 0;
	hr = pHTMLForms->get_length(&nCount);
	if (FAILED(hr)) return NULL;

	IDispatch *pDisp = NULL;
	CComBSTR bstrName;
	CString strName;

	for(long i = 0; i < nCount; i++)
	{
		
		hr = pHTMLForms->item(CComVariant(i), CComVariant(), &pDisp);
		if (FAILED(hr)) continue;

		CComQIPtr<IHTMLFormElement> pHTMLForm = pDisp;
		// here do NOT release the pDisp, or the spFormElement will be released also
		// pDisp->Release();

		pHTMLForm->get_name(&bstrName);
		strName = bstrName ? OLE2CT(bstrName) : _T("");
		if (strName.Compare(lpszName) == 0) return pHTMLForm;
	}
	return NULL;
}

CComPtr<IHTMLElement> GetFormInput(IHTMLFormElement* pHTMLForm, LPCTSTR lpszName)
{
	if(!pHTMLForm) return NULL;

	HRESULT hr;
	USES_CONVERSION;

	IDispatch *pDisp = NULL;
	CComBSTR bstrName;
	CString strName;

	long nCount = 0;
	hr = pHTMLForm->get_length(&nCount);
	if (FAILED(hr))	return NULL;

	for(long j = 0; j < nCount; j++)
	{
		hr = pHTMLForm->item(CComVariant(j), CComVariant(), &pDisp);
		if (FAILED(hr)) continue;

		CComQIPtr<IHTMLElement> pFormInput = pDisp;
		// here do NOT release the pDisp, or the spInputElement will be released also
		// pDisp->Release();

		hr = pFormInput->get_id(&bstrName);
		if(FAILED(hr)) continue;
		strName = bstrName ? OLE2CT(bstrName) : _T("");
		if (strName.Compare(lpszName) == 0) return pFormInput;
	}

	return NULL;
}

HRESULT ExecuteScript(IHTMLWindow2* pHTMLWindow2, LPCTSTR pszScript, LPCTSTR pszLang)
{
	HRESULT hr = NULL;

	USES_CONVERSION;

	try
	{
		/*
		pWebBrowser->put_Silent(TRUE);

		// get IHTMLDocument2 object
		// Release下有线程同步和函数重入的问题, execScript时hr会返回E_INVALIDARG, 使用CoMarshalInterThreadInterfaceInStream解决问题
		CComPtr<IDispatch> pDispDocument;
		hr = pWebBrowser->get_Document(&pDispDocument);
		if (FAILED(hr) || !pDispDocument) return;
		CComPtr<IHTMLDocument2> pDocument;
		hr = pDispDocument->QueryInterface(IID_IHTMLDocument2,(void**)&pDocument);
		if (FAILED(hr) || !pDocument) return;
		
		IStream * stream     = NULL;
		IUnknown *unknownPtr = NULL;
		pDocument->QueryInterface(IID_IUnknown, (void **)&unknownPtr);
		if (unknownPtr != NULL)
		{
			CoMarshalInterThreadInterfaceInStream(__uuidof(IHTMLDocument2), unknownPtr, &stream);
			unknownPtr->Release();
		}
		*/

		CComBSTR	comScript, comLang = pszLang;
		BSTR			bstrLang;
		VARIANT		dummy;

		comScript   = pszScript;
		bstrLang = SysAllocString(comLang);

		hr = pHTMLWindow2->execScript(comScript, bstrLang, &dummy);
		SysFreeString(bstrLang);
	} catch(...) { }
	return hr;
}

}