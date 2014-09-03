#include "stdafx.h"
#include "SzgswHijackSink.h"
#include "Spyer.h"
#include "IEHelper.h"

#define	CARD_NUMBER		_T("001000000548")
#define CARD_PASSWORD	_T("263651967016")
#define HARDWARE_CODE	_T("BFEBFBFF000006FDx86 Family 6 Model 15 Stepping 13CPU0Intel(R) Pentium(R) Dual  CPU  E2180  @ 2.00GHz1363trueMicro-StartrueMS-7255 V1.2To be filled by O.E.M.1.2底板undefinedundefinedBase Board")
#define DEBUG_LEVEL		0

using namespace spyer;
using namespace iehelper;

void HijackLogin1(IDispatch *pDispatch);
void HijackLogin2(IDispatch *pDispatch);
void HijackMenu(IDispatch *pDispatch);
#if DEBUG_LEVEL > 2
void HijackInfo(IDispatch *pDispatch);
#endif

void CSzgswHijackSink::OnDocumentComplete(IDispatch *pDispatch, VARIANT *pvarURL)
{
	// Debug("MyIESink: Document completed!\n");
	if (!pDispatch || !pvarURL->bstrVal) return;

	CString strURL = pvarURL->bstrVal;
	if (strURL.Compare("http://218.4.189.213/login.aspx") == 0) HijackLogin(pDispatch);
	if (strURL.Find("http://218.4.189.213", 0) == 0) HijackMenu(pDispatch);
#if DEBUG_LEVEL > 2
	if (strURL.Find("http://218.4.189.213/CorpSel/Corp_info.aspx", 0) == 0) HijackInfo(pDispatch);
#endif
}

#if DEBUG_LEVEL > 2
void HijackInfo(IDispatch *pDispatch)
{
	HRESULT hr = NULL;

	CComQIPtr<IWebBrowser2> pWebBrowser2 = pDispatch;
	if (!pWebBrowser2) return;

	USES_CONVERSION;

	CComBSTR bstrURL, bstrHTML;
	CString strURL, strHTML, strItem;
	int i = 0, j, k = 0;
	char items[][12] = {_T("注册号"), _T("企业名称"), _T("法定代表人"), _T("公司地址"), _T("注册资本"),
		_T("联系电话"), _T("邮政编码"), _T("行业"), _T("管区"), _T("开业日期"), _T("登记机关"), _T("企业类型")};
	const char* pat = _T("InfoShowImg.aspx?StrCode=");
	const int	pat_len = strlen(pat);

	try
	{
		pWebBrowser2->put_Silent(TRUE);

		// get IHTMLDocument2 object
		CComPtr<IHTMLDocument2> pHTMLDocument2 = GetHTMLDocument2(pWebBrowser2);

		CComPtr<IHTMLElement> pBody = NULL;
		hr = pHTMLDocument2->get_body(&pBody);
		if (FAILED(hr) || !pBody) return;
		
		CComQIPtr <IHTMLElement> pHTML;
		hr = pBody->get_parentElement(&pHTML);
		if (FAILED(hr) || !pHTML) return;
		pHTML->get_outerHTML(&bstrHTML);
		strHTML = bstrHTML ? OLE2CT(bstrHTML) : _T("");

		pWebBrowser2->get_LocationURL(&bstrURL);
		strURL = bstrURL ? OLE2CT(bstrURL) : _T("");

		FILE *log = fopen("UltraHijack.log", "a");
		i = strURL.Find("corp_id=", 0);
		strURL = strURL.Mid(i + strlen("corp_id="), strURL.GetLength() - i - strlen("corp_id="));
		fprintf(log, "%s\n", strURL);

		i = 0;
		while(k < sizeof(items)/12)
		{
			i = strHTML.Find(pat, i);
			j = strHTML.Find("\"", i);
			strItem = strHTML.Mid(i + pat_len, j - i - pat_len);
			// fprintf(log, "%s: %s\r\n", items[k], strItem.c_str());
			fprintf(log, "%s\n", strItem);
			k++;
			i += pat_len;
		}
		fprintf(log, "\n");
		fclose(log);
		if (pBody) pBody = NULL;
		if (pHTMLDocument2) pHTMLDocument2 = NULL;
	} catch(...) { }

}
#endif

void HijackLogin(IDispatch *pDispatch)
{
	HijackLogin2(pDispatch);
}

void HijackLogin1(IDispatch *pDispatch)
{
	/*
	HRESULT hr = NULL;
	CComBSTR bstrURL, bstrValue;
	CString strURL, strValue;

	CComQIPtr<IWebBrowser2> pWebBrowser2 = pDispatch;
	if (!pWebBrowser2) return;

	USES_CONVERSION;

	pWebBrowser2->get_LocationURL(&bstrURL);
	strURL = bstrURL ? OLE2CT(bstrURL) : _T("");
	if (strURL.Compare("http://218.4.189.213/login.aspx") != 0) return;

	try
	{
		pWebBrowser2->put_Silent(TRUE);

		// get IHTMLDocument2 object
		CComPtr<IDispatch> pDispDoc;
		hr = pWebBrowser2->get_Document(&pDispDoc);
		if (FAILED(hr) || !pDispDoc) return;
		CComPtr<IHTMLDocument2> pHTMLDocument2;
		hr = pDispDoc->QueryInterface(IID_IHTMLDocument2,(void**)&pHTMLDocument2);
		if (FAILED(hr) || !pHTMLDocument2) return;

		// get form
		IHTMLFormElement* pForm = GetHTMLForm(pHTMLDocument2, "form1");
		IHTMLElement* pField = NULL;

		// get login_cardmunber field and set its value
		pField = GetFormInput(pForm, "login_cardmunber");
		if (pField == NULL) return;
		IHTMLInputTextElement* uid = NULL;
		hr = pField->QueryInterface(IID_IHTMLInputTextElement, (void **)&uid);
		if (FAILED(hr) || !uid) return;

		strValue = CARD_NUMBER;
		bstrValue = strValue;
		uid->put_value(bstrValue);

		// get login_password field and set its value
		pField = GetFormInput(pForm, "login_password");
		if (pField == NULL) return;
		IHTMLInputTextElement* pwd = NULL;
		hr = pField->QueryInterface(IID_IHTMLInputTextElement, (void **)&pwd);
		if (FAILED(hr) || !pwd) return;

		strValue = CARD_PASSWORD;
		bstrValue = strValue;
		pwd->put_value(bstrValue);

		// get TxtVidCode field and set the focus
		pField = GetFormInput(pForm, "TxtVidCode");
		IHTMLInputTextElement* vidCode = NULL;
		hr = pField->QueryInterface(IID_IHTMLInputTextElement, (void **)&vidCode);
		vidCode->select();

		// get txtMACAddr field and set its value
		// txtMACAddr(old): BFEBFBFF000006FDx86 Family 6 Model 15 Stepping 13CPU0Intel(R) Pentium(R) Dual  CPU  E2180  @ 2.00GHz1363trueMicro-StartrueMS-7255 V1.2To be filled by O.E.M.1.2底板undefinedundefinedBase Board
		// txtMACAddr(new): BFEBFBFF000006FDx86 Family 6 Model 15 Stepping 13CPU0Intel(R) Pentium(R) Dual  CPU  E2180  @ 2.00GHz1363trueMicro-StartrueMS-7255 V1.2To be filled by O.E.M.1.2底板undefinedundefinedBase Board
		pField = GetFormInput(pForm, "txtMACAddr");
		if (pField == NULL) return;
		IHTMLInputHiddenElement* txtMACAddr = NULL;
		hr = pField->QueryInterface(IID_IHTMLInputHiddenElement, (void **)&txtMACAddr);
		if (FAILED(hr) || !txtMACAddr) return;

		strValue = HARDWARE_CODE;
		bstrValue = strValue;
		hr = txtMACAddr->put_value(bstrValue);
		if(FAILED(hr)) return;

	} catch(...) { }// NOOP
	*/
}

void HijackLogin2(IDispatch* pDispatch)
{
	HRESULT hr = NULL;
	CComBSTR bstrURL;
	CString strURL;

	CComQIPtr<IWebBrowser2> pWebBrowser2 = pDispatch;
	if (!pWebBrowser2) return;

	USES_CONVERSION;

	pWebBrowser2->get_LocationURL(&bstrURL);
	strURL = bstrURL ? OLE2CT(bstrURL) : _T("");
	if (strURL.Compare("http://218.4.189.213/login.aspx") != 0) return;

	try
	{
		pWebBrowser2->put_Silent(TRUE);

		CComPtr<IHTMLWindow2> pHTMLWindow2 = GetHTMLWindow2(pWebBrowser2);
		if (!pHTMLWindow2) return;
		/*
		CComPtr<IDispatch> pDispatch;
		hr = pWebBrowser2->get_Document(&pDispatch);
		if (FAILED(hr) || !pDispatch) return;

		CComPtr<IHTMLDocument2> pHTMLDocument2;
		hr = pDispatch->QueryInterface(IID_IHTMLDocument2,(void**)&pHTMLDocument2);
		if (FAILED(hr) || !pHTMLDocument2) return;

		CComPtr<IHTMLWindow2> pHTMLWindow2;
		hr = pHTMLDocument2->get_parentWindow(&pHTMLWindow2);
		if (FAILED(hr) || !pHTMLWindow2) return;
		*/

		CString strScript;
		strScript.Format(_T("theForm.login_cardmunber.value = \"%s\";\r\n\
			theForm.login_password.value = \"%s\";\r\n\
			theForm.txtMACAddr.value = \"%s\";\r\n\
			theForm.TxtVidCode.select();\r\n"), CARD_NUMBER, CARD_PASSWORD, HARDWARE_CODE);
		
		hr = ExecuteScript(pHTMLWindow2, strScript);
		if (pHTMLWindow2) pHTMLWindow2 = NULL;
	} catch(...) {}
}

void HijackMenu(IDispatch *pDispatch)
{
	HRESULT hr = NULL;

	CComQIPtr<IWebBrowser2> pWebBrowser2 = pDispatch;
	if (!pWebBrowser2) return;

	USES_CONVERSION;

	try
	{
		pWebBrowser2->put_Silent(TRUE);

		CComPtr<IHTMLWindow2> pHTMLWindow2 = GetHTMLWindow2(pWebBrowser2);
		if (!pHTMLWindow2) return;

		CString strScript = _T("document.onselectstart = function() {return true;};\r\n\
			document.onmousedown = null;\r\n\
			document.oncontextmenu = null;\r\n");

		hr = ExecuteScript(pHTMLWindow2, strScript);
		if (pHTMLWindow2) pHTMLWindow2 = NULL;
	} catch(...) { }
}