// IESeeker.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "IESpy.h"
#include "MainDlg.h"
#include "IESeeker.h"
#include "IESink2.h"
#include "FlashSink.h"

//#import "C:\\WINDOWS\\system32\\Macromed\\Flash\\Flash9d.ocx" named_guids 
#include "Flash9d.tlh"

#include <mshtml.h>		// 所有 IHTMLxxxx 的接口声明
#include <oleacc.h>
#include <afxctl.h>

#pragma comment (lib, "oleacc")

using namespace std;
using namespace ShockwaveFlashObjects;

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

namespace seeker {
struct GUID_HELPER {
	char *name;
	GUID iid;
};

void EnumIE1();								//枚举浏览器函数 -- 通过IShellWindows取得实例
void EnumIE2();								//枚举浏览器函数 -- 向窗口发消息取得文档的接口指针
void EnumDocument(IHTMLDocument2 *pIHTMLDocument2, HWND hWnd = NULL);	//枚举文档函数
void AdviceDocument(IHTMLDocument2 *pIHTMLDocument2);
void EnumFrames(IHTMLDocument2 *pIHTMLDocument2);	//枚举子框架函数
void EnumForms(IHTMLDocument2 *pIHTMLDocument2);	//枚举表单函数
void EnumObjects(IHTMLDocument2 *pIHTMLDocument2);	//枚举对象函数
void EnumEmbeds(IHTMLDocument2 *pIHTMLDocument2);	//枚举对象函数
void EnumInterface(IDispatch *pDisp, LPCSTR comment = NULL);
void EnumConnectionPoint(IDispatch *pDisp, LPCSTR comment = NULL);
void MonitorFlash(IDispatch *pDisp);				//查看连接点
void MonitorWebBrowser(IWebBrowser2 *pWebBrowser2);
char* GetIIDName(IID iid);

static const GUID_HELPER COM_IIDS[] = 
{
	// flash9d.ocx
	{_T("ShockwaveFlashObjects"), {0xd27cdb6b,0xae6d,0x11cf,{0x96,0xb8,0x44,0x45,0x53,0x54,0x00,0x00}}},
	{_T("IShockwaveFlash"), {0xd27cdb6c,0xae6d,0x11cf,{0x96,0xb8,0x44,0x45,0x53,0x54,0x00,0x00}}},
	{_T("IShockwaveFlashEvents"), {0xd27cdb6d,0xae6d,0x11cf,{0x96,0xb8,0x44,0x45,0x53,0x54,0x00,0x00}}},
	{_T("ShockwaveFlash"), {0xd27cdb6e,0xae6d,0x11cf,{0x96,0xb8,0x44,0x45,0x53,0x54,0x00,0x00}}},
	{_T("FlashProp"), {0x1171a62f,0x05d2,0x11d1,{0x83,0xfc,0x00,0xa0,0xc9,0x08,0x9c,0x5a}}},
	{_T("IFlashFactory"), {0xd27cdb70,0xae6d,0x11cf,{0x96,0xb8,0x44,0x45,0x53,0x54,0x00,0x00}}},
	{_T("IDispatchEx"), {0xa6ef9860,0xc720,0x11d0,{0x93,0x37,0x00,0xa0,0xc9,0x0d,0xca,0xa9}}},
	{_T("IFlashObjectInterface"), {0xd27cdb72,0xae6d,0x11cf,{0x96,0xb8,0x44,0x45,0x53,0x54,0x00,0x00}}},
	{_T("IServiceProvider"), {0x6d5140c1,0x7436,0x11ce,{0x80,0x34,0x00,0xaa,0x00,0x60,0x09,0xfa}}},
	{_T("FlashObjectInterface"), {0xd27cdb71,0xae6d,0x11cf,{0x96,0xb8,0x44,0x45,0x53,0x54,0x00,0x00}}},

	// Flash9.ocx\2
	{_T("IFlashAccessibility"), {0x57a0e747, 0x3863, 0x4d20, {0xa8, 0x11, 0x95, 0x0c, 0x84, 0xf1, 0xdb, 0x9b}}},
	{_T("ISimpleTextSelection"), {0x307f64c0, 0x621d, 0x4d56, {0xbb, 0xc6, 0x91, 0xef, 0xc1, 0x3c, 0xe4, 0x0d}}},

	// FlashUtil9d.ocx
	{_T("IFlashBroker"), {0x2e4bb6be, 0xa75f, 0x4dc0, {0x95, 0x00, 0x68, 0x20, 0x36, 0x55, 0xa2, 0xc4}}},

	// mshtml.tlb
	{_T("IHTMLObjectElement"), {0x3050f24f, 0x98b5, 0x11cf, {0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b}}},
	{_T("IHTMLObjectElement2"), {0x3050f4cd, 0x98b5, 0x11cf, {0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b}}},
	{_T("IHTMLObjectElement3"), {0x3050f827, 0x98b5, 0x11cf, {0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b}}},
	{_T("IHTMLEventObj"), {0x3050f32d, 0x98b5, 0x11cf, {0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b}}},
};

static const GUID_HELPER EVENT_IIDS[] = 
{
	// flash9d.ocx
	{_T("DIID__IShockwaveFlashEvents"), {0xd27cdb6d, 0xae6d, 0x11cf, {0x96, 0xb8, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00}}},
	// mshtml.tlb
	{_T("DIID_DWebBrowserEvents2"), {0x34a715a0, 0x6587, 0x11d0, {0x92, 0x4a, 0x00, 0x20, 0xaf, 0xc7, 0xac, 0x4d}}},
};

CIESeeker* CIESeeker::_instance = NULL;

void CIESeeker::Start()
{
	REDIR_OUT2LOG;
	//EnumIE1();
	EnumIE2();
	REDIR_OUT2OUT;
}

void CIESeeker::Stop()
{
	REDIR_OUT2LOG;
	HRESULT hr = S_FALSE;
	RemoveAllSinks();
	REDIR_OUT2OUT;
}

void CIESeeker::AddSink(EVENT_SINK eventSink)
{
	m_eventSinks.insert(EVENT_SINK_PAIR(eventSink.pUnkSink, eventSink));
	//m_eventSinks[eventSink.pUnkSink] = eventSink;
}

void CIESeeker::RemoveAllSinks()
{
	EVENT_SINK_MAP::iterator it;
	for(it = m_eventSinks.begin(); it != m_eventSinks.end(); ++it)
	{
		RemoveSink((*it).first, (*it).second);
	}
	m_eventSinks.clear();
}

void CIESeeker::RemoveSink(LPVOID pSink, EVENT_SINK eventSink)
{
	HRESULT hr = S_FALSE;
	if(eventSink.dwCookie > 0 && eventSink.pUnkSink)
	{
		hr = AtlUnadvise(eventSink.pUnkSrc, eventSink.iid, eventSink.dwCookie);
		// hr = AfxConnectionUnadvise(eventSink.pUnkSrc, eventSink.iid, eventSink.pUnkSink, FALSE, eventSink.dwCookie);
		if (SUCCEEDED(hr))
		{
			cout << _T("Unadvised[") << GetIIDName(eventSink.iid) << _T("], dwCookie = ") << eventSink.dwCookie << endl;
		}	
	}
}

char* GetIIDName(IID iid)
{
	for(int i = 0; i < sizeof(EVENT_IIDS)/sizeof(GUID_HELPER); i++)
	{
		if(IsEqualGUID(EVENT_IIDS[i].iid, iid)) return EVENT_IIDS[i].name;
	}
	return "IUnknown";
}

void Log(LPCTSTR psz)
{
	if(g_pMainWin) g_pMainWin->Log(psz);
	else cout << psz << endl;
}

void EnumIE1()
{
	cout << _T("开始扫描系统中正在运行的浏览器实例... ...") << endl;

	CComPtr< IShellWindows > spShellWin;
	HRESULT hr = spShellWin.CoCreateInstance(CLSID_ShellWindows);
	if (FAILED(hr))
	{
		cout << _T("错误: 获取IShellWindows接口!") << endl;
		return;
	}

	long nCount = 0;		// 取得浏览器实例个数(Explorer 和 IExplorer)
	spShellWin->get_Count(&nCount);
	if(0 == nCount)
	{
		cout << _T("信息: 没有找到运行中的浏览器!") << endl;
		return;
	}

	for(int i = 0; i < nCount; i++)
	{
		CComPtr<IDispatch> spDispIE;
		hr = spShellWin->Item(CComVariant((long)i), &spDispIE);
		if (FAILED (hr)) continue;

		CComQIPtr<IWebBrowser2> spBrowser = spDispIE;
		if (!spBrowser) continue;

		CComPtr<IDispatch> spDispDoc;
		hr = spBrowser->get_Document(&spDispDoc);
		if (FAILED(hr) || !spDispDoc) continue;

		CComQIPtr<IHTMLDocument2> spDocument = spDispDoc;
		if (!spDocument) continue;
		// 程序运行到此，已经找到了 IHTMLDocument2 的接口指针

		// 删除下行语句的注释，把浏览器的背景改变看看
		// spDocument2->put_bgColor(CComVariant("green"));

		EnumDocument(spDocument);
	}
}

BOOL CALLBACK EnumIEWin(HWND hWnd, LPARAM lParam)
{
	TCHAR szClassName[256];
	::GetClassName(hWnd,  szClassName,  sizeof(szClassName));

	if (_tcscmp(szClassName,  _T("Internet Explorer_Server") ) == 0)
	{
		UINT nMsg = ::RegisterWindowMessage(_T("WM_HTML_GETOBJECT"));
		LRESULT lRes;
		::SendMessageTimeout(hWnd, nMsg, 0L, 0L, SMTO_ABORTIFHUNG, 1000, (DWORD*)&lRes);

		CComPtr<IHTMLDocument2> spDispDoc = NULL;
		HRESULT hr = ::ObjectFromLresult (lRes, IID_IHTMLDocument2, 0 , (LPVOID*)&spDispDoc);
		if (FAILED(hr) || !spDispDoc) return TRUE;
		// 程序运行到此，已经找到了 IHTMLDocument2 的接口指针

		EnumDocument(spDispDoc, hWnd);
	}
	return TRUE;		// 继续枚举子窗口
};

void EnumIE2() 
{
	::EnumChildWindows(GetDesktopWindow(), EnumIEWin, NULL);
}

IWebBrowser2* GetWebBrowser(IHTMLDocument2 *pIHTMLDocument2)
{
	if(!pIHTMLDocument2) return NULL;

	HRESULT hr = S_OK;
	IWebBrowser2 *pWebBrowser2 = NULL;
	CComPtr<::IServiceProvider> spServiceProvider = NULL;

	hr = pIHTMLDocument2->QueryInterface(::IID_IServiceProvider, (void**)&spServiceProvider);
	if(SUCCEEDED(hr) && spServiceProvider)
	{
		hr = spServiceProvider->QueryService(SID_SInternetExplorer, IID_IWebBrowser2, (void**)&pWebBrowser2);
	}

	return pWebBrowser2;
}

HWND GetWindowHandle(IWebBrowser2 *pWebBrowser2)
{
	if(!pWebBrowser2) return NULL;
	HWND hWnd = NULL;
	pWebBrowser2->get_HWND((long*)&hWnd);
	return hWnd;
}

HWND GetWindowHandle(IHTMLDocument2 *pIHTMLDocument2)
{
	if(!pIHTMLDocument2) return NULL;
	return GetWindowHandle(GetWebBrowser(pIHTMLDocument2));
}

void EnumDocument(IHTMLDocument2 *pIHTMLDocument2, HWND hWnd)
{
	HRESULT hr;
	USES_CONVERSION;

	if(hWnd == NULL) hWnd = GetWindowHandle(pIHTMLDocument2);
	if (hWnd != NULL)
	{
		HWND hParentWnd = ::GetAncestor(hWnd, GA_ROOTOWNER);
		TCHAR lpCaption[MAX_PATH] = {0};
		::GetWindowText(hParentWnd, lpCaption, MAX_PATH);
		cout << _T("============") << hParentWnd << lpCaption << _T("============") << endl;
	}
	if (!pIHTMLDocument2) return;

	CComBSTR bstrTitle, bstrURL, bstrHTML;
	pIHTMLDocument2->get_title(&bstrTitle);	//取得文档标题
	pIHTMLDocument2->get_URL(&bstrURL);

	CComQIPtr <IHTMLElement> spBody;
	hr = pIHTMLDocument2->get_body(&spBody);
	if (FAILED(hr)) return;
	
	CComQIPtr <IHTMLElement> spHTML;
	hr = spBody->get_parentElement(&spHTML);
	if (FAILED(hr)) return;
	spHTML->get_outerHTML(&bstrHTML);
	
	cout << _T("Title: ") << (bstrTitle ? OLE2CT(bstrTitle) : _T("")) << endl;
	cout << _T("URL: ") << (bstrURL ? OLE2CT(bstrURL) : _T("")) << endl;
	// cout << _T("HTML: ") << (bstrHTML ? OLE2CT(bstrHTML) : _T("")) << endl;
	EnumForms(pIHTMLDocument2);
	EnumObjects(pIHTMLDocument2);
	EnumEmbeds(pIHTMLDocument2);
	EnumFrames(pIHTMLDocument2);
	IWebBrowser2 *spWebBrowser2 = GetWebBrowser(pIHTMLDocument2);
	//CreateThread(NULL, 0, MonitorWebBrowser, this, 0, 0);
	MonitorWebBrowser(spWebBrowser2);
}

void EnumFrames(IHTMLDocument2 *pIHTMLDocument2)
{
	if (!pIHTMLDocument2) return;

	HRESULT hr;
	CComPtr< IHTMLFramesCollection2 > spFramesCollection2;
	hr = pIHTMLDocument2->get_frames(&spFramesCollection2);	//取得框架frame的集合
	if (FAILED(hr))
	{
		cout << _T("错误: 获取Frame集合IHTMLFramesCollection2!") << endl;
		return;
	}

	long nFrameCount = 0;				//取得子框架个数
	hr = spFramesCollection2->get_length(&nFrameCount);
	if (FAILED(hr))
	{
		cout << _T("错误: 获取Frame集合IHTMLFramesCollection2的长度!") << endl;
	}

	cout << "Frame对象共: " << nFrameCount << endl;

	for(long i = 0; i < nFrameCount; i++)
	{
		CComVariant vDispWin2;		//取得子框架的自动化接口
		hr = spFramesCollection2->item(&CComVariant(i), &vDispWin2);
		if (FAILED(hr))	continue;

		CComQIPtr<IHTMLWindow2> spWin2 = vDispWin2.pdispVal;
		if(!spWin2)	continue;	//取得子框架的IHTMLWindow2接口

		CComQIPtr <IHTMLDocument2> spDoc2;
		spWin2->get_document(&spDoc2);	//取得字框架的IHTMLDocument2接口

		cout << _T("************ Frame: ") << i + 1 << _T(" ************") << endl;
		EnumDocument(spDoc2);			//递归枚举当前子框架的IHTMLDocument2
	}
}

void EnumEmbeds(IHTMLDocument2 *pIHTMLDocument2)
{
	if(!pIHTMLDocument2) return;

	HRESULT hr;
	USES_CONVERSION;

	CComPtr< IHTMLElementCollection > spEmbedCollection;
	hr = pIHTMLDocument2->get_embeds(&spEmbedCollection);	//取得Embed的集合
	if (FAILED(hr))
	{
		cout << _T("错误: 获取Embed集合IHTMLElementCollection!") << endl;
		return;
	}

	long nEmbedCount = 0;				//取得Embed个数
	hr = spEmbedCollection->get_length(&nEmbedCount);
	if (FAILED(hr))
	{
		cout << _T("错误: 获取Embed集合IHTMLElementCollection的长度!") << endl;
	}

	cout << "Embed对象共: " << nEmbedCount << endl;

	CComBSTR bstrPluginsPage, bstrSrc;
	IDispatch *pDisp = NULL;
	CComQIPtr< IHTMLEmbedElement > spEmbedElement;
	for(long i = 0; i < nEmbedCount; i++)
	{
		hr = spEmbedCollection->item(CComVariant(i), CComVariant(), &pDisp);
		if (FAILED(hr) || !pDisp) continue;

		spEmbedElement = pDisp;
		pDisp->Release();
		
		spEmbedElement->get_pluginspage(&bstrPluginsPage);
		spEmbedElement->get_src(&bstrSrc);
		cout << i + 1
			<< ": pluginspage = " << (bstrPluginsPage ? OLE2CT(bstrPluginsPage) : _T(""))
			<< ", src = " << (bstrSrc ? OLE2CT(bstrSrc) : _T("")) 
			<< endl;
	}
	if (pDisp) 
	{
		pDisp->Release();
		pDisp = NULL;
	}
}

void EnumObjects(IHTMLDocument2 *pIHTMLDocument2)
{
	if(!pIHTMLDocument2) return;

	HRESULT hr;
	USES_CONVERSION;

	CComPtr< IHTMLElementCollection > spAllCollection;
	hr = pIHTMLDocument2->get_all(&spAllCollection);	//取得document.all的集合
	if (FAILED(hr))
	{
		cout << _T("错误: 获取document.all集合IHTMLElementCollection!") << endl;
		return;
	}

	IDispatch *pDisp = NULL;
	hr = spAllCollection->tags(CComVariant(_T("OBJECT")), &pDisp);
	if (FAILED(hr))
	{
		cout << _T("错误: 获取document.all.tags('OBJECT')集合IHTMLElementCollection!") << endl;
		return;
	}

	CComQIPtr <IHTMLElementCollection> spObjectCollection = pDisp;
	long nObjectCount = 0;				//取得Embed个数
	hr = spObjectCollection->get_length(&nObjectCount);
	if (FAILED(hr))
	{
		cout << _T("错误: 获取Object集合IHTMLElementCollection的长度!") << endl;
	}

	cout << "Object对象共: " << nObjectCount << endl;

	CComBSTR bstrClassId, bstrData, bstrName, bstrCodeBase;
	CComBSTR bstrCodeType, bstrCode, bstrBaseHref, bstrType;

	CComQIPtr< IHTMLObjectElement > spObjectElement;

	for(long i = 0; i < nObjectCount; i++)
	{
		hr = spObjectCollection->item(CComVariant(i), CComVariant(), &pDisp);
		if (FAILED(hr)) continue;
		EnumInterface(pDisp, _T("OBJECT"));
		EnumConnectionPoint(pDisp, _T("OBJECT"));
		spObjectElement = pDisp;
		
		spObjectElement->get_classid(&bstrClassId);
		spObjectElement->get_data(&bstrData);
		spObjectElement->get_name(&bstrName);
		spObjectElement->get_codeBase(&bstrCodeBase);
		spObjectElement->get_codeType(&bstrCodeType);
		spObjectElement->get_code(&bstrCode);
		spObjectElement->get_BaseHref(&bstrBaseHref);
		spObjectElement->get_type(&bstrType);

		cout << i + 1 
			<< ": classid = " << (bstrClassId ? OLE2CT(bstrClassId) : _T(""))
			<< ", data = " << (bstrData ? OLE2CT(bstrData) : _T(""))
			<< ", name = " << (bstrName ? OLE2CT(bstrName) : _T(""))
			<< ", codeBase = " << (bstrCodeBase ? OLE2CT(bstrCodeBase) : _T(""))
			<< ", codeType = " << (bstrCodeType ? OLE2CT(bstrCodeType) : _T(""))
			<< ", code = " << (bstrCode ? OLE2CT(bstrCode) : _T(""))
			<< ", base href = " << (bstrBaseHref ? OLE2CT(bstrBaseHref) : _T(""))
			<< ", type = " << (bstrType ? OLE2CT(bstrType) :_T("")) 
			<< endl;
	
		hr = spObjectElement->get_object(&pDisp);
		if (FAILED(hr) || !pDisp) continue;
		//ViewObject(pDisp);
		EnumInterface(pDisp, _T("FLASH"));
		EnumConnectionPoint(pDisp, _T("FLASH"));
		MonitorFlash(pDisp);
	}
	if (pDisp) 
	{
		pDisp->Release();
		pDisp = NULL;
	}	
}

void EnumInterface(IDispatch *pDisp, LPCSTR comment)
{
	if (!pDisp) return;
	HRESULT hr = S_OK;
	void *pObject;
	if(comment) cout << comment << _T(" ");
	cout << "Interface: [";
	for(int i = 0; i < sizeof(COM_IIDS)/sizeof(GUID_HELPER); i++)
	{
		hr = pDisp->QueryInterface(COM_IIDS[i].iid, &pObject);
		if(SUCCEEDED(hr) && pObject) cout << COM_IIDS[i].name << _T(", ");
	}
	cout << "]" << endl;
}

void EnumConnectionPoint(IDispatch *pDisp, LPCSTR comment)
{
	if (!pDisp) return;
	HRESULT hr = S_OK;

	USES_CONVERSION;
	IConnectionPointContainer *spCPC = NULL;
	IConnectionPoint *spCP = NULL;
	IEnumConnectionPoints *ppEnum = NULL;
	IID id;
	LPOLESTR strGUID = 0;
	CComBSTR bstrId = NULL;

	if(comment) cout << comment << _T(" ");
	cout << _T("ConnectionPoint: [");
	hr = pDisp->QueryInterface(IID_IConnectionPointContainer, reinterpret_cast<void**>(&spCPC));
	if (SUCCEEDED(hr) && spCPC)
	{
		hr = spCPC->EnumConnectionPoints(&ppEnum);
		while(SUCCEEDED(ppEnum->Next(1, &spCP, NULL)) && spCP) 
		{
			spCP->GetConnectionInterface(&id);
			// 注意此处StringFromIID的正确用法, 否则会在Debug状态下出错: user breakpoint called from code 0x7c901230
			StringFromIID(id, &strGUID);
			bstrId = strGUID;
			CoTaskMemFree(strGUID);
			cout << (bstrId ? OLE2CT(bstrId) : _T("")) << _T(", ");
		}
	}
	cout << "]" << endl;
	if(spCP) spCP->Release();
	if(ppEnum) ppEnum->Release();
	if(spCPC) spCPC->Release();
/*
	{9bfbbc02-eff1-101a-84ed-00aa00341d07}	IPropertyNotifySink
	{3050f3c4-98b5-11cf-bb82-00aa00bdce0b}	HTMLObjectElementEvents
	{3050f620-98b5-11cf-bb82-00aa00bdce0b}	HTMLObjectElementEvents2
	{00020400-0000-0000-c000-000000000046}	IDispatch
	{1dc9ca50-06ef-11d2-8415-006008c3fbfc}	ITridentEventSink
	{d27cdb6d-ae6d-11cf-96b8-444553540000}	ShockwaveFlashObjects::_IShockwaveFlashEvents
*/
}

void EnumForms(IHTMLDocument2 *pIHTMLDocument2)
{
	if(!pIHTMLDocument2) return;

	HRESULT hr;
	USES_CONVERSION;

	CComQIPtr< IHTMLElementCollection > spElementCollection;
	hr = pIHTMLDocument2->get_forms(&spElementCollection);	//取得表单集合
	if (FAILED(hr))
	{
		cout << _T("错误: 获取Form集合IHTMLElementCollection!") << endl;
		return;
	}

	long nFormCount = 0;				//取得表单数目
	hr = spElementCollection->get_length(&nFormCount);
	if (FAILED(hr))
	{
		cout << _T("错误: 获取Form集合IHTMLElementCollection的长度!") << endl;
		return;
	}
	
	cout << "Form对象共: " << nFormCount << endl;

	for(long i = 0; i < nFormCount; i++)
	{
		IDispatch *pDisp = NULL;		//取得第i个Form表单
		hr = spElementCollection->item(CComVariant(i), CComVariant(), &pDisp);
		if (FAILED(hr)) continue;

		CComQIPtr< IHTMLFormElement > spFormElement = pDisp;
		pDisp->Release();

		long nElemCount = 0;			//取得表单中域的数目
		hr = spFormElement->get_length(&nElemCount);
		if (FAILED(hr))	continue;

		for(long j = 0; j < nElemCount; j++)
		{
			CComDispatchDriver spInputElement;	//取得第 j 项表单域
			hr = spFormElement->item(CComVariant(j), CComVariant(), &spInputElement);
			if (FAILED(hr)) continue;

			CComVariant vName,vVal,vType;		//取得表单域的 名，值，类型
			hr = spInputElement.GetPropertyByName(L"name", &vName);
			if(FAILED(hr)) continue;
			hr = spInputElement.GetPropertyByName(L"value", &vVal);
			if(FAILED(hr)) continue;
			hr = spInputElement.GetPropertyByName(L"type", &vType);
			if(FAILED(hr)) continue;

			cout << _T("[") << (vType.bstrVal? OLE2CT(vType.bstrVal) : _T("")) << _T("] ")
				<< (vName.bstrVal? OLE2CT(vName.bstrVal) : _T(""))
				<< _T(" = ") 
				<< (vVal.bstrVal? OLE2CT(vVal.bstrVal ) : _T(""))
				<< endl;
		}
		//想提交这个表单吗？删除下面语句的注释吧
		//pForm->submit();
	}
}

void MonitorWebBrowser(IWebBrowser2 *pWebBrowser2)
{
	if(!pWebBrowser2) return;

	HRESULT hr = S_FALSE;
	/*
	// 此处也可以用方法局部变量, 用类变量主要是为了Unadvise.
	// 注意编译开关MTd(多线程)
	CComObject<CIESink> *m_ieSink = new CComObject<CIESink>();
	DWORD m_ieCookie = 0;
	*/

	EVENT_SINK es;
	CIESink2 *sink2 = new CIESink2();
	LPDISPATCH pSinkDisp = sink2->GetIDispatch(FALSE);
	hr = AfxConnectionAdvise(pWebBrowser2, DIID_DWebBrowserEvents2, pSinkDisp, FALSE, &es.dwCookie);
	if (hr)
	{
		es.pUnkSink = pSinkDisp;
		es.pUnkSrc = pWebBrowser2;
		es.iid = DIID_DWebBrowserEvents2;
		CIESeeker::Instance()->AddSink(es);
		cout << _T("Advised[DIID_DWebBrowserEvents2], dwCookie = ") << es.dwCookie << endl;
	}

	/*
	EVENT_SINK es;
	CComObject<CIESink> *sink = new CComObject<CIESink>();
	hr = AtlAdvise(pWebBrowser2, (DWebBrowserEvents2*)sink, DIID_DWebBrowserEvents2, &es.dwCookie);
	if (SUCCEEDED(hr))
	{
		es.pUnkSink = (LPUNKNOWN)sink;
		es.pUnkSrc = pWebBrowser2;
		es.iid = DIID_DWebBrowserEvents2;
		CIESeeker::Instance()->AddSink(es);
		cout << _T("Advised[DIID_DWebBrowserEvents2], dwCookie = ") << es.dwCookie << endl;
	}
	*/

	/*
	IConnectionPointContainer *spCPC = NULL;
	IConnectionPoint *spCP = NULL;

	hr = pWebBrowser2->QueryInterface(IID_IConnectionPointContainer, reinterpret_cast<void**>(&spCPC));
	if(FAILED(hr) || !spCPC) return;

	hr = spCPC->FindConnectionPoint(DIID_DWebBrowserEvents2, &spCP);
	if (SUCCEEDED(hr) && spCP) 
	{
		cout << _T("Found IWebBrowser2>>IConnectionPointContainer>>IConnectionPoint[DIID_DWebBrowserEvents2]!") << endl;
		hr = spCP->Advise((IDispatch*)m_ieSink, &m_ieCookie);
		if (SUCCEEDED(hr))
		{
			cout << _T("Advised[DIID_DWebBrowserEvents2], dwCookie = ") << m_ieCookie << endl;
		}
	}
	spCP->Release();
	spCPC->Release();
	*/
}


void MonitorFlash(IDispatch *pDisp)
{
	if (!pDisp) return;
	HRESULT hr = S_OK;

	// 使用CFlashSink3, 连接点无法成功Advise, 报错: E_POINTER 0x80004003(无效指针)
	/*
	EVENT_SINK es;
	CFlashSink3* sink3 = new CFlashSink3();
	IDispatch* pSinkDisp = NULL;
	hr = sink3->QueryInterface(DIID__IShockwaveFlashEvents, reinterpret_cast<void**>(&pSinkDisp));
	hr = AtlAdvise(pDisp, sink3, DIID__IShockwaveFlashEvents, &es.dwCookie);
	if (SUCCEEDED(hr))
	{
		es.pUnkSink = pSinkDisp;
		es.pUnkSrc = pDisp;
		es.iid = DIID__IShockwaveFlashEvents;
		CIESeeker::Instance()->AddSink(es);
		cout << _T("Advised[DIID__IShockwaveFlashEvents], dwCookie = ") << es.dwCookie << endl;
	}
	*/

	// 使用CFlashSink2, 连接点无法成功Advise, dwCookie = 0. 报错: CONNECT_E_CANNOTCONNECT 0x80040202(无法连接).
	// IShockwaveFlashEvents不支持CCmdTarget方式的连接点
	/*
	EVENT_SINK es;
	CFlashSink2* sink2 = new CFlashSink2();
	IDispatch* pSinkDisp = sink2->GetIDispatch(FALSE);
	//hr = AfxConnectionAdvise(pDisp, DIID__IShockwaveFlashEvents, pSinkDisp, FALSE, &es.dwCookie);
	hr = AtlAdvise(pDisp, pSinkDisp, DIID__IShockwaveFlashEvents, &es.dwCookie);
	if (hr)
	{
		es.pUnkSink = pSinkDisp;
		es.pUnkSrc = pDisp;
		es.iid = DIID__IShockwaveFlashEvents;
		CIESeeker::Instance()->AddSink(es);
		cout << _T("Advised[DIID__IShockwaveFlashEvents], dwCookie = ") << es.dwCookie << endl;
	}
	*/

	EVENT_SINK es;
	CComObject<CFlashSink> *sink = new CComObject<CFlashSink>;
	hr = AtlAdvise(pDisp, sink, DIID__IShockwaveFlashEvents, &es.dwCookie);
	if (SUCCEEDED(hr))
	{
		es.pUnkSink = sink;
		es.pUnkSrc = pDisp;
		es.iid = DIID__IShockwaveFlashEvents;
		CIESeeker::Instance()->AddSink(es);
		cout << _T("Advised[DIID__IShockwaveFlashEvents], dwCookie = ") << es.dwCookie << endl;
	}

	/*
	IConnectionPointContainer *spCPC = NULL;
	IConnectionPoint *spCP = NULL;
	//CFlashSunk *sink = new CFlashSunk();
	//CComObject<CFlashSunk> *sink = new CComObject<CFlashSunk>();
	//DWORD dwCookie = 0;

	hr = pDisp->QueryInterface(IID_IConnectionPointContainer, reinterpret_cast<void**>(&spCPC));
	if (SUCCEEDED(hr) && spCPC)
	{
		cout << _T("Found IHTMLObjectElement>>IConnectionPointContainer!") << endl;

		hr = spCPC->FindConnectionPoint(DIID__IShockwaveFlashEvents, &spCP);
		if (SUCCEEDED(hr) && spCP) 
		{
			cout << _T("Found IHTMLObjectElement>>IConnectionPointContainer>>IConnectionPoint[DIID__IShockwaveFlashEvents]!") << endl;
			
			hr = spCP->Advise(m_flashSink, &m_flashCookie);
			if (SUCCEEDED(hr))
			{
				cout << _T("Advised[DIID__IShockwaveFlashEvents], dwCookie = ") << m_flashCookie << endl;
			}
		}
	}
	*/
	//if(spCP) spCP->Release();
	//if(spCPC) spCPC->Release();
}
}