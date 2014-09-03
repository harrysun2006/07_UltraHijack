// IESink.h: interface for the CIESink class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(IESINK__INCLUDED_)
#define IESINK__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"

#define DISPID_PRINTTEMPLATEINSTANTIATION	225
#define DISPID_PRINTTEMPLATETEARDOWN		226
#define DISPID_UPDATEPAGESTATUS				227

#define DISPID_WINDOWSETRESIZABLE			262
#define DISPID_WINDOWCLOSING				263
#define DISPID_WINDOWSETLEFT				264
#define DISPID_WINDOWSETTOP					265
#define DISPID_WINDOWSETWIDTH				266
#define DISPID_WINDOWSETHEIGHT				267
#define DISPID_CLIENTTOHOSTWINDOW			268
#define DISPID_SETSECURELOCKICON			269
#define DISPID_FILEDOWNLOAD					270
#define DISPID_NAVIGATEERROR				271
#define DISPID_PRIVACYIMPACTEDSTATECHANGE	272
#define DISPID_NEWWINDOW3					273

class CIESink : public CCmdTarget
{
public:
	CIESink() {this->EnableAutomation();}
protected:
	virtual void OnBeforeNavigate2(LPDISPATCH, VARIANT*, VARIANT*, VARIANT*, VARIANT* PostData, VARIANT* Headers, VARIANT_BOOL* Cancel);
	virtual void OnNavigateComplete2(LPDISPATCH pDisp, VARIANT* URL);
	virtual void OnStatusTextChange(BSTR Text);
	virtual void OnProgressChange(long Progress, long ProgressMax);
	virtual void OnDocumentComplete(IDispatch *pDisp, VARIANT *pvarURL);
	virtual void OnDownloadBegin();
	virtual void OnDownloadComplete();
	virtual void OnCommandStateChange(long Command, BOOL Enable);
	virtual void OnNewWindow2(LPDISPATCH* ppDisp, VARIANT_BOOL* Cancel);
	virtual void OnNewWindow3(LPDISPATCH* ppDisp, VARIANT_BOOL* Cancel, 
		unsigned long dwFlags, BSTR bstrUrlContext, BSTR bstrUrl);
	virtual void OnTitleChange(BSTR Text);
	virtual void OnPropertyChange(BSTR szProperty);
	virtual void OnQuit();
	virtual void OnVisible(VARIANT_BOOL Visible);
	virtual void OnToolBar(VARIANT_BOOL ToolBar);
	virtual void OnMenuBar(VARIANT_BOOL MenuBar);
	virtual void OnStatusBar(VARIANT_BOOL StatusBar);
	virtual void OnFullScreen(VARIANT_BOOL FullScreen);
	virtual void OnTheaterMode(VARIANT_BOOL TheaterMode);
	virtual void OnWindowSetResizable(VARIANT_BOOL Resizable);
	virtual void OnWindowSetLeft(long Left);
	virtual void OnWindowSetTop(long Top);
	virtual void OnWindowSetWidth(long Width);
	virtual void OnWindowSetHeight(long Height);
	virtual void OnWindowClosing(VARIANT_BOOL IsChildWindow, VARIANT_BOOL* Cancel);
	virtual void OnClientToHostWindow(long* CX, long* CY);
	virtual void OnSetSecureLockIcon(long SecureLockIcon);
	virtual void OnFileDownload(VARIANT_BOOL* Cancel);
	virtual void OnNavigateError(LPDISPATCH pDisp, VARIANT* URL, VARIANT* Frame,
		VARIANT* StatusCode, VARIANT_BOOL* Cancel);
	virtual void OnPrintTemplateInstantiation(LPDISPATCH pDisp);
	virtual void OnPrintTemplateTeardown(LPDISPATCH pDisp);
	virtual void OnUpdatePageStatus(LPDISPATCH pDisp, VARIANT* nPage, VARIANT* fDone);
	virtual void OnPrivacyImpactedStateChange(VARIANT_BOOL bImpacted);
private:
	void DoQuit();
DECLARE_DISPATCH_MAP()
};

#endif // !defined(IESINK__INCLUDED_)