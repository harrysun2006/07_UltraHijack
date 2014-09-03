#include "stdafx.h"
#include "IESink.h"
#include "Spyer.h"
#include <exdispid.h>

using namespace spyer;

BEGIN_DISPATCH_MAP(CIESink, CCmdTarget)
	DISP_FUNCTION_ID(CIESink, "BeforeNavigate2", DISPID_BEFORENAVIGATE2, OnBeforeNavigate2, VT_EMPTY, VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PBOOL)
	DISP_FUNCTION_ID(CIESink, "NavigateComplete2", DISPID_NAVIGATECOMPLETE2, OnNavigateComplete2, VT_EMPTY, VTS_DISPATCH VTS_PVARIANT)
	DISP_FUNCTION_ID(CIESink, "StatusTextChange", DISPID_STATUSTEXTCHANGE, OnStatusTextChange, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(CIESink, "ProgressChange", DISPID_PROGRESSCHANGE, OnProgressChange, VT_EMPTY, VTS_I4 VTS_I4)
	DISP_FUNCTION_ID(CIESink, "DocumentComplete", DISPID_DOCUMENTCOMPLETE, OnDocumentComplete, VT_EMPTY, VTS_DISPATCH VTS_PVARIANT)
	DISP_FUNCTION_ID(CIESink, "DownloadBegin", DISPID_DOWNLOADBEGIN, OnDownloadBegin, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(CIESink, "DownloadComplete", DISPID_DOWNLOADCOMPLETE, OnDownloadComplete, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(CIESink, "CommandStateChange", DISPID_COMMANDSTATECHANGE, OnCommandStateChange, VT_EMPTY, VTS_I4 VTS_BOOL)
	DISP_FUNCTION_ID(CIESink, "NewWindow2", DISPID_NEWWINDOW2, OnNewWindow2, VT_EMPTY, VTS_PDISPATCH VTS_PBOOL)
	DISP_FUNCTION_ID(CIESink, "NewWindow3", DISPID_NEWWINDOW3, OnNewWindow3, VT_EMPTY, VTS_PDISPATCH VTS_PBOOL VTS_I4 VTS_BSTR VTS_BSTR)
	DISP_FUNCTION_ID(CIESink, "TitleChange", DISPID_TITLECHANGE, OnTitleChange, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(CIESink, "PropertyChange", DISPID_PROPERTYCHANGE, OnPropertyChange, VT_EMPTY, VTS_BSTR)
	DISP_FUNCTION_ID(CIESink, "Quit", DISPID_QUIT, DoQuit, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(CIESink, "OnQuit", DISPID_ONQUIT, DoQuit, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(CIESink, "OnVisible", DISPID_ONVISIBLE, OnVisible, VT_EMPTY, VTS_BOOL)
	DISP_FUNCTION_ID(CIESink, "OnToolBar", DISPID_ONTOOLBAR, OnToolBar, VT_EMPTY, VTS_BOOL)
	DISP_FUNCTION_ID(CIESink, "OnMenuBar", DISPID_ONMENUBAR, OnMenuBar, VT_EMPTY, VTS_BOOL)
	DISP_FUNCTION_ID(CIESink, "OnStatusBar", DISPID_ONSTATUSBAR, OnStatusBar, VT_EMPTY, VTS_BOOL)
	DISP_FUNCTION_ID(CIESink, "OnFullScreen", DISPID_ONFULLSCREEN, OnFullScreen, VT_EMPTY, VTS_BOOL)
	DISP_FUNCTION_ID(CIESink, "OnTheaterMode", DISPID_ONTHEATERMODE, OnTheaterMode, VT_EMPTY, VTS_BOOL)
	DISP_FUNCTION_ID(CIESink, "WindowSetResizable", DISPID_WINDOWSETRESIZABLE, OnWindowSetResizable, VT_EMPTY, VTS_BOOL)
	DISP_FUNCTION_ID(CIESink, "WindowSetLeft", DISPID_WINDOWSETLEFT, OnWindowSetLeft, VT_EMPTY, VTS_I4)
	DISP_FUNCTION_ID(CIESink, "WindowSetTop", DISPID_WINDOWSETTOP, OnWindowSetTop, VT_EMPTY, VTS_I4)
	DISP_FUNCTION_ID(CIESink, "WindowSetWidth", DISPID_WINDOWSETWIDTH, OnWindowSetWidth, VT_EMPTY, VTS_I4)
	DISP_FUNCTION_ID(CIESink, "WindowSetHeight", DISPID_WINDOWSETHEIGHT, OnWindowSetHeight, VT_EMPTY, VTS_I4)
	DISP_FUNCTION_ID(CIESink, "WindowClosing", DISPID_WINDOWCLOSING, OnWindowClosing, VT_EMPTY, VTS_BOOL VTS_PBOOL)
	DISP_FUNCTION_ID(CIESink, "ClientToHostWindow", DISPID_CLIENTTOHOSTWINDOW, OnClientToHostWindow, VT_EMPTY, VTS_PI4 VTS_PI4)
	DISP_FUNCTION_ID(CIESink, "SetSecureLockIcon", DISPID_SETSECURELOCKICON, OnSetSecureLockIcon, VT_EMPTY, VTS_I4)
	DISP_FUNCTION_ID(CIESink, "FileDownload", DISPID_FILEDOWNLOAD, OnFileDownload, VT_EMPTY, VTS_PBOOL)
	DISP_FUNCTION_ID(CIESink, "NavigateError", DISPID_NAVIGATEERROR, OnNavigateError, VT_EMPTY, VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PBOOL)
	DISP_FUNCTION_ID(CIESink, "PrintTemplateInstantiation", DISPID_PRINTTEMPLATEINSTANTIATION, OnPrintTemplateInstantiation, VT_EMPTY, VTS_DISPATCH)
	DISP_FUNCTION_ID(CIESink, "PrintTemplateTeardown", DISPID_PRINTTEMPLATETEARDOWN, OnPrintTemplateTeardown, VT_EMPTY, VTS_DISPATCH)
	DISP_FUNCTION_ID(CIESink, "UpdatePageStatus", DISPID_UPDATEPAGESTATUS, OnUpdatePageStatus, VT_EMPTY, VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT)
	DISP_FUNCTION_ID(CIESink, "PrivacyImpactedStateChange", DISPID_PRIVACYIMPACTEDSTATECHANGE, OnPrivacyImpactedStateChange, VT_EMPTY, VTS_BOOL)
END_DISPATCH_MAP()

// DWebBrowserEvents2 methods
// Fired before navigate occurs in the given WebBrowser (window or frameset element). 
// The processing of this navigation may be modified.
void CIESink::OnBeforeNavigate2(LPDISPATCH pDisp,
			 				   VARIANT* URL,
							   VARIANT* Flags,
							   VARIANT* TargetFrameName,
							   VARIANT* PostData,
							   VARIANT* Headers,
							   VARIANT_BOOL* Cancel)
{
}

// Fired when the document being navigated to becomes visible and enters the navigation stack.
void CIESink::OnNavigateComplete2(LPDISPATCH pDisp, VARIANT* URL)
{
}

// Statusbar text changed.
void CIESink::OnStatusTextChange(BSTR Text)
{
}

// Fired when download progress is updated.
void CIESink::OnProgressChange(long Progress, long ProgressMax)
{
}

// Fired when the document being navigated to reaches ReadyState_Complete.
void CIESink::OnDocumentComplete(LPDISPATCH pDisp, VARIANT *pvarURL)
{
}

// Download of a page started.
void CIESink::OnDownloadBegin()
{
}

// Download of page complete.
void CIESink::OnDownloadComplete()
{
}

// The enabled state of a command changed.
void CIESink::OnCommandStateChange(long Command, BOOL Enable)
{
}

// A new, hidden, non-navigated WebBrowser window is needed.
void CIESink::OnNewWindow2(LPDISPATCH* ppDisp, VARIANT_BOOL* Cancel)
{
}

// A new, hidden, non-navigated WebBrowser window is needed.
void CIESink::OnNewWindow3(LPDISPATCH* ppDisp, VARIANT_BOOL* Cancel, unsigned long dwFlags, 
	BSTR bstrUrlContext, BSTR bstrUrl)
{
}

// Document title changed.
void CIESink::OnTitleChange(BSTR Text)
{
}

// Fired when the PutProperty method has been called.
void CIESink::OnPropertyChange(BSTR szProperty)
{
}

void CIESink::DoQuit()
{
	OnQuit();
	CSpyer::Instance()->RemoveSink(this);
}

// Fired when application is quiting.
void CIESink::OnQuit()
{
}

// Fired when the window should be shown/hidden.
void CIESink::OnVisible(VARIANT_BOOL Visible)
{
}

// Fired when the toolbar  should be shown/hidden.
void CIESink::OnToolBar(VARIANT_BOOL ToolBar)
{
}

// Fired when the menubar should be shown/hidden.
void CIESink::OnMenuBar(VARIANT_BOOL MenuBar)
{
}

// Fired when the statusbar should be shown/hidden.
void CIESink::OnStatusBar(VARIANT_BOOL StatusBar)
{
}

// Fired when fullscreen mode should be on/off.
void CIESink::OnFullScreen(VARIANT_BOOL FullScreen)
{
}

// Fired when theater mode should be on/off.
void CIESink::OnTheaterMode(VARIANT_BOOL TheaterMode)
{
}

// Fired when the host window should allow/disallow resizing.
void CIESink::OnWindowSetResizable(VARIANT_BOOL Resizable)
{
}

// Fired when the host window should change its Left coordinate.
void CIESink::OnWindowSetLeft(long Left)
{
}

// Fired when the host window should change its Top coordinate.
void CIESink::OnWindowSetTop(long Top)
{
}

// Fired when the host window should change its width.
void CIESink::OnWindowSetWidth(long Width)
{
}

// Fired when the host window should change its height.
void CIESink::OnWindowSetHeight(long Height)
{
}

// Fired when the WebBrowser is about to be closed by script.
void CIESink::OnWindowClosing(VARIANT_BOOL IsChildWindow, VARIANT_BOOL* Cancel)
{
}

// Fired to request client sizes be converted to host window sizes.
void CIESink::OnClientToHostWindow(long* CX, long* CY)
{
}

// Fired to indicate the security level of the current web page contents.
void CIESink::OnSetSecureLockIcon(long SecureLockIcon)
{
}

// Fired to indicate the File Download dialog is opening.
void CIESink::OnFileDownload(VARIANT_BOOL* Cancel)
{
}

// Fired when a binding error occurs (window or frameset element).
void CIESink::OnNavigateError(LPDISPATCH pDisp, VARIANT* URL, VARIANT* Frame, 
	VARIANT* StatusCode, VARIANT_BOOL* Cancel)
{
}

// Fired when a print template is instantiated.
void CIESink::OnPrintTemplateInstantiation(LPDISPATCH pDisp)
{
}

// Fired when a print template destroyed."
void CIESink::OnPrintTemplateTeardown(LPDISPATCH pDisp)
{
}

// Fired when a page is spooled. When it is fired can be changed by a custom template.
void CIESink::OnUpdatePageStatus(LPDISPATCH pDisp, VARIANT* nPage, VARIANT* fDone)
{
}

// Fired when the global privacy impacted state changes.
void CIESink::OnPrivacyImpactedStateChange(VARIANT_BOOL bImpacted)
{
}
