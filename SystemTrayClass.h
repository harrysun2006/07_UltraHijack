#define WM_TRAYNOTIFY 0xA44C

#include <Shellapi.h>

class clsSysTray
{
public:
	clsSysTray();
	
	BOOL SetIcon(HICON hNewIcon);
	HICON GetIcon();
	
	BOOL SetTipText(char *lpstrNewTipText);
	char *GetTipText();
	
	BOOL AddIcon();
	BOOL RemoveIcon();
	
	HWND hWnd;
	
protected:
	NOTIFYICONDATA NotifyIconData;
	bool bInTray;
	
};
clsSysTray::clsSysTray()
{
	bInTray = false;
	
	NotifyIconData.cbSize = sizeof(NotifyIconData);
	NotifyIconData.uID = 2;
	NotifyIconData.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE;
	NotifyIconData.uCallbackMessage = WM_TRAYNOTIFY;
	NotifyIconData.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	NotifyIconData.szTip[0] = '\0';
	NotifyIconData.hWnd = NULL;
	
}
HICON clsSysTray::GetIcon()
{
	return(NotifyIconData.hIcon);
}
BOOL clsSysTray::SetIcon(HICON hNewIcon)
{
	NotifyIconData.hIcon = hNewIcon;
	
	if(bInTray)
	{
		BOOL iRetVal;
		
		iRetVal = Shell_NotifyIcon(NIM_MODIFY, &NotifyIconData);
		
		if(iRetVal)
		{
			bInTray = true;
		}
		
		return(iRetVal);
	}
	else
		return(1);
}
char *clsSysTray::GetTipText()
{
	return(NotifyIconData.szTip);
}
BOOL clsSysTray::SetTipText(char *lpstrNewTipText)
{
	strcpy(NotifyIconData.szTip, lpstrNewTipText);
	
	if(bInTray)
	{
		BOOL iRetVal;
		
		iRetVal = Shell_NotifyIcon(NIM_MODIFY, &NotifyIconData);
		
		if(iRetVal)
		{
			bInTray = true;
		}
		
		return(iRetVal);
	}
	else
		return(1);
}
BOOL clsSysTray::AddIcon()
{
	BOOL iRetVal;
	
	NotifyIconData.hWnd = hWnd;
	
	iRetVal = Shell_NotifyIcon(NIM_ADD, &NotifyIconData);
	
	if(iRetVal)
	{
		bInTray = true;
	}
	
	return(iRetVal);
}
BOOL clsSysTray::RemoveIcon()
{
	BOOL iRetVal;
	
	iRetVal = Shell_NotifyIcon(NIM_DELETE, &NotifyIconData);
	
	if(iRetVal)
	{
		bInTray = false;
	}
	
	return(iRetVal);
}