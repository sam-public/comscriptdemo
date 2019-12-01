// MainFrm.cpp : implmentation of the CMainFrame class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "aboutdlg.h"
#include "scriptdemoView.h"
#include "MainFrm.h"

LPCTSTR c_pszScript = "\
// simple demonstration of javascript calling back and forth to C++ with COM/IDispatch \r\n\
// this whole app is < 1000 lines, and about 2hrs to write \r\n\
// change this script if you want, F5 to run it \r\n\
\r\n\
// alert is a function in C++ ( CScriptObj::alert ) which just shows a MessageBox \r\n\
// CallbackTest is a function in C++ which expects to take a COM object implementing IDispatch\r\n\
// and then calls its Foo method\r\n\
\r\n\
\r\n\
var globalvariable = 1;\r\n\
\r\n\
bar = new Object(); 		// make an object in javascript\r\n\
bar.Foo = function()\r\n\
{\r\n\
  alert(\"inside object bar\"); \r\n\
  globalvariable++;\r\n\
};\r\n\
\r\n\
alert(\"start \\r\\nglobalvariable=\" + globalvariable);\r\n\
CallbackTest(bar);\r\n\
alert(\"done \\r\\nglobalvariable=\" + globalvariable);\r\n\
\r\n\
\r\n\
// demonstrate standard COM stuff you get for free- literally zero C++ code written to support this\r\n\
\r\n\
var ExcelSheet;\r\n\
ExcelApp = new ActiveXObject(\"Excel.Application\");\r\n\
ExcelSheet = new ActiveXObject(\"Excel.Sheet\");\r\n\
ExcelSheet.Application.Visible = true;\r\n\
sleep(3000);\r\n\
ExcelSheet.Application.Quit();\r\n\
";

BOOL CMainFrame::PreTranslateMessage(MSG* pMsg)
{
	if(CFrameWindowImpl<CMainFrame>::PreTranslateMessage(pMsg))
		return TRUE;

	return m_view.PreTranslateMessage(pMsg);
}

BOOL CMainFrame::OnIdle()
{
	return FALSE;
}

LRESULT CMainFrame::OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{

	CreateSimpleStatusBar();

	m_hWndClient = m_view.Create(m_hWnd, rcDefault, NULL, ES_MULTILINE | WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, WS_EX_CLIENTEDGE);
	UISetCheck(ID_VIEW_STATUS_BAR, 1);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);
	pLoop->AddIdleHandler(this);

	CScriptObj::CreateInstance(&m_pobj);
	m_pobj->AddRef();
	m_pobj->CreateScriptingEngine("JScript");

	m_view.SetWindowText(c_pszScript);
	m_view.SendMessage(EM_SETSEL,-1,-1);


	return 0;
}

LRESULT CMainFrame::OnEditCut(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	m_view.SendMessage(WM_CUT);
	return 0;
}
LRESULT CMainFrame::OnEditCopy(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	m_view.SendMessage(WM_COPY);
	return 0;
}
LRESULT CMainFrame::OnEditPaste(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	m_view.SendMessage(WM_PASTE);
	return 0;
}


LRESULT CMainFrame::OnFileExit(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	PostMessage(WM_CLOSE);
	return 0;
}

LRESULT CMainFrame::OnFileNew(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	// TODO: add code to initialize document

	return 0;
}

LRESULT CMainFrame::OnViewStatusBar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	BOOL bVisible = !::IsWindowVisible(m_hWndStatusBar);
	::ShowWindow(m_hWndStatusBar, bVisible ? SW_SHOWNOACTIVATE : SW_HIDE);
	UISetCheck(ID_VIEW_STATUS_BAR, bVisible);
	UpdateLayout();
	return 0;
}

LRESULT CMainFrame::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	CAboutDlg dlg;
	dlg.DoModal();
	return 0;
}

LRESULT CMainFrame::OnRunScript(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	CString s;
	m_view.GetWindowText(s);
	bool bRet = m_pobj->RunScript(s);
	return 0;
}
