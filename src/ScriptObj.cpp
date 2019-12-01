#include "StdAfx.h"
#include "ScriptObj.h"

CScriptObj::CScriptObj(void)
{
}

CScriptObj::~CScriptObj(void)
{
	CloseEngine();
}

STDMETHODIMP CScriptObj::alert(BSTR bsText)
{
	MessageBoxW(NULL, bsText, L"CScriptObj message", MB_OK);
	return S_OK;
}

STDMETHODIMP CScriptObj::sleep(UINT ms)
{
	Sleep(ms);
	return S_OK;
}


STDMETHODIMP CScriptObj::CallbackTest(IDispatch *pDisp)
{
	CComDispatchDriver pSmartDisp(pDisp);
	CComVariant vRet;

	// note that CComDispatchDriver is just a specialization of CComQIPtr for IDispatch
	// which has helper methods for calling IDispatch::Invoke
	HRESULT hr = pSmartDisp.Invoke0(L"Foo", &vRet);
	if (FAILED(hr))
	{
		::MessageBox(GetActiveWindow(), "Unable to find Foo member on object!", "CScriptObj error!", MB_OK | MB_ICONERROR);
	}
	return S_OK;
}


HRESULT STDMETHODCALLTYPE CScriptObj::GetDocVersionString(/* [out] */ BSTR __RPC_FAR *pbstrVersion)
{
	return E_NOTIMPL;
}
HRESULT STDMETHODCALLTYPE CScriptObj::GetLCID(/* [out] */ LCID __RPC_FAR *plcid)
{
	return E_NOTIMPL;
}

// This method is very important. It "tells" the scripting engine about the site it is running in. 
// This method is called when a script needs to access one of the external objects. It is possible 
// to access second level objects via the main host object like: Preprocessor.Reader.Advance, or it
// is also possible to expose "named" second-level objects to the scripting engine. This is very 
// important when those objects fire events that should be processed by script. Remember: Only named 
// objects' events can be caught by a script. Objects naming is done in CreateEngine, we'll see it 
// later in the code.


HRESULT STDMETHODCALLTYPE CScriptObj::GetItemInfo(/* [in] */ LPCOLESTR pstrName,
									  /* [in] */ DWORD dwReturnMask,
								      /* [out] */ IUnknown __RPC_FAR *__RPC_FAR *ppiunkItem,
									  /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppti)
{



	ATLTRACE("GetItemInfo: Name = %s Mask = %x\n", pstrName, dwReturnMask);

	if (dwReturnMask & SCRIPTINFO_ITYPEINFO)
	{
		if (!ppti)
			return E_INVALIDARG;
		*ppti = NULL;
	}

	if (dwReturnMask & SCRIPTINFO_IUNKNOWN)
	{
		if (!ppiunkItem)
			return E_INVALIDARG;
		*ppiunkItem = NULL;
	}

	// Global object
	if (!_wcsicmp(L"Window", pstrName))
	{
		if (dwReturnMask & SCRIPTINFO_ITYPEINFO)
		{
			GetTypeInfo(0 /* lcid unknown! */, NULL, ppti);
			(*ppti)->AddRef();      // because returning
		}

		if (dwReturnMask & SCRIPTINFO_IUNKNOWN)
		{
			*ppiunkItem = (IDispatch*)this;//pThis->GetIDispatch(TRUE);
			(*ppiunkItem)->AddRef();    // because returning
		}
		return S_OK;
	}

	return E_INVALIDARG;
}

HRESULT STDMETHODCALLTYPE CScriptObj::OnScriptTerminate(/* [in] */ const VARIANT __RPC_FAR *pvarResult,
											/* [in] */ const EXCEPINFO __RPC_FAR *pexcepinfo)
{
	ATLTRACE("\nOnScriptTerminate\r\n");
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CScriptObj::OnStateChange(/* [in] */ SCRIPTSTATE ssScriptState)
{
	ATLTRACE("\nOnStateChange\r\n");
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CScriptObj::OnEnterScript( void)
{
	ATLTRACE("\nOnEnterScript\r\n");
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CScriptObj::OnLeaveScript( void)
{
	ATLTRACE("\nOnLeaveScript\r\n");
	return S_OK;
}


HRESULT STDMETHODCALLTYPE CScriptObj::OnScriptError(/* [in] */ IActiveScriptError __RPC_FAR *pscripterror)
{

	EXCEPINFO ei;
	DWORD     dwSrcContext;
	ULONG     ulLine;
	LONG      ichError;
	BSTR      bstrLine = NULL;
	CString strError;

	pscripterror->GetExceptionInfo(&ei);
	pscripterror->GetSourcePosition(&dwSrcContext, &ulLine, &ichError);
	pscripterror->GetSourceLineText(&bstrLine);

	CString desc;
	CString src;

	desc = (LPCWSTR)ei.bstrDescription;
	src = (LPCWSTR)ei.bstrSource;

	strError.Format("%s\nSrc: %s\nLine:%d Error:%d Scode:%x", desc, src, ulLine, (int)ei.wCode, ei.scode);
	::MessageBox(::GetActiveWindow(),strError,"error",MB_ICONEXCLAMATION);

	ATLTRACE(strError);
	ATLTRACE("\n");

	return S_OK;

}
// The following is an implementation of IActiveScriptSiteWindow: In our case those functions 
// really don't do too much. When asked a window handle, we return null - our host doesn't have
// a window. If we had a window, we'd like it to become modal when a message box is popped from
// the script, in this case, again we do nothing.
HRESULT STDMETHODCALLTYPE CScriptObj::GetWindow(/*[out]*/HWND __RPC_FAR *phwnd)
{
	phwnd = NULL;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CScriptObj::EnableModeless(/*[in]*/BOOL fEnable)
{
	return S_OK;
}


bool CScriptObj::CreateScriptingEngine(LPCTSTR pszLanguage)
{
    USES_CONVERSION;

	HRESULT hr;

    CLSID clsid;
    //HRESULT hr = CLSIDFromProgID(L"VBScript", &clsid);
    //HRESULT hr = CLSIDFromProgID(L"VBS", &clsid);
    //HRESULT hr = CLSIDFromProgID(L"JavaScript", &clsid);
    // hr = CLSIDFromProgID(L"JScript", &clsid);
    hr = CLSIDFromProgID(T2W(pszLanguage), &clsid);

	// XX ActiveX Scripting XX 
    hr = m_ScriptEngine.CoCreateInstance(clsid);
	if (FAILED(hr))
		// If this happens, the scripting engine is probably not properly registered
		return FALSE;
	// Script Engine must support IActiveScriptParse for us to use it
	hr = m_ScriptEngine.QueryInterface(&m_ScriptParser);

	if (FAILED(hr))
		goto Failure;

	hr = m_ScriptEngine->SetScriptSite((IActiveScriptSite*)this);

	if (FAILED(hr))
		goto Failure;

	// InitNew the object:
	hr = m_ScriptParser->InitNew();

	if (FAILED(hr))
		goto Failure;

	// Add Top-level Global Named Item
	hr = m_ScriptEngine->AddNamedItem(L"Window", SCRIPTITEM_ISVISIBLE | SCRIPTITEM_ISSOURCE | SCRIPTITEM_GLOBALMEMBERS);
	if (FAILED(hr))
		ATLTRACE("Couldn't add named item Scripter\n");

	return TRUE;

Failure:
	if (m_ScriptEngine)
		m_ScriptEngine.Release();
	if (m_ScriptParser)
		m_ScriptParser.Release();

	// XX ActiveX Scripting XX
	return FALSE;
}

HRESULT CScriptObj::CloseEngine()
{
	HRESULT	hr = S_OK;
	// Close Script
    if (m_ScriptEngine)
	    hr = m_ScriptEngine->Close();
	if (FAILED(hr))
	{
		ATLTRACE("Attempt to Close AXS failed\n");
		goto FailedToStopScript;
	}
	// NOTE: Handle OLESCRIPT_S_PENDING
	// The documentation reports that if you get this return value,
	// you'll need to wait until IActiveScriptSite::OnStateChange before
	// finishing the Close. However, no current script engine implementation
	// will do this and this return value is not defined in any headers currently
	// (This is true as of IE 3.01)

	if (m_ScriptParser)
	{
		m_ScriptParser = NULL;
	}
	if (m_ScriptEngine)
	{
		m_ScriptEngine->Close();
		m_ScriptEngine = NULL;
	}

	return S_OK;

FailedToStopScript:
	::MessageBox(NULL,"Attempt to stop running script failed!\n", "error", MB_ICONEXCLAMATION | MB_OK);
    return hr;
}

bool CScriptObj::RunScript(LPCTSTR pszScript)
{
    if (!m_ScriptParser)
        return false;

    EXCEPINFO ei = {0};
    CComBSTR bstrParseMe(pszScript);
	HRESULT hr = m_ScriptParser->ParseScriptText(bstrParseMe, L"Window", NULL, NULL, 0, 0, 0L, NULL, &ei);
    if (SUCCEEDED(hr))
    {

	    // Start the script running...
	    hr = m_ScriptEngine->SetScriptState(SCRIPTSTATE_CONNECTED);
	    if (FAILED(hr))
	    {
		    ATLTRACE("SetScriptState to CONNECTED failed\n");
	    }
    }
	else
	{
		MessageBox(GetActiveWindow(), "Failed to parse script", "error", MB_OK | MB_ICONERROR);
	}

	return SUCCEEDED(hr);
}
