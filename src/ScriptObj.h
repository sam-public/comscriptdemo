#pragma once

[uuid("34BB42F4-FB90-4DB3-9F80-B30FA9E27352")]
class CScriptObj :
	public CComObjectRoot,
    public CComCoClass<CScriptObj>,
	public IDispatchImpl<IScriptObj, &IID_IScriptObj, &LIBID_ScriptDemo, 0xffff, 0xffff>,
	public ISupportErrorInfoImpl<&IID_IScriptObj>,
	public IActiveScriptSite,
	public IActiveScriptSiteWindow
{
public:
	CScriptObj(void);
	~CScriptObj(void);

    BEGIN_COM_MAP(CScriptObj)
	    COM_INTERFACE_ENTRY(CScriptObj)	// for CreateInstance to work
	    COM_INTERFACE_ENTRY(IDispatch)
	    COM_INTERFACE_ENTRY(IScriptObj)
	    COM_INTERFACE_ENTRY(IActiveScriptSite)
	    COM_INTERFACE_ENTRY(IActiveScriptSiteWindow)
	    COM_INTERFACE_ENTRY(ISupportErrorInfo)
    END_COM_MAP()
    // IActiveScriptSite
	virtual HRESULT STDMETHODCALLTYPE GetLCID(LCID *plcid);
    virtual HRESULT STDMETHODCALLTYPE GetItemInfo(
		LPCOLESTR pstrName,DWORD dwReturnMask, IUnknown **ppiunkItem, ITypeInfo **ppti);
    virtual HRESULT STDMETHODCALLTYPE GetDocVersionString(BSTR *pbstrVersion);
    virtual HRESULT STDMETHODCALLTYPE OnScriptTerminate(const VARIANT *pvarResult,const EXCEPINFO *pexcepinfo);
    virtual HRESULT STDMETHODCALLTYPE OnStateChange(SCRIPTSTATE ssScriptState);
    virtual HRESULT STDMETHODCALLTYPE OnScriptError(IActiveScriptError *pscripterror);
    virtual HRESULT STDMETHODCALLTYPE OnEnterScript( void);
    virtual HRESULT STDMETHODCALLTYPE OnLeaveScript( void);

    // IActiveScriptSiteWindow
	virtual HRESULT STDMETHODCALLTYPE GetWindow(HWND *phwnd);
    virtual HRESULT STDMETHODCALLTYPE EnableModeless(BOOL fEnable);

	//IScriptObj methods
	STDMETHODIMP alert(BSTR bsText);
    STDMETHODIMP CallbackTest(IDispatch *pDisp);
	STDMETHODIMP sleep(UINT ms);

	// public class methods
	bool CreateScriptingEngine(LPCTSTR pszLanguage);
	bool RunScript(LPCTSTR pszScript);


private:
	HRESULT CloseEngine();

	CComPtr<IActiveScriptParse> m_ScriptParser;
	CComPtr<IActiveScript>      m_ScriptEngine;

};
