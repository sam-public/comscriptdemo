

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0500 */
/* at Tue Sep 09 23:43:50 2008
 */
/* Compiler settings for .\scriptdemo.idl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __scriptdemo_h__
#define __scriptdemo_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IScriptObj_FWD_DEFINED__
#define __IScriptObj_FWD_DEFINED__
typedef interface IScriptObj IScriptObj;
#endif 	/* __IScriptObj_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 



#ifndef __ScriptDemo_LIBRARY_DEFINED__
#define __ScriptDemo_LIBRARY_DEFINED__

/* library ScriptDemo */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_ScriptDemo;

#ifndef __IScriptObj_INTERFACE_DEFINED__
#define __IScriptObj_INTERFACE_DEFINED__

/* interface IScriptObj */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IScriptObj;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("D5AD8914-14A4-4707-8F53-6CDC55258E56")
    IScriptObj : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE alert( 
            /* [in] */ BSTR bsText) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CallbackTest( 
            /* [in] */ IDispatch *pDisp) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE sleep( 
            /* [in] */ UINT ms) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IScriptObjVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IScriptObj * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IScriptObj * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IScriptObj * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IScriptObj * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IScriptObj * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IScriptObj * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IScriptObj * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *alert )( 
            IScriptObj * This,
            /* [in] */ BSTR bsText);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *CallbackTest )( 
            IScriptObj * This,
            /* [in] */ IDispatch *pDisp);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *sleep )( 
            IScriptObj * This,
            /* [in] */ UINT ms);
        
        END_INTERFACE
    } IScriptObjVtbl;

    interface IScriptObj
    {
        CONST_VTBL struct IScriptObjVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IScriptObj_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IScriptObj_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IScriptObj_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IScriptObj_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IScriptObj_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IScriptObj_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IScriptObj_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IScriptObj_alert(This,bsText)	\
    ( (This)->lpVtbl -> alert(This,bsText) ) 

#define IScriptObj_CallbackTest(This,pDisp)	\
    ( (This)->lpVtbl -> CallbackTest(This,pDisp) ) 

#define IScriptObj_sleep(This,ms)	\
    ( (This)->lpVtbl -> sleep(This,ms) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IScriptObj_INTERFACE_DEFINED__ */

#endif /* __ScriptDemo_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


