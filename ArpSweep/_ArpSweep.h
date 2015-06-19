

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 6.00.0361 */
/* at Sat Oct 29 21:02:33 2005
 */
/* Compiler settings for _ArpSweep.idl:
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

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef ___ArpSweep_h__
#define ___ArpSweep_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IArpScanner_FWD_DEFINED__
#define __IArpScanner_FWD_DEFINED__
typedef interface IArpScanner IArpScanner;
#endif 	/* __IArpScanner_FWD_DEFINED__ */


#ifndef ___IArpScannerEvents_FWD_DEFINED__
#define ___IArpScannerEvents_FWD_DEFINED__
typedef interface _IArpScannerEvents _IArpScannerEvents;
#endif 	/* ___IArpScannerEvents_FWD_DEFINED__ */


#ifndef __CArpScanner_FWD_DEFINED__
#define __CArpScanner_FWD_DEFINED__

#ifdef __cplusplus
typedef class CArpScanner CArpScanner;
#else
typedef struct CArpScanner CArpScanner;
#endif /* __cplusplus */

#endif 	/* __CArpScanner_FWD_DEFINED__ */


/* header files for imported files */
#include "prsht.h"
#include "mshtml.h"
#include "mshtmhst.h"
#include "exdisp.h"
#include "objsafe.h"

#ifdef __cplusplus
extern "C"{
#endif 

void * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void * ); 

#ifndef __IArpScanner_INTERFACE_DEFINED__
#define __IArpScanner_INTERFACE_DEFINED__

/* interface IArpScanner */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IArpScanner;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("12F0C3AF-5936-4601-A810-F0326F805952")
    IArpScanner : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_StartIP( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_StartIP( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EndIP( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EndIP( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Key( 
            /* [in] */ LONG newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Adapter( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Adapter( 
            /* [in] */ LONG newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_HostCount( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE StartScan( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_AdapterCount( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetAdapterInfo( 
            /* [in] */ LONG adapter,
            /* [out] */ BSTR *name,
            /* [out] */ BSTR *description,
            /* [out] */ BSTR *ip,
            /* [out] */ BSTR *mac,
            /* [retval][out] */ BOOL *ok) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Scan( 
            /* [in] */ LONG adapter,
            /* [in] */ BSTR ip,
            /* [in] */ BSTR mac,
            /* [retval][out] */ LONG *response) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE StopScan( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LookupAdapterName( 
            /* [in] */ BSTR AdapterName,
            /* [retval][out] */ LONG *Adapter) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ScannedHostCount( 
            /* [retval][out] */ LONG *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IArpScannerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IArpScanner * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IArpScanner * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IArpScanner * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IArpScanner * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IArpScanner * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IArpScanner * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IArpScanner * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_StartIP )( 
            IArpScanner * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_StartIP )( 
            IArpScanner * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_EndIP )( 
            IArpScanner * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_EndIP )( 
            IArpScanner * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Key )( 
            IArpScanner * This,
            /* [in] */ LONG newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Adapter )( 
            IArpScanner * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Adapter )( 
            IArpScanner * This,
            /* [in] */ LONG newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HostCount )( 
            IArpScanner * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *StartScan )( 
            IArpScanner * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_AdapterCount )( 
            IArpScanner * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetAdapterInfo )( 
            IArpScanner * This,
            /* [in] */ LONG adapter,
            /* [out] */ BSTR *name,
            /* [out] */ BSTR *description,
            /* [out] */ BSTR *ip,
            /* [out] */ BSTR *mac,
            /* [retval][out] */ BOOL *ok);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Scan )( 
            IArpScanner * This,
            /* [in] */ LONG adapter,
            /* [in] */ BSTR ip,
            /* [in] */ BSTR mac,
            /* [retval][out] */ LONG *response);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *StopScan )( 
            IArpScanner * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *LookupAdapterName )( 
            IArpScanner * This,
            /* [in] */ BSTR AdapterName,
            /* [retval][out] */ LONG *Adapter);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ScannedHostCount )( 
            IArpScanner * This,
            /* [retval][out] */ LONG *pVal);
        
        END_INTERFACE
    } IArpScannerVtbl;

    interface IArpScanner
    {
        CONST_VTBL struct IArpScannerVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IArpScanner_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IArpScanner_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IArpScanner_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IArpScanner_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IArpScanner_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IArpScanner_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IArpScanner_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IArpScanner_get_StartIP(This,pVal)	\
    (This)->lpVtbl -> get_StartIP(This,pVal)

#define IArpScanner_put_StartIP(This,newVal)	\
    (This)->lpVtbl -> put_StartIP(This,newVal)

#define IArpScanner_get_EndIP(This,pVal)	\
    (This)->lpVtbl -> get_EndIP(This,pVal)

#define IArpScanner_put_EndIP(This,newVal)	\
    (This)->lpVtbl -> put_EndIP(This,newVal)

#define IArpScanner_put_Key(This,newVal)	\
    (This)->lpVtbl -> put_Key(This,newVal)

#define IArpScanner_get_Adapter(This,pVal)	\
    (This)->lpVtbl -> get_Adapter(This,pVal)

#define IArpScanner_put_Adapter(This,newVal)	\
    (This)->lpVtbl -> put_Adapter(This,newVal)

#define IArpScanner_get_HostCount(This,pVal)	\
    (This)->lpVtbl -> get_HostCount(This,pVal)

#define IArpScanner_StartScan(This)	\
    (This)->lpVtbl -> StartScan(This)

#define IArpScanner_get_AdapterCount(This,pVal)	\
    (This)->lpVtbl -> get_AdapterCount(This,pVal)

#define IArpScanner_GetAdapterInfo(This,adapter,name,description,ip,mac,ok)	\
    (This)->lpVtbl -> GetAdapterInfo(This,adapter,name,description,ip,mac,ok)

#define IArpScanner_Scan(This,adapter,ip,mac,response)	\
    (This)->lpVtbl -> Scan(This,adapter,ip,mac,response)

#define IArpScanner_StopScan(This)	\
    (This)->lpVtbl -> StopScan(This)

#define IArpScanner_LookupAdapterName(This,AdapterName,Adapter)	\
    (This)->lpVtbl -> LookupAdapterName(This,AdapterName,Adapter)

#define IArpScanner_get_ScannedHostCount(This,pVal)	\
    (This)->lpVtbl -> get_ScannedHostCount(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArpScanner_get_StartIP_Proxy( 
    IArpScanner * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IArpScanner_get_StartIP_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArpScanner_put_StartIP_Proxy( 
    IArpScanner * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IArpScanner_put_StartIP_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArpScanner_get_EndIP_Proxy( 
    IArpScanner * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IArpScanner_get_EndIP_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArpScanner_put_EndIP_Proxy( 
    IArpScanner * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IArpScanner_put_EndIP_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArpScanner_put_Key_Proxy( 
    IArpScanner * This,
    /* [in] */ LONG newVal);


void __RPC_STUB IArpScanner_put_Key_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArpScanner_get_Adapter_Proxy( 
    IArpScanner * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IArpScanner_get_Adapter_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArpScanner_put_Adapter_Proxy( 
    IArpScanner * This,
    /* [in] */ LONG newVal);


void __RPC_STUB IArpScanner_put_Adapter_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArpScanner_get_HostCount_Proxy( 
    IArpScanner * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IArpScanner_get_HostCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArpScanner_StartScan_Proxy( 
    IArpScanner * This);


void __RPC_STUB IArpScanner_StartScan_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArpScanner_get_AdapterCount_Proxy( 
    IArpScanner * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IArpScanner_get_AdapterCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArpScanner_GetAdapterInfo_Proxy( 
    IArpScanner * This,
    /* [in] */ LONG adapter,
    /* [out] */ BSTR *name,
    /* [out] */ BSTR *description,
    /* [out] */ BSTR *ip,
    /* [out] */ BSTR *mac,
    /* [retval][out] */ BOOL *ok);


void __RPC_STUB IArpScanner_GetAdapterInfo_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArpScanner_Scan_Proxy( 
    IArpScanner * This,
    /* [in] */ LONG adapter,
    /* [in] */ BSTR ip,
    /* [in] */ BSTR mac,
    /* [retval][out] */ LONG *response);


void __RPC_STUB IArpScanner_Scan_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArpScanner_StopScan_Proxy( 
    IArpScanner * This);


void __RPC_STUB IArpScanner_StopScan_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArpScanner_LookupAdapterName_Proxy( 
    IArpScanner * This,
    /* [in] */ BSTR AdapterName,
    /* [retval][out] */ LONG *Adapter);


void __RPC_STUB IArpScanner_LookupAdapterName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArpScanner_get_ScannedHostCount_Proxy( 
    IArpScanner * This,
    /* [retval][out] */ LONG *pVal);


void __RPC_STUB IArpScanner_get_ScannedHostCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IArpScanner_INTERFACE_DEFINED__ */



#ifndef __ArpSweep_LIBRARY_DEFINED__
#define __ArpSweep_LIBRARY_DEFINED__

/* library ArpSweep */
/* [helpstring][uuid][version] */ 


EXTERN_C const IID LIBID_ArpSweep;

#ifndef ___IArpScannerEvents_DISPINTERFACE_DEFINED__
#define ___IArpScannerEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IArpScannerEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__IArpScannerEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("59753AB6-4CD0-40D2-9EFC-FA452CBF2051")
    _IArpScannerEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IArpScannerEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _IArpScannerEvents * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _IArpScannerEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _IArpScannerEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _IArpScannerEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _IArpScannerEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _IArpScannerEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _IArpScannerEvents * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } _IArpScannerEventsVtbl;

    interface _IArpScannerEvents
    {
        CONST_VTBL struct _IArpScannerEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IArpScannerEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _IArpScannerEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _IArpScannerEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _IArpScannerEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _IArpScannerEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _IArpScannerEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _IArpScannerEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IArpScannerEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_CArpScanner;

#ifdef __cplusplus

class DECLSPEC_UUID("F6CDAC27-27C7-4030-BAC7-E3E7143DDE0B")
CArpScanner;
#endif
#endif /* __ArpSweep_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


