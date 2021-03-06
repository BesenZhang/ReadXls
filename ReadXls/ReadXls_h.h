

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Wed May 17 20:41:32 2017
 */
/* Compiler settings for ReadXls.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

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


#ifndef __ReadXls_h_h__
#define __ReadXls_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IReadXls_FWD_DEFINED__
#define __IReadXls_FWD_DEFINED__
typedef interface IReadXls IReadXls;

#endif 	/* __IReadXls_FWD_DEFINED__ */


#ifndef __ReadXls_FWD_DEFINED__
#define __ReadXls_FWD_DEFINED__

#ifdef __cplusplus
typedef class ReadXls ReadXls;
#else
typedef struct ReadXls ReadXls;
#endif /* __cplusplus */

#endif 	/* __ReadXls_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 



#ifndef __ReadXls_LIBRARY_DEFINED__
#define __ReadXls_LIBRARY_DEFINED__

/* library ReadXls */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_ReadXls;

#ifndef __IReadXls_DISPINTERFACE_DEFINED__
#define __IReadXls_DISPINTERFACE_DEFINED__

/* dispinterface IReadXls */
/* [uuid] */ 


EXTERN_C const IID DIID_IReadXls;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("6B3EFF69-A4AD-4306-BD9C-D7863A4202B2")
    IReadXls : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IReadXlsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IReadXls * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IReadXls * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IReadXls * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IReadXls * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IReadXls * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IReadXls * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IReadXls * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IReadXlsVtbl;

    interface IReadXls
    {
        CONST_VTBL struct IReadXlsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IReadXls_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IReadXls_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IReadXls_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IReadXls_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IReadXls_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IReadXls_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IReadXls_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IReadXls_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_ReadXls;

#ifdef __cplusplus

class DECLSPEC_UUID("174D6E32-4900-420B-9006-1C6D299D4ABC")
ReadXls;
#endif
#endif /* __ReadXls_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


