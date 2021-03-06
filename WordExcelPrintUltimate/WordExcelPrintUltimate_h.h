

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Fri Jul 14 11:44:21 2017
 */
/* Compiler settings for WordExcelPrintUltimate.idl:
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


#ifndef __WordExcelPrintUltimate_h_h__
#define __WordExcelPrintUltimate_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IWordExcelPrintUltimate_FWD_DEFINED__
#define __IWordExcelPrintUltimate_FWD_DEFINED__
typedef interface IWordExcelPrintUltimate IWordExcelPrintUltimate;

#endif 	/* __IWordExcelPrintUltimate_FWD_DEFINED__ */


#ifndef __WordExcelPrintUltimate_FWD_DEFINED__
#define __WordExcelPrintUltimate_FWD_DEFINED__

#ifdef __cplusplus
typedef class WordExcelPrintUltimate WordExcelPrintUltimate;
#else
typedef struct WordExcelPrintUltimate WordExcelPrintUltimate;
#endif /* __cplusplus */

#endif 	/* __WordExcelPrintUltimate_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 



#ifndef __WordExcelPrintUltimate_LIBRARY_DEFINED__
#define __WordExcelPrintUltimate_LIBRARY_DEFINED__

/* library WordExcelPrintUltimate */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_WordExcelPrintUltimate;

#ifndef __IWordExcelPrintUltimate_DISPINTERFACE_DEFINED__
#define __IWordExcelPrintUltimate_DISPINTERFACE_DEFINED__

/* dispinterface IWordExcelPrintUltimate */
/* [uuid] */ 


EXTERN_C const IID DIID_IWordExcelPrintUltimate;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("D835C838-DC60-4B8B-AD42-FB9E54417843")
    IWordExcelPrintUltimate : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IWordExcelPrintUltimateVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IWordExcelPrintUltimate * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IWordExcelPrintUltimate * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IWordExcelPrintUltimate * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IWordExcelPrintUltimate * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IWordExcelPrintUltimate * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IWordExcelPrintUltimate * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IWordExcelPrintUltimate * This,
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
    } IWordExcelPrintUltimateVtbl;

    interface IWordExcelPrintUltimate
    {
        CONST_VTBL struct IWordExcelPrintUltimateVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IWordExcelPrintUltimate_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IWordExcelPrintUltimate_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IWordExcelPrintUltimate_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IWordExcelPrintUltimate_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IWordExcelPrintUltimate_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IWordExcelPrintUltimate_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IWordExcelPrintUltimate_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IWordExcelPrintUltimate_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_WordExcelPrintUltimate;

#ifdef __cplusplus

class DECLSPEC_UUID("5CE108BC-836E-41CE-9946-02DD894E5A53")
WordExcelPrintUltimate;
#endif
#endif /* __WordExcelPrintUltimate_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


