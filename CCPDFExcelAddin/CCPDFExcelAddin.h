

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0500 */
/* at Mon May 07 20:18:00 2012
 */
/* Compiler settings for CCPDFExcelAddin.idl:
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

#ifndef __CCPDFExcelAddin_h__
#define __CCPDFExcelAddin_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __ICCPDFExcelAddinObj_FWD_DEFINED__
#define __ICCPDFExcelAddinObj_FWD_DEFINED__
typedef interface ICCPDFExcelAddinObj ICCPDFExcelAddinObj;
#endif 	/* __ICCPDFExcelAddinObj_FWD_DEFINED__ */


#ifndef __CCPDFExcelAddinObj_FWD_DEFINED__
#define __CCPDFExcelAddinObj_FWD_DEFINED__

#ifdef __cplusplus
typedef class CCPDFExcelAddinObj CCPDFExcelAddinObj;
#else
typedef struct CCPDFExcelAddinObj CCPDFExcelAddinObj;
#endif /* __cplusplus */

#endif 	/* __CCPDFExcelAddinObj_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __ICCPDFExcelAddinObj_INTERFACE_DEFINED__
#define __ICCPDFExcelAddinObj_INTERFACE_DEFINED__

/* interface ICCPDFExcelAddinObj */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ICCPDFExcelAddinObj;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("1FD93D67-9174-4BD0-A7D2-9BA241870079")
    ICCPDFExcelAddinObj : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct ICCPDFExcelAddinObjVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICCPDFExcelAddinObj * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICCPDFExcelAddinObj * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICCPDFExcelAddinObj * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICCPDFExcelAddinObj * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICCPDFExcelAddinObj * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICCPDFExcelAddinObj * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICCPDFExcelAddinObj * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } ICCPDFExcelAddinObjVtbl;

    interface ICCPDFExcelAddinObj
    {
        CONST_VTBL struct ICCPDFExcelAddinObjVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICCPDFExcelAddinObj_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICCPDFExcelAddinObj_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICCPDFExcelAddinObj_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICCPDFExcelAddinObj_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICCPDFExcelAddinObj_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICCPDFExcelAddinObj_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICCPDFExcelAddinObj_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICCPDFExcelAddinObj_INTERFACE_DEFINED__ */



#ifndef __CCPDFEXCELADDINLib_LIBRARY_DEFINED__
#define __CCPDFEXCELADDINLib_LIBRARY_DEFINED__

/* library CCPDFEXCELADDINLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_CCPDFEXCELADDINLib;

EXTERN_C const CLSID CLSID_CCPDFExcelAddinObj;

#ifdef __cplusplus

class DECLSPEC_UUID("7A13F11B-2986-4C4C-82B4-D5FEA66BB36F")
CCPDFExcelAddinObj;
#endif
#endif /* __CCPDFEXCELADDINLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


