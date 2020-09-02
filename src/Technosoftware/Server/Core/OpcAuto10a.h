/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 3.01.75 */
/* at Tue Mar 02 14:11:12 1999
 */
/* Compiler settings for opcauto10a.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: none
*/
//@@MIDL_FILE_HEADING(  )
#include "rpc.h"
#include "rpcndr.h"
#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __opcauto10a_h__
#define __opcauto10a_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IOPCServerDisp_FWD_DEFINED__
#define __IOPCServerDisp_FWD_DEFINED__
typedef interface IOPCServerDisp IOPCServerDisp;
#endif 	/* __IOPCServerDisp_FWD_DEFINED__ */


#ifndef __IOPCServerPublicGroupsDisp_FWD_DEFINED__
#define __IOPCServerPublicGroupsDisp_FWD_DEFINED__
typedef interface IOPCServerPublicGroupsDisp IOPCServerPublicGroupsDisp;
#endif 	/* __IOPCServerPublicGroupsDisp_FWD_DEFINED__ */


#ifndef __IOPCBrowseServerAddressSpaceDisp_FWD_DEFINED__
#define __IOPCBrowseServerAddressSpaceDisp_FWD_DEFINED__
typedef interface IOPCBrowseServerAddressSpaceDisp IOPCBrowseServerAddressSpaceDisp;
#endif 	/* __IOPCBrowseServerAddressSpaceDisp_FWD_DEFINED__ */


#ifndef __IOPCItemMgtDisp_FWD_DEFINED__
#define __IOPCItemMgtDisp_FWD_DEFINED__
typedef interface IOPCItemMgtDisp IOPCItemMgtDisp;
#endif 	/* __IOPCItemMgtDisp_FWD_DEFINED__ */


#ifndef __IOPCGroupStateMgtDisp_FWD_DEFINED__
#define __IOPCGroupStateMgtDisp_FWD_DEFINED__
typedef interface IOPCGroupStateMgtDisp IOPCGroupStateMgtDisp;
#endif 	/* __IOPCGroupStateMgtDisp_FWD_DEFINED__ */


#ifndef __IOPCSyncIODisp_FWD_DEFINED__
#define __IOPCSyncIODisp_FWD_DEFINED__
typedef interface IOPCSyncIODisp IOPCSyncIODisp;
#endif 	/* __IOPCSyncIODisp_FWD_DEFINED__ */


#ifndef __IOPCAsyncIODisp_FWD_DEFINED__
#define __IOPCAsyncIODisp_FWD_DEFINED__
typedef interface IOPCAsyncIODisp IOPCAsyncIODisp;
#endif 	/* __IOPCAsyncIODisp_FWD_DEFINED__ */


#ifndef __IOPCPublicGroupStateMgtDisp_FWD_DEFINED__
#define __IOPCPublicGroupStateMgtDisp_FWD_DEFINED__
typedef interface IOPCPublicGroupStateMgtDisp IOPCPublicGroupStateMgtDisp;
#endif 	/* __IOPCPublicGroupStateMgtDisp_FWD_DEFINED__ */


#ifndef __IOPCItemDisp_FWD_DEFINED__
#define __IOPCItemDisp_FWD_DEFINED__
typedef interface IOPCItemDisp IOPCItemDisp;
#endif 	/* __IOPCItemDisp_FWD_DEFINED__ */


#ifndef __OPCServer_FWD_DEFINED__
#define __OPCServer_FWD_DEFINED__

#ifdef __cplusplus
typedef class OPCServer OPCServer;
#else
typedef struct OPCServer OPCServer;
#endif /* __cplusplus */

#endif 	/* __OPCServer_FWD_DEFINED__ */


#ifndef __OPCGroup_FWD_DEFINED__
#define __OPCGroup_FWD_DEFINED__

#ifdef __cplusplus
typedef class OPCGroup OPCGroup;
#else
typedef struct OPCGroup OPCGroup;
#endif /* __cplusplus */

#endif 	/* __OPCGroup_FWD_DEFINED__ */


#ifndef __OPCItem_FWD_DEFINED__
#define __OPCItem_FWD_DEFINED__

#ifdef __cplusplus
typedef class OPCItem OPCItem;
#else
typedef struct OPCItem OPCItem;
#endif /* __cplusplus */

#endif 	/* __OPCItem_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

#ifndef __IOPCServerDisp_INTERFACE_DEFINED__
#define __IOPCServerDisp_INTERFACE_DEFINED__

/****************************************
 * Generated header for interface: IOPCServerDisp
 * at Tue Mar 02 14:11:12 1999
 * using MIDL 3.01.75
 ****************************************/
/* [object][unique][helpstring][dual][uuid] */ 



EXTERN_C const IID IID_IOPCServerDisp;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface DECLSPEC_UUID("39c13a57-011e-11d0-9675-0020afd8adb3")
    IOPCServerDisp : public IDispatch
    {
    public:
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pCount) = 0;
        
        virtual /* [id][restricted][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_StartTime( 
            /* [retval][out] */ DATE __RPC_FAR *pStartTime) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_CurrentTime( 
            /* [retval][out] */ DATE __RPC_FAR *pCurrentTime) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_LastUpdateTime( 
            /* [retval][out] */ DATE __RPC_FAR *pLastUpdateTime) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_MajorVersion( 
            /* [retval][out] */ short __RPC_FAR *pMajorVersion) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_MinorVersion( 
            /* [retval][out] */ short __RPC_FAR *pMinorVersion) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_BuildNumber( 
            /* [retval][out] */ short __RPC_FAR *pBuildNumber) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_VendorInfo( 
            /* [retval][out] */ BSTR __RPC_FAR *pVendorInfo) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Item( 
            /* [in] */ VARIANT Item,
            /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE AddGroup( 
            /* [in] */ BSTR Name,
            /* [in] */ VARIANT_BOOL Active,
            /* [in] */ long RequestedUpdateRate,
            /* [in] */ long ClientGroupHandle,
            /* [in] */ float __RPC_FAR *pPercentDeadband,
            /* [in] */ long LCID,
            /* [out] */ long __RPC_FAR *pServerGroupHandle,
            /* [out] */ long __RPC_FAR *pRevisedUpdateRate,
            /* [optional][in] */ VARIANT __RPC_FAR *pTimeBias,
            /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetErrorString( 
            /* [in] */ long Error,
            /* [in] */ long Locale,
            /* [retval][out] */ BSTR __RPC_FAR *ErrorString) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetGroupByName( 
            /* [in] */ BSTR Name,
            /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE RemoveGroup( 
            /* [in] */ long ServerGroupHandle,
            /* [in] */ VARIANT_BOOL Force) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SaveConfig( 
            /* [in] */ BSTR FileName) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE LoadConfig( 
            /* [in] */ BSTR FileName) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetEnumeratorType( 
            /* [in] */ long Scope,
            /* [in] */ short Type) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOPCServerDispVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOPCServerDisp __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOPCServerDisp __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pCount);
        
        /* [id][restricted][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_StartTime )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [retval][out] */ DATE __RPC_FAR *pStartTime);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CurrentTime )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [retval][out] */ DATE __RPC_FAR *pCurrentTime);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LastUpdateTime )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [retval][out] */ DATE __RPC_FAR *pLastUpdateTime);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_MajorVersion )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pMajorVersion);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_MinorVersion )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pMinorVersion);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_BuildNumber )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pBuildNumber);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_VendorInfo )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVendorInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Item )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [in] */ VARIANT Item,
            /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddGroup )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [in] */ BSTR Name,
            /* [in] */ VARIANT_BOOL Active,
            /* [in] */ long RequestedUpdateRate,
            /* [in] */ long ClientGroupHandle,
            /* [in] */ float __RPC_FAR *pPercentDeadband,
            /* [in] */ long LCID,
            /* [out] */ long __RPC_FAR *pServerGroupHandle,
            /* [out] */ long __RPC_FAR *pRevisedUpdateRate,
            /* [optional][in] */ VARIANT __RPC_FAR *pTimeBias,
            /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetErrorString )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [in] */ long Error,
            /* [in] */ long Locale,
            /* [retval][out] */ BSTR __RPC_FAR *ErrorString);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetGroupByName )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [in] */ BSTR Name,
            /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RemoveGroup )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [in] */ long ServerGroupHandle,
            /* [in] */ VARIANT_BOOL Force);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SaveConfig )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [in] */ BSTR FileName);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *LoadConfig )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [in] */ BSTR FileName);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetEnumeratorType )( 
            IOPCServerDisp __RPC_FAR * This,
            /* [in] */ long Scope,
            /* [in] */ short Type);
        
        END_INTERFACE
    } IOPCServerDispVtbl;

    interface IOPCServerDisp
    {
        CONST_VTBL struct IOPCServerDispVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOPCServerDisp_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOPCServerDisp_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOPCServerDisp_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOPCServerDisp_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IOPCServerDisp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IOPCServerDisp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IOPCServerDisp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IOPCServerDisp_get_Count(This,pCount)	\
    (This)->lpVtbl -> get_Count(This,pCount)

#define IOPCServerDisp_get__NewEnum(This,ppUnk)	\
    (This)->lpVtbl -> get__NewEnum(This,ppUnk)

#define IOPCServerDisp_get_StartTime(This,pStartTime)	\
    (This)->lpVtbl -> get_StartTime(This,pStartTime)

#define IOPCServerDisp_get_CurrentTime(This,pCurrentTime)	\
    (This)->lpVtbl -> get_CurrentTime(This,pCurrentTime)

#define IOPCServerDisp_get_LastUpdateTime(This,pLastUpdateTime)	\
    (This)->lpVtbl -> get_LastUpdateTime(This,pLastUpdateTime)

#define IOPCServerDisp_get_MajorVersion(This,pMajorVersion)	\
    (This)->lpVtbl -> get_MajorVersion(This,pMajorVersion)

#define IOPCServerDisp_get_MinorVersion(This,pMinorVersion)	\
    (This)->lpVtbl -> get_MinorVersion(This,pMinorVersion)

#define IOPCServerDisp_get_BuildNumber(This,pBuildNumber)	\
    (This)->lpVtbl -> get_BuildNumber(This,pBuildNumber)

#define IOPCServerDisp_get_VendorInfo(This,pVendorInfo)	\
    (This)->lpVtbl -> get_VendorInfo(This,pVendorInfo)

#define IOPCServerDisp_Item(This,Item,ppDisp)	\
    (This)->lpVtbl -> Item(This,Item,ppDisp)

#define IOPCServerDisp_AddGroup(This,Name,Active,RequestedUpdateRate,ClientGroupHandle,pPercentDeadband,LCID,pServerGroupHandle,pRevisedUpdateRate,pTimeBias,ppDisp)	\
    (This)->lpVtbl -> AddGroup(This,Name,Active,RequestedUpdateRate,ClientGroupHandle,pPercentDeadband,LCID,pServerGroupHandle,pRevisedUpdateRate,pTimeBias,ppDisp)

#define IOPCServerDisp_GetErrorString(This,Error,Locale,ErrorString)	\
    (This)->lpVtbl -> GetErrorString(This,Error,Locale,ErrorString)

#define IOPCServerDisp_GetGroupByName(This,Name,ppDisp)	\
    (This)->lpVtbl -> GetGroupByName(This,Name,ppDisp)

#define IOPCServerDisp_RemoveGroup(This,ServerGroupHandle,Force)	\
    (This)->lpVtbl -> RemoveGroup(This,ServerGroupHandle,Force)

#define IOPCServerDisp_SaveConfig(This,FileName)	\
    (This)->lpVtbl -> SaveConfig(This,FileName)

#define IOPCServerDisp_LoadConfig(This,FileName)	\
    (This)->lpVtbl -> LoadConfig(This,FileName)

#define IOPCServerDisp_SetEnumeratorType(This,Scope,Type)	\
    (This)->lpVtbl -> SetEnumeratorType(This,Scope,Type)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCServerDisp_get_Count_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pCount);


void __RPC_STUB IOPCServerDisp_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][restricted][propget] */ HRESULT STDMETHODCALLTYPE IOPCServerDisp_get__NewEnum_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk);


void __RPC_STUB IOPCServerDisp_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCServerDisp_get_StartTime_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [retval][out] */ DATE __RPC_FAR *pStartTime);


void __RPC_STUB IOPCServerDisp_get_StartTime_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCServerDisp_get_CurrentTime_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [retval][out] */ DATE __RPC_FAR *pCurrentTime);


void __RPC_STUB IOPCServerDisp_get_CurrentTime_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCServerDisp_get_LastUpdateTime_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [retval][out] */ DATE __RPC_FAR *pLastUpdateTime);


void __RPC_STUB IOPCServerDisp_get_LastUpdateTime_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCServerDisp_get_MajorVersion_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pMajorVersion);


void __RPC_STUB IOPCServerDisp_get_MajorVersion_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCServerDisp_get_MinorVersion_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pMinorVersion);


void __RPC_STUB IOPCServerDisp_get_MinorVersion_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCServerDisp_get_BuildNumber_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pBuildNumber);


void __RPC_STUB IOPCServerDisp_get_BuildNumber_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCServerDisp_get_VendorInfo_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVendorInfo);


void __RPC_STUB IOPCServerDisp_get_VendorInfo_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCServerDisp_Item_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [in] */ VARIANT Item,
    /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp);


void __RPC_STUB IOPCServerDisp_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCServerDisp_AddGroup_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [in] */ BSTR Name,
    /* [in] */ VARIANT_BOOL Active,
    /* [in] */ long RequestedUpdateRate,
    /* [in] */ long ClientGroupHandle,
    /* [in] */ float __RPC_FAR *pPercentDeadband,
    /* [in] */ long LCID,
    /* [out] */ long __RPC_FAR *pServerGroupHandle,
    /* [out] */ long __RPC_FAR *pRevisedUpdateRate,
    /* [optional][in] */ VARIANT __RPC_FAR *pTimeBias,
    /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp);


void __RPC_STUB IOPCServerDisp_AddGroup_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCServerDisp_GetErrorString_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [in] */ long Error,
    /* [in] */ long Locale,
    /* [retval][out] */ BSTR __RPC_FAR *ErrorString);


void __RPC_STUB IOPCServerDisp_GetErrorString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCServerDisp_GetGroupByName_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [in] */ BSTR Name,
    /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp);


void __RPC_STUB IOPCServerDisp_GetGroupByName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCServerDisp_RemoveGroup_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [in] */ long ServerGroupHandle,
    /* [in] */ VARIANT_BOOL Force);


void __RPC_STUB IOPCServerDisp_RemoveGroup_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCServerDisp_SaveConfig_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [in] */ BSTR FileName);


void __RPC_STUB IOPCServerDisp_SaveConfig_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCServerDisp_LoadConfig_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [in] */ BSTR FileName);


void __RPC_STUB IOPCServerDisp_LoadConfig_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCServerDisp_SetEnumeratorType_Proxy( 
    IOPCServerDisp __RPC_FAR * This,
    /* [in] */ long Scope,
    /* [in] */ short Type);


void __RPC_STUB IOPCServerDisp_SetEnumeratorType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOPCServerDisp_INTERFACE_DEFINED__ */


#ifndef __IOPCServerPublicGroupsDisp_INTERFACE_DEFINED__
#define __IOPCServerPublicGroupsDisp_INTERFACE_DEFINED__

/****************************************
 * Generated header for interface: IOPCServerPublicGroupsDisp
 * at Tue Mar 02 14:11:12 1999
 * using MIDL 3.01.75
 ****************************************/
/* [object][unique][helpstring][dual][uuid] */ 



EXTERN_C const IID IID_IOPCServerPublicGroupsDisp;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface DECLSPEC_UUID("39c13a58-011e-11d0-9675-0020afd8adb3")
    IOPCServerPublicGroupsDisp : public IDispatch
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE GetPublicGroupByName( 
            /* [in] */ BSTR Name,
            /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE RemovePublicGroup( 
            /* [in] */ long ServerGroupHandle,
            /* [in] */ VARIANT_BOOL Force) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOPCServerPublicGroupsDispVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOPCServerPublicGroupsDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOPCServerPublicGroupsDisp __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOPCServerPublicGroupsDisp __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IOPCServerPublicGroupsDisp __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IOPCServerPublicGroupsDisp __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IOPCServerPublicGroupsDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IOPCServerPublicGroupsDisp __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPublicGroupByName )( 
            IOPCServerPublicGroupsDisp __RPC_FAR * This,
            /* [in] */ BSTR Name,
            /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RemovePublicGroup )( 
            IOPCServerPublicGroupsDisp __RPC_FAR * This,
            /* [in] */ long ServerGroupHandle,
            /* [in] */ VARIANT_BOOL Force);
        
        END_INTERFACE
    } IOPCServerPublicGroupsDispVtbl;

    interface IOPCServerPublicGroupsDisp
    {
        CONST_VTBL struct IOPCServerPublicGroupsDispVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOPCServerPublicGroupsDisp_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOPCServerPublicGroupsDisp_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOPCServerPublicGroupsDisp_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOPCServerPublicGroupsDisp_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IOPCServerPublicGroupsDisp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IOPCServerPublicGroupsDisp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IOPCServerPublicGroupsDisp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IOPCServerPublicGroupsDisp_GetPublicGroupByName(This,Name,ppDisp)	\
    (This)->lpVtbl -> GetPublicGroupByName(This,Name,ppDisp)

#define IOPCServerPublicGroupsDisp_RemovePublicGroup(This,ServerGroupHandle,Force)	\
    (This)->lpVtbl -> RemovePublicGroup(This,ServerGroupHandle,Force)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IOPCServerPublicGroupsDisp_GetPublicGroupByName_Proxy( 
    IOPCServerPublicGroupsDisp __RPC_FAR * This,
    /* [in] */ BSTR Name,
    /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp);


void __RPC_STUB IOPCServerPublicGroupsDisp_GetPublicGroupByName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCServerPublicGroupsDisp_RemovePublicGroup_Proxy( 
    IOPCServerPublicGroupsDisp __RPC_FAR * This,
    /* [in] */ long ServerGroupHandle,
    /* [in] */ VARIANT_BOOL Force);


void __RPC_STUB IOPCServerPublicGroupsDisp_RemovePublicGroup_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOPCServerPublicGroupsDisp_INTERFACE_DEFINED__ */


#ifndef __IOPCBrowseServerAddressSpaceDisp_INTERFACE_DEFINED__
#define __IOPCBrowseServerAddressSpaceDisp_INTERFACE_DEFINED__

/****************************************
 * Generated header for interface: IOPCBrowseServerAddressSpaceDisp
 * at Tue Mar 02 14:11:12 1999
 * using MIDL 3.01.75
 ****************************************/
/* [object][unique][helpstring][dual][uuid] */ 



EXTERN_C const IID IID_IOPCBrowseServerAddressSpaceDisp;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface DECLSPEC_UUID("39c13a59-011e-11d0-9675-0020afd8adb3")
    IOPCBrowseServerAddressSpaceDisp : public IDispatch
    {
    public:
        virtual /* [id][restricted][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum_BROWSE(
            /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Organization( 
            /* [retval][out] */ long __RPC_FAR *pNameSpaceType) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ChangeBrowsePosition( 
            /* [in] */ long BrowseDirection,
            /* [in] */ BSTR Position) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetItemIDEnumerator( 
            /* [in] */ long BrowseFilterType,
            /* [in] */ BSTR FilterCriteria,
            /* [in] */ VARIANT DataTypeFilter,
            /* [in] */ long AccessRightsFilter) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetItemIDString( 
            /* [in] */ BSTR ItemDataID,
            /* [retval][out] */ BSTR __RPC_FAR *ItemID) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetAccessPathEnumerator( 
            /* [in] */ BSTR ItemID) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOPCBrowseServerAddressSpaceDispVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [id][restricted][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum_BROWSE )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
            /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Organization )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pNameSpaceType);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ChangeBrowsePosition )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
            /* [in] */ long BrowseDirection,
            /* [in] */ BSTR Position);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetItemIDEnumerator )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
            /* [in] */ long BrowseFilterType,
            /* [in] */ BSTR FilterCriteria,
            /* [in] */ VARIANT DataTypeFilter,
            /* [in] */ long AccessRightsFilter);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetItemIDString )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
            /* [in] */ BSTR ItemDataID,
            /* [retval][out] */ BSTR __RPC_FAR *ItemID);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetAccessPathEnumerator )( 
            IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
            /* [in] */ BSTR ItemID);
        
        END_INTERFACE
    } IOPCBrowseServerAddressSpaceDispVtbl;

    interface IOPCBrowseServerAddressSpaceDisp
    {
        CONST_VTBL struct IOPCBrowseServerAddressSpaceDispVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOPCBrowseServerAddressSpaceDisp_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOPCBrowseServerAddressSpaceDisp_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOPCBrowseServerAddressSpaceDisp_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOPCBrowseServerAddressSpaceDisp_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IOPCBrowseServerAddressSpaceDisp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IOPCBrowseServerAddressSpaceDisp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IOPCBrowseServerAddressSpaceDisp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IOPCBrowseServerAddressSpaceDisp_get__NewEnum(This,ppUnk)	\
    (This)->lpVtbl -> get__NewEnum_BROWSE(This,ppUnk)

#define IOPCBrowseServerAddressSpaceDisp_get_Organization(This,pNameSpaceType)	\
    (This)->lpVtbl -> get_Organization(This,pNameSpaceType)

#define IOPCBrowseServerAddressSpaceDisp_ChangeBrowsePosition(This,BrowseDirection,Position)	\
    (This)->lpVtbl -> ChangeBrowsePosition(This,BrowseDirection,Position)

#define IOPCBrowseServerAddressSpaceDisp_SetItemIDEnumerator(This,BrowseFilterType,FilterCriteria,DataTypeFilter,AccessRightsFilter)	\
    (This)->lpVtbl -> SetItemIDEnumerator(This,BrowseFilterType,FilterCriteria,DataTypeFilter,AccessRightsFilter)

#define IOPCBrowseServerAddressSpaceDisp_GetItemIDString(This,ItemDataID,ItemID)	\
    (This)->lpVtbl -> GetItemIDString(This,ItemDataID,ItemID)

#define IOPCBrowseServerAddressSpaceDisp_SetAccessPathEnumerator(This,ItemID)	\
    (This)->lpVtbl -> SetAccessPathEnumerator(This,ItemID)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [id][restricted][propget] */ HRESULT STDMETHODCALLTYPE IOPCBrowseServerAddressSpaceDisp_get__NewEnum_Proxy( 
    IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
    /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk);


void __RPC_STUB IOPCBrowseServerAddressSpaceDisp_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCBrowseServerAddressSpaceDisp_get_Organization_Proxy( 
    IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pNameSpaceType);


void __RPC_STUB IOPCBrowseServerAddressSpaceDisp_get_Organization_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCBrowseServerAddressSpaceDisp_ChangeBrowsePosition_Proxy( 
    IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
    /* [in] */ long BrowseDirection,
    /* [in] */ BSTR Position);


void __RPC_STUB IOPCBrowseServerAddressSpaceDisp_ChangeBrowsePosition_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCBrowseServerAddressSpaceDisp_SetItemIDEnumerator_Proxy( 
    IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
    /* [in] */ long BrowseFilterType,
    /* [in] */ BSTR FilterCriteria,
    /* [in] */ VARIANT DataTypeFilter,
    /* [in] */ long AccessRightsFilter);


void __RPC_STUB IOPCBrowseServerAddressSpaceDisp_SetItemIDEnumerator_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCBrowseServerAddressSpaceDisp_GetItemIDString_Proxy( 
    IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
    /* [in] */ BSTR ItemDataID,
    /* [retval][out] */ BSTR __RPC_FAR *ItemID);


void __RPC_STUB IOPCBrowseServerAddressSpaceDisp_GetItemIDString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCBrowseServerAddressSpaceDisp_SetAccessPathEnumerator_Proxy( 
    IOPCBrowseServerAddressSpaceDisp __RPC_FAR * This,
    /* [in] */ BSTR ItemID);


void __RPC_STUB IOPCBrowseServerAddressSpaceDisp_SetAccessPathEnumerator_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOPCBrowseServerAddressSpaceDisp_INTERFACE_DEFINED__ */


#ifndef __IOPCItemMgtDisp_INTERFACE_DEFINED__
#define __IOPCItemMgtDisp_INTERFACE_DEFINED__

/****************************************
 * Generated header for interface: IOPCItemMgtDisp
 * at Tue Mar 02 14:11:12 1999
 * using MIDL 3.01.75
 ****************************************/
/* [object][unique][helpstring][dual][uuid] */ 



EXTERN_C const IID IID_IOPCItemMgtDisp;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface DECLSPEC_UUID("39c13a5a-011e-11d0-9675-0020afd8adb3")
    IOPCItemMgtDisp : public IDispatch
    {
    public:
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pCount) = 0;
        
        virtual /* [id][restricted][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Item( 
            /* [in] */ VARIANT ItemSpecifier,
            /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE AddItems( 
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ItemIDs,
            /* [in] */ VARIANT ActiveStates,
            /* [in] */ VARIANT ClientHandles,
            /* [out] */ VARIANT __RPC_FAR *pServerHandles,
            /* [out] */ VARIANT __RPC_FAR *errors,
            /* [out] */ VARIANT __RPC_FAR *pItemObjects,
            /* [optional][in] */ VARIANT AccessPaths,
            /* [optional][in] */ VARIANT RequestedDataTypes,
            /* [optional][out][in] */ VARIANT __RPC_FAR *pBlobs,
            /* [optional][out] */ VARIANT __RPC_FAR *pCanonicalDataTypes,
            /* [optional][out] */ VARIANT __RPC_FAR *pAccessRights) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE ValidateItems( 
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ItemIDs,
            /* [out] */ VARIANT __RPC_FAR *errors,
            /* [optional][in] */ VARIANT AccessPaths,
            /* [optional][in] */ VARIANT RequestedDataTypes,
            /* [optional][in] */ VARIANT BlobUpdate,
            /* [optional][out][in] */ VARIANT __RPC_FAR *pBlobs,
            /* [optional][out] */ VARIANT __RPC_FAR *pCanonicalDataTypes,
            /* [optional][out] */ VARIANT __RPC_FAR *pAccessRights) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE RemoveItems( 
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [out] */ VARIANT __RPC_FAR *errors,
            /* [in] */ VARIANT_BOOL Force) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetActiveState( 
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [in] */ VARIANT_BOOL ActiveState,
            /* [out] */ VARIANT __RPC_FAR *errors) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetClientHandles( 
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [in] */ VARIANT ClientHandles,
            /* [out] */ VARIANT __RPC_FAR *errors) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE SetDatatypes( 
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [in] */ VARIANT RequestedDataTypes,
            /* [out] */ VARIANT __RPC_FAR *errors) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOPCItemMgtDispVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOPCItemMgtDisp __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOPCItemMgtDisp __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pCount);
        
        /* [id][restricted][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Item )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [in] */ VARIANT ItemSpecifier,
            /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddItems )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ItemIDs,
            /* [in] */ VARIANT ActiveStates,
            /* [in] */ VARIANT ClientHandles,
            /* [out] */ VARIANT __RPC_FAR *pServerHandles,
            /* [out] */ VARIANT __RPC_FAR *errors,
            /* [out] */ VARIANT __RPC_FAR *pItemObjects,
            /* [optional][in] */ VARIANT AccessPaths,
            /* [optional][in] */ VARIANT RequestedDataTypes,
            /* [optional][out][in] */ VARIANT __RPC_FAR *pBlobs,
            /* [optional][out] */ VARIANT __RPC_FAR *pCanonicalDataTypes,
            /* [optional][out] */ VARIANT __RPC_FAR *pAccessRights);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ValidateItems )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ItemIDs,
            /* [out] */ VARIANT __RPC_FAR *errors,
            /* [optional][in] */ VARIANT AccessPaths,
            /* [optional][in] */ VARIANT RequestedDataTypes,
            /* [optional][in] */ VARIANT BlobUpdate,
            /* [optional][out][in] */ VARIANT __RPC_FAR *pBlobs,
            /* [optional][out] */ VARIANT __RPC_FAR *pCanonicalDataTypes,
            /* [optional][out] */ VARIANT __RPC_FAR *pAccessRights);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RemoveItems )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [out] */ VARIANT __RPC_FAR *errors,
            /* [in] */ VARIANT_BOOL Force);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetActiveState )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [in] */ VARIANT_BOOL ActiveState,
            /* [out] */ VARIANT __RPC_FAR *errors);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetClientHandles )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [in] */ VARIANT ClientHandles,
            /* [out] */ VARIANT __RPC_FAR *errors);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetDatatypes )( 
            IOPCItemMgtDisp __RPC_FAR * This,
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [in] */ VARIANT RequestedDataTypes,
            /* [out] */ VARIANT __RPC_FAR *errors);
        
        END_INTERFACE
    } IOPCItemMgtDispVtbl;

    interface IOPCItemMgtDisp
    {
        CONST_VTBL struct IOPCItemMgtDispVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOPCItemMgtDisp_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOPCItemMgtDisp_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOPCItemMgtDisp_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOPCItemMgtDisp_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IOPCItemMgtDisp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IOPCItemMgtDisp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IOPCItemMgtDisp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IOPCItemMgtDisp_get_Count(This,pCount)	\
    (This)->lpVtbl -> get_Count(This,pCount)

#define IOPCItemMgtDisp_get__NewEnum(This,ppUnk)	\
    (This)->lpVtbl -> get__NewEnum(This,ppUnk)

#define IOPCItemMgtDisp_Item(This,ItemSpecifier,ppDisp)	\
    (This)->lpVtbl -> Item(This,ItemSpecifier,ppDisp)

#define IOPCItemMgtDisp_AddItems(This,NumItems,ItemIDs,ActiveStates,ClientHandles,pServerHandles,errors,pItemObjects,AccessPaths,RequestedDataTypes,pBlobs,pCanonicalDataTypes,pAccessRights)	\
    (This)->lpVtbl -> AddItems(This,NumItems,ItemIDs,ActiveStates,ClientHandles,pServerHandles,errors,pItemObjects,AccessPaths,RequestedDataTypes,pBlobs,pCanonicalDataTypes,pAccessRights)

#define IOPCItemMgtDisp_ValidateItems(This,NumItems,ItemIDs,errors,AccessPaths,RequestedDataTypes,BlobUpdate,pBlobs,pCanonicalDataTypes,pAccessRights)	\
    (This)->lpVtbl -> ValidateItems(This,NumItems,ItemIDs,errors,AccessPaths,RequestedDataTypes,BlobUpdate,pBlobs,pCanonicalDataTypes,pAccessRights)

#define IOPCItemMgtDisp_RemoveItems(This,NumItems,ServerHandles,errors,Force)	\
    (This)->lpVtbl -> RemoveItems(This,NumItems,ServerHandles,errors,Force)

#define IOPCItemMgtDisp_SetActiveState(This,NumItems,ServerHandles,ActiveState,errors)	\
    (This)->lpVtbl -> SetActiveState(This,NumItems,ServerHandles,ActiveState,errors)

#define IOPCItemMgtDisp_SetClientHandles(This,NumItems,ServerHandles,ClientHandles,errors)	\
    (This)->lpVtbl -> SetClientHandles(This,NumItems,ServerHandles,ClientHandles,errors)

#define IOPCItemMgtDisp_SetDatatypes(This,NumItems,ServerHandles,RequestedDataTypes,errors)	\
    (This)->lpVtbl -> SetDatatypes(This,NumItems,ServerHandles,RequestedDataTypes,errors)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemMgtDisp_get_Count_Proxy( 
    IOPCItemMgtDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pCount);


void __RPC_STUB IOPCItemMgtDisp_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][restricted][propget] */ HRESULT STDMETHODCALLTYPE IOPCItemMgtDisp_get__NewEnum_Proxy( 
    IOPCItemMgtDisp __RPC_FAR * This,
    /* [retval][out] */ IUnknown __RPC_FAR *__RPC_FAR *ppUnk);


void __RPC_STUB IOPCItemMgtDisp_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCItemMgtDisp_Item_Proxy( 
    IOPCItemMgtDisp __RPC_FAR * This,
    /* [in] */ VARIANT ItemSpecifier,
    /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp);


void __RPC_STUB IOPCItemMgtDisp_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCItemMgtDisp_AddItems_Proxy( 
    IOPCItemMgtDisp __RPC_FAR * This,
    /* [in] */ long NumItems,
    /* [in] */ VARIANT ItemIDs,
    /* [in] */ VARIANT ActiveStates,
    /* [in] */ VARIANT ClientHandles,
    /* [out] */ VARIANT __RPC_FAR *pServerHandles,
    /* [out] */ VARIANT __RPC_FAR *errors,
    /* [out] */ VARIANT __RPC_FAR *pItemObjects,
    /* [optional][in] */ VARIANT AccessPaths,
    /* [optional][in] */ VARIANT RequestedDataTypes,
    /* [optional][out][in] */ VARIANT __RPC_FAR *pBlobs,
    /* [optional][out] */ VARIANT __RPC_FAR *pCanonicalDataTypes,
    /* [optional][out] */ VARIANT __RPC_FAR *pAccessRights);


void __RPC_STUB IOPCItemMgtDisp_AddItems_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCItemMgtDisp_ValidateItems_Proxy( 
    IOPCItemMgtDisp __RPC_FAR * This,
    /* [in] */ long NumItems,
    /* [in] */ VARIANT ItemIDs,
    /* [out] */ VARIANT __RPC_FAR *errors,
    /* [optional][in] */ VARIANT AccessPaths,
    /* [optional][in] */ VARIANT RequestedDataTypes,
    /* [optional][in] */ VARIANT BlobUpdate,
    /* [optional][out][in] */ VARIANT __RPC_FAR *pBlobs,
    /* [optional][out] */ VARIANT __RPC_FAR *pCanonicalDataTypes,
    /* [optional][out] */ VARIANT __RPC_FAR *pAccessRights);


void __RPC_STUB IOPCItemMgtDisp_ValidateItems_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCItemMgtDisp_RemoveItems_Proxy( 
    IOPCItemMgtDisp __RPC_FAR * This,
    /* [in] */ long NumItems,
    /* [in] */ VARIANT ServerHandles,
    /* [out] */ VARIANT __RPC_FAR *errors,
    /* [in] */ VARIANT_BOOL Force);


void __RPC_STUB IOPCItemMgtDisp_RemoveItems_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCItemMgtDisp_SetActiveState_Proxy( 
    IOPCItemMgtDisp __RPC_FAR * This,
    /* [in] */ long NumItems,
    /* [in] */ VARIANT ServerHandles,
    /* [in] */ VARIANT_BOOL ActiveState,
    /* [out] */ VARIANT __RPC_FAR *errors);


void __RPC_STUB IOPCItemMgtDisp_SetActiveState_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCItemMgtDisp_SetClientHandles_Proxy( 
    IOPCItemMgtDisp __RPC_FAR * This,
    /* [in] */ long NumItems,
    /* [in] */ VARIANT ServerHandles,
    /* [in] */ VARIANT ClientHandles,
    /* [out] */ VARIANT __RPC_FAR *errors);


void __RPC_STUB IOPCItemMgtDisp_SetClientHandles_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCItemMgtDisp_SetDatatypes_Proxy( 
    IOPCItemMgtDisp __RPC_FAR * This,
    /* [in] */ long NumItems,
    /* [in] */ VARIANT ServerHandles,
    /* [in] */ VARIANT RequestedDataTypes,
    /* [out] */ VARIANT __RPC_FAR *errors);


void __RPC_STUB IOPCItemMgtDisp_SetDatatypes_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOPCItemMgtDisp_INTERFACE_DEFINED__ */


#ifndef __IOPCGroupStateMgtDisp_INTERFACE_DEFINED__
#define __IOPCGroupStateMgtDisp_INTERFACE_DEFINED__

/****************************************
 * Generated header for interface: IOPCGroupStateMgtDisp
 * at Tue Mar 02 14:11:12 1999
 * using MIDL 3.01.75
 ****************************************/
/* [object][oleautomation][dual][uuid] */ 



EXTERN_C const IID IID_IOPCGroupStateMgtDisp;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface DECLSPEC_UUID("39c13a5b-011e-11d0-9675-0020afd8adb3")
    IOPCGroupStateMgtDisp : public IDispatch
    {
    public:
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_ActiveStatus( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pActiveStatus) = 0;
        
        virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_ActiveStatus( 
            /* [in] */ VARIANT_BOOL ActiveStatus) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_ClientGroupHandle( 
            /* [retval][out] */ long __RPC_FAR *phClientGroupHandle) = 0;
        
        virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_ClientGroupHandle( 
            /* [in] */ long ClientGroupHandle) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_ServerGroupHandle( 
            /* [retval][out] */ long __RPC_FAR *phServerGroupHandle) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR __RPC_FAR *pName) = 0;
        
        virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_Name( 
            /* [in] */ BSTR Name) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_UpdateRate( 
            /* [retval][out] */ long __RPC_FAR *pUpdateRate) = 0;
        
        virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_UpdateRate( 
            /* [in] */ long UpdateRate) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_TimeBias( 
            /* [retval][out] */ long __RPC_FAR *pTimeBias) = 0;
        
        virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_TimeBias( 
            /* [in] */ long TimeBias) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_PercentDeadBand( 
            /* [retval][out] */ float __RPC_FAR *pPercentDeadBand) = 0;
        
        virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_PercentDeadBand( 
            /* [in] */ float PercentDeadBand) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_LCID( 
            /* [retval][out] */ long __RPC_FAR *pLCID) = 0;
        
        virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_LCID( 
            /* [in] */ long LCID) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE CloneGroup( 
            /* [optional][in] */ VARIANT Name,
            /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOPCGroupStateMgtDispVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ActiveStatus )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pActiveStatus);
        
        /* [propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ActiveStatus )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL ActiveStatus);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ClientGroupHandle )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *phClientGroupHandle);
        
        /* [propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ClientGroupHandle )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ long ClientGroupHandle);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ServerGroupHandle )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *phServerGroupHandle);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Name )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pName);
        
        /* [propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Name )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ BSTR Name);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_UpdateRate )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pUpdateRate);
        
        /* [propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_UpdateRate )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ long UpdateRate);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TimeBias )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pTimeBias);
        
        /* [propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_TimeBias )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ long TimeBias);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_PercentDeadBand )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [retval][out] */ float __RPC_FAR *pPercentDeadBand);
        
        /* [propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_PercentDeadBand )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ float PercentDeadBand);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LCID )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pLCID);
        
        /* [propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_LCID )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ long LCID);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CloneGroup )( 
            IOPCGroupStateMgtDisp __RPC_FAR * This,
            /* [optional][in] */ VARIANT Name,
            /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp);
        
        END_INTERFACE
    } IOPCGroupStateMgtDispVtbl;

    interface IOPCGroupStateMgtDisp
    {
        CONST_VTBL struct IOPCGroupStateMgtDispVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOPCGroupStateMgtDisp_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOPCGroupStateMgtDisp_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOPCGroupStateMgtDisp_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOPCGroupStateMgtDisp_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IOPCGroupStateMgtDisp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IOPCGroupStateMgtDisp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IOPCGroupStateMgtDisp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IOPCGroupStateMgtDisp_get_ActiveStatus(This,pActiveStatus)	\
    (This)->lpVtbl -> get_ActiveStatus(This,pActiveStatus)

#define IOPCGroupStateMgtDisp_put_ActiveStatus(This,ActiveStatus)	\
    (This)->lpVtbl -> put_ActiveStatus(This,ActiveStatus)

#define IOPCGroupStateMgtDisp_get_ClientGroupHandle(This,phClientGroupHandle)	\
    (This)->lpVtbl -> get_ClientGroupHandle(This,phClientGroupHandle)

#define IOPCGroupStateMgtDisp_put_ClientGroupHandle(This,ClientGroupHandle)	\
    (This)->lpVtbl -> put_ClientGroupHandle(This,ClientGroupHandle)

#define IOPCGroupStateMgtDisp_get_ServerGroupHandle(This,phServerGroupHandle)	\
    (This)->lpVtbl -> get_ServerGroupHandle(This,phServerGroupHandle)

#define IOPCGroupStateMgtDisp_get_Name(This,pName)	\
    (This)->lpVtbl -> get_Name(This,pName)

#define IOPCGroupStateMgtDisp_put_Name(This,Name)	\
    (This)->lpVtbl -> put_Name(This,Name)

#define IOPCGroupStateMgtDisp_get_UpdateRate(This,pUpdateRate)	\
    (This)->lpVtbl -> get_UpdateRate(This,pUpdateRate)

#define IOPCGroupStateMgtDisp_put_UpdateRate(This,UpdateRate)	\
    (This)->lpVtbl -> put_UpdateRate(This,UpdateRate)

#define IOPCGroupStateMgtDisp_get_TimeBias(This,pTimeBias)	\
    (This)->lpVtbl -> get_TimeBias(This,pTimeBias)

#define IOPCGroupStateMgtDisp_put_TimeBias(This,TimeBias)	\
    (This)->lpVtbl -> put_TimeBias(This,TimeBias)

#define IOPCGroupStateMgtDisp_get_PercentDeadBand(This,pPercentDeadBand)	\
    (This)->lpVtbl -> get_PercentDeadBand(This,pPercentDeadBand)

#define IOPCGroupStateMgtDisp_put_PercentDeadBand(This,PercentDeadBand)	\
    (This)->lpVtbl -> put_PercentDeadBand(This,PercentDeadBand)

#define IOPCGroupStateMgtDisp_get_LCID(This,pLCID)	\
    (This)->lpVtbl -> get_LCID(This,pLCID)

#define IOPCGroupStateMgtDisp_put_LCID(This,LCID)	\
    (This)->lpVtbl -> put_LCID(This,LCID)

#define IOPCGroupStateMgtDisp_CloneGroup(This,Name,ppDisp)	\
    (This)->lpVtbl -> CloneGroup(This,Name,ppDisp)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_get_ActiveStatus_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pActiveStatus);


void __RPC_STUB IOPCGroupStateMgtDisp_get_ActiveStatus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_put_ActiveStatus_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL ActiveStatus);


void __RPC_STUB IOPCGroupStateMgtDisp_put_ActiveStatus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_get_ClientGroupHandle_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *phClientGroupHandle);


void __RPC_STUB IOPCGroupStateMgtDisp_get_ClientGroupHandle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_put_ClientGroupHandle_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [in] */ long ClientGroupHandle);


void __RPC_STUB IOPCGroupStateMgtDisp_put_ClientGroupHandle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_get_ServerGroupHandle_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *phServerGroupHandle);


void __RPC_STUB IOPCGroupStateMgtDisp_get_ServerGroupHandle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_get_Name_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pName);


void __RPC_STUB IOPCGroupStateMgtDisp_get_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_put_Name_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [in] */ BSTR Name);


void __RPC_STUB IOPCGroupStateMgtDisp_put_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_get_UpdateRate_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pUpdateRate);


void __RPC_STUB IOPCGroupStateMgtDisp_get_UpdateRate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_put_UpdateRate_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [in] */ long UpdateRate);


void __RPC_STUB IOPCGroupStateMgtDisp_put_UpdateRate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_get_TimeBias_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pTimeBias);


void __RPC_STUB IOPCGroupStateMgtDisp_get_TimeBias_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_put_TimeBias_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [in] */ long TimeBias);


void __RPC_STUB IOPCGroupStateMgtDisp_put_TimeBias_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_get_PercentDeadBand_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [retval][out] */ float __RPC_FAR *pPercentDeadBand);


void __RPC_STUB IOPCGroupStateMgtDisp_get_PercentDeadBand_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_put_PercentDeadBand_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [in] */ float PercentDeadBand);


void __RPC_STUB IOPCGroupStateMgtDisp_put_PercentDeadBand_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_get_LCID_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pLCID);


void __RPC_STUB IOPCGroupStateMgtDisp_get_LCID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput] */ HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_put_LCID_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [in] */ long LCID);


void __RPC_STUB IOPCGroupStateMgtDisp_put_LCID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCGroupStateMgtDisp_CloneGroup_Proxy( 
    IOPCGroupStateMgtDisp __RPC_FAR * This,
    /* [optional][in] */ VARIANT Name,
    /* [retval][out] */ IDispatch __RPC_FAR *__RPC_FAR *ppDisp);


void __RPC_STUB IOPCGroupStateMgtDisp_CloneGroup_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOPCGroupStateMgtDisp_INTERFACE_DEFINED__ */


#ifndef __IOPCSyncIODisp_INTERFACE_DEFINED__
#define __IOPCSyncIODisp_INTERFACE_DEFINED__

/****************************************
 * Generated header for interface: IOPCSyncIODisp
 * at Tue Mar 02 14:11:12 1999
 * using MIDL 3.01.75
 ****************************************/
/* [object][unique][helpstring][dual][uuid] */ 



EXTERN_C const IID IID_IOPCSyncIODisp;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface DECLSPEC_UUID("39c13a5c-011e-11d0-9675-0020afd8adb3")
    IOPCSyncIODisp : public IDispatch
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE OPCRead( 
            /* [in] */ short Source,
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [out] */ VARIANT __RPC_FAR *pValues,
            /* [optional][out] */ VARIANT __RPC_FAR *pQualities,
            /* [optional][out] */ VARIANT __RPC_FAR *pTimeStamps,
            /* [optional][out] */ VARIANT __RPC_FAR *errors) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE OPCWrite( 
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [in] */ VARIANT Values,
            /* [optional][out] */ VARIANT __RPC_FAR *errors) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOPCSyncIODispVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOPCSyncIODisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOPCSyncIODisp __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOPCSyncIODisp __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IOPCSyncIODisp __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IOPCSyncIODisp __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IOPCSyncIODisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IOPCSyncIODisp __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OPCRead )( 
            IOPCSyncIODisp __RPC_FAR * This,
            /* [in] */ short Source,
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [out] */ VARIANT __RPC_FAR *pValues,
            /* [optional][out] */ VARIANT __RPC_FAR *pQualities,
            /* [optional][out] */ VARIANT __RPC_FAR *pTimeStamps,
            /* [optional][out] */ VARIANT __RPC_FAR *errors);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OPCWrite )( 
            IOPCSyncIODisp __RPC_FAR * This,
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [in] */ VARIANT Values,
            /* [optional][out] */ VARIANT __RPC_FAR *errors);
        
        END_INTERFACE
    } IOPCSyncIODispVtbl;

    interface IOPCSyncIODisp
    {
        CONST_VTBL struct IOPCSyncIODispVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOPCSyncIODisp_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOPCSyncIODisp_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOPCSyncIODisp_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOPCSyncIODisp_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IOPCSyncIODisp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IOPCSyncIODisp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IOPCSyncIODisp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IOPCSyncIODisp_OPCRead(This,Source,NumItems,ServerHandles,pValues,pQualities,pTimeStamps,errors)	\
    (This)->lpVtbl -> OPCRead(This,Source,NumItems,ServerHandles,pValues,pQualities,pTimeStamps,errors)

#define IOPCSyncIODisp_OPCWrite(This,NumItems,ServerHandles,Values,errors)	\
    (This)->lpVtbl -> OPCWrite(This,NumItems,ServerHandles,Values,errors)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IOPCSyncIODisp_OPCRead_Proxy( 
    IOPCSyncIODisp __RPC_FAR * This,
    /* [in] */ short Source,
    /* [in] */ long NumItems,
    /* [in] */ VARIANT ServerHandles,
    /* [out] */ VARIANT __RPC_FAR *pValues,
    /* [optional][out] */ VARIANT __RPC_FAR *pQualities,
    /* [optional][out] */ VARIANT __RPC_FAR *pTimeStamps,
    /* [optional][out] */ VARIANT __RPC_FAR *errors);


void __RPC_STUB IOPCSyncIODisp_OPCRead_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCSyncIODisp_OPCWrite_Proxy( 
    IOPCSyncIODisp __RPC_FAR * This,
    /* [in] */ long NumItems,
    /* [in] */ VARIANT ServerHandles,
    /* [in] */ VARIANT Values,
    /* [optional][out] */ VARIANT __RPC_FAR *errors);


void __RPC_STUB IOPCSyncIODisp_OPCWrite_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOPCSyncIODisp_INTERFACE_DEFINED__ */


#ifndef __IOPCAsyncIODisp_INTERFACE_DEFINED__
#define __IOPCAsyncIODisp_INTERFACE_DEFINED__

/****************************************
 * Generated header for interface: IOPCAsyncIODisp
 * at Tue Mar 02 14:11:12 1999
 * using MIDL 3.01.75
 ****************************************/
/* [object][unique][helpstring][dual][uuid] */ 



EXTERN_C const IID IID_IOPCAsyncIODisp;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface DECLSPEC_UUID("39c13a5d-011e-11d0-9675-0020afd8adb3")
    IOPCAsyncIODisp : public IDispatch
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE AddCallbackReference( 
            /* [in] */ long Context,
            /* [in] */ IDispatch __RPC_FAR *pCallback,
            /* [retval][out] */ long __RPC_FAR *pConnection) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE DropCallbackReference( 
            /* [in] */ long Connection) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE OPCRead( 
            /* [in] */ short Source,
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [optional][out] */ VARIANT __RPC_FAR *errors,
            /* [retval][out] */ long __RPC_FAR *pTransactionID) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE OPCWrite( 
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [in] */ VARIANT Values,
            /* [optional][out] */ VARIANT __RPC_FAR *errors,
            /* [retval][out] */ long __RPC_FAR *pTransactionID) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Cancel( 
            /* [in] */ long TransactionID) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE Refresh( 
            /* [in] */ short Source,
            /* [retval][out] */ long __RPC_FAR *pTransactionID) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOPCAsyncIODispVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOPCAsyncIODisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOPCAsyncIODisp __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOPCAsyncIODisp __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IOPCAsyncIODisp __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IOPCAsyncIODisp __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IOPCAsyncIODisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IOPCAsyncIODisp __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddCallbackReference )( 
            IOPCAsyncIODisp __RPC_FAR * This,
            /* [in] */ long Context,
            /* [in] */ IDispatch __RPC_FAR *pCallback,
            /* [retval][out] */ long __RPC_FAR *pConnection);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DropCallbackReference )( 
            IOPCAsyncIODisp __RPC_FAR * This,
            /* [in] */ long Connection);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OPCRead )( 
            IOPCAsyncIODisp __RPC_FAR * This,
            /* [in] */ short Source,
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [optional][out] */ VARIANT __RPC_FAR *errors,
            /* [retval][out] */ long __RPC_FAR *pTransactionID);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OPCWrite )( 
            IOPCAsyncIODisp __RPC_FAR * This,
            /* [in] */ long NumItems,
            /* [in] */ VARIANT ServerHandles,
            /* [in] */ VARIANT Values,
            /* [optional][out] */ VARIANT __RPC_FAR *errors,
            /* [retval][out] */ long __RPC_FAR *pTransactionID);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Cancel )( 
            IOPCAsyncIODisp __RPC_FAR * This,
            /* [in] */ long TransactionID);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Refresh )( 
            IOPCAsyncIODisp __RPC_FAR * This,
            /* [in] */ short Source,
            /* [retval][out] */ long __RPC_FAR *pTransactionID);
        
        END_INTERFACE
    } IOPCAsyncIODispVtbl;

    interface IOPCAsyncIODisp
    {
        CONST_VTBL struct IOPCAsyncIODispVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOPCAsyncIODisp_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOPCAsyncIODisp_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOPCAsyncIODisp_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOPCAsyncIODisp_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IOPCAsyncIODisp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IOPCAsyncIODisp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IOPCAsyncIODisp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IOPCAsyncIODisp_AddCallbackReference(This,Context,pCallback,pConnection)	\
    (This)->lpVtbl -> AddCallbackReference(This,Context,pCallback,pConnection)

#define IOPCAsyncIODisp_DropCallbackReference(This,Connection)	\
    (This)->lpVtbl -> DropCallbackReference(This,Connection)

#define IOPCAsyncIODisp_OPCRead(This,Source,NumItems,ServerHandles,errors,pTransactionID)	\
    (This)->lpVtbl -> OPCRead(This,Source,NumItems,ServerHandles,errors,pTransactionID)

#define IOPCAsyncIODisp_OPCWrite(This,NumItems,ServerHandles,Values,errors,pTransactionID)	\
    (This)->lpVtbl -> OPCWrite(This,NumItems,ServerHandles,Values,errors,pTransactionID)

#define IOPCAsyncIODisp_Cancel(This,TransactionID)	\
    (This)->lpVtbl -> Cancel(This,TransactionID)

#define IOPCAsyncIODisp_Refresh(This,Source,pTransactionID)	\
    (This)->lpVtbl -> Refresh(This,Source,pTransactionID)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT STDMETHODCALLTYPE IOPCAsyncIODisp_AddCallbackReference_Proxy( 
    IOPCAsyncIODisp __RPC_FAR * This,
    /* [in] */ long Context,
    /* [in] */ IDispatch __RPC_FAR *pCallback,
    /* [retval][out] */ long __RPC_FAR *pConnection);


void __RPC_STUB IOPCAsyncIODisp_AddCallbackReference_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCAsyncIODisp_DropCallbackReference_Proxy( 
    IOPCAsyncIODisp __RPC_FAR * This,
    /* [in] */ long Connection);


void __RPC_STUB IOPCAsyncIODisp_DropCallbackReference_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCAsyncIODisp_OPCRead_Proxy( 
    IOPCAsyncIODisp __RPC_FAR * This,
    /* [in] */ short Source,
    /* [in] */ long NumItems,
    /* [in] */ VARIANT ServerHandles,
    /* [optional][out] */ VARIANT __RPC_FAR *errors,
    /* [retval][out] */ long __RPC_FAR *pTransactionID);


void __RPC_STUB IOPCAsyncIODisp_OPCRead_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCAsyncIODisp_OPCWrite_Proxy( 
    IOPCAsyncIODisp __RPC_FAR * This,
    /* [in] */ long NumItems,
    /* [in] */ VARIANT ServerHandles,
    /* [in] */ VARIANT Values,
    /* [optional][out] */ VARIANT __RPC_FAR *errors,
    /* [retval][out] */ long __RPC_FAR *pTransactionID);


void __RPC_STUB IOPCAsyncIODisp_OPCWrite_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCAsyncIODisp_Cancel_Proxy( 
    IOPCAsyncIODisp __RPC_FAR * This,
    /* [in] */ long TransactionID);


void __RPC_STUB IOPCAsyncIODisp_Cancel_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCAsyncIODisp_Refresh_Proxy( 
    IOPCAsyncIODisp __RPC_FAR * This,
    /* [in] */ short Source,
    /* [retval][out] */ long __RPC_FAR *pTransactionID);


void __RPC_STUB IOPCAsyncIODisp_Refresh_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOPCAsyncIODisp_INTERFACE_DEFINED__ */


#ifndef __IOPCPublicGroupStateMgtDisp_INTERFACE_DEFINED__
#define __IOPCPublicGroupStateMgtDisp_INTERFACE_DEFINED__

/****************************************
 * Generated header for interface: IOPCPublicGroupStateMgtDisp
 * at Tue Mar 02 14:11:12 1999
 * using MIDL 3.01.75
 ****************************************/
/* [object][unique][helpstring][dual][uuid] */ 



EXTERN_C const IID IID_IOPCPublicGroupStateMgtDisp;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface DECLSPEC_UUID("39c13a5e-011e-11d0-9675-0020afd8adb3")
    IOPCPublicGroupStateMgtDisp : public IDispatch
    {
    public:
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_State( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pPublic) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE MoveToPublic_DISP( void) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOPCPublicGroupStateMgtDispVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOPCPublicGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOPCPublicGroupStateMgtDisp __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOPCPublicGroupStateMgtDisp __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IOPCPublicGroupStateMgtDisp __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IOPCPublicGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IOPCPublicGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IOPCPublicGroupStateMgtDisp __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_State )( 
            IOPCPublicGroupStateMgtDisp __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pPublic);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *MoveToPublic_DISP )( 
            IOPCPublicGroupStateMgtDisp __RPC_FAR * This);
        
        END_INTERFACE
    } IOPCPublicGroupStateMgtDispVtbl;

    interface IOPCPublicGroupStateMgtDisp
    {
        CONST_VTBL struct IOPCPublicGroupStateMgtDispVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOPCPublicGroupStateMgtDisp_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOPCPublicGroupStateMgtDisp_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOPCPublicGroupStateMgtDisp_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOPCPublicGroupStateMgtDisp_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IOPCPublicGroupStateMgtDisp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IOPCPublicGroupStateMgtDisp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IOPCPublicGroupStateMgtDisp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IOPCPublicGroupStateMgtDisp_get_State(This,pPublic)	\
    (This)->lpVtbl -> get_State(This,pPublic)

#define IOPCPublicGroupStateMgtDisp_MoveToPublic(This)	\
    (This)->lpVtbl -> MoveToPublic_DISP(This)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCPublicGroupStateMgtDisp_get_State_Proxy( 
    IOPCPublicGroupStateMgtDisp __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pPublic);


void __RPC_STUB IOPCPublicGroupStateMgtDisp_get_State_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCPublicGroupStateMgtDisp_MoveToPublic_Proxy( 
    IOPCPublicGroupStateMgtDisp __RPC_FAR * This);


void __RPC_STUB IOPCPublicGroupStateMgtDisp_MoveToPublic_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOPCPublicGroupStateMgtDisp_INTERFACE_DEFINED__ */


#ifndef __IOPCItemDisp_INTERFACE_DEFINED__
#define __IOPCItemDisp_INTERFACE_DEFINED__

/****************************************
 * Generated header for interface: IOPCItemDisp
 * at Tue Mar 02 14:11:12 1999
 * using MIDL 3.01.75
 ****************************************/
/* [object][unique][helpstring][dual][uuid] */ 



EXTERN_C const IID IID_IOPCItemDisp;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    interface DECLSPEC_UUID("39c13a5f-011e-11d0-9675-0020afd8adb3")
    IOPCItemDisp : public IDispatch
    {
    public:
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_AccessPath( 
            /* [retval][out] */ BSTR __RPC_FAR *pAccessPath) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_AccessRights( 
            /* [retval][out] */ long __RPC_FAR *pAccessRights) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_ActiveStatus( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pActiveStatus) = 0;
        
        virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_ActiveStatus( 
            /* [in] */ VARIANT_BOOL ActiveStatus) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Blob( 
            /* [retval][out] */ VARIANT __RPC_FAR *pBlob) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_ClientHandle( 
            /* [retval][out] */ long __RPC_FAR *phClient) = 0;
        
        virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_ClientHandle( 
            /* [in] */ long Client) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_ItemID( 
            /* [retval][out] */ BSTR __RPC_FAR *pItemID) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_ServerHandle( 
            /* [retval][out] */ long __RPC_FAR *pServer) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_RequestedDataType( 
            /* [retval][out] */ short __RPC_FAR *pRequestedDataType) = 0;
        
        virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_RequestedDataType( 
            /* [in] */ short RequestedDataType) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Value( 
            /* [retval][out] */ VARIANT __RPC_FAR *pValue) = 0;
        
        virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_Value( 
            /* [in] */ VARIANT NewValue) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Quality( 
            /* [retval][out] */ short __RPC_FAR *pQuality) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Timestamp( 
            /* [retval][out] */ DATE __RPC_FAR *pTimeStamp) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_ReadError( 
            /* [retval][out] */ long __RPC_FAR *pError) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_WriteError( 
            /* [retval][out] */ long __RPC_FAR *pError) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_EUType( 
            /* [retval][out] */ short __RPC_FAR *pError) = 0;
        
        virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_EUInfo( 
            /* [retval][out] */ VARIANT __RPC_FAR *pError) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE OPCRead( 
            /* [in] */ short Source,
            /* [out] */ VARIANT __RPC_FAR *pValue,
            /* [optional][out] */ VARIANT __RPC_FAR *pQuality,
            /* [optional][out] */ VARIANT __RPC_FAR *pTimestamp) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE OPCWrite( 
            /* [in] */ VARIANT NewValue) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOPCItemDispVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOPCItemDisp __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOPCItemDisp __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_AccessPath )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pAccessPath);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_AccessRights )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pAccessRights);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ActiveStatus )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pActiveStatus);
        
        /* [propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ActiveStatus )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL ActiveStatus);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Blob )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ VARIANT __RPC_FAR *pBlob);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ClientHandle )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *phClient);
        
        /* [propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ClientHandle )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [in] */ long Client);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ItemID )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pItemID);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ServerHandle )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pServer);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_RequestedDataType )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pRequestedDataType);
        
        /* [propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_RequestedDataType )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [in] */ short RequestedDataType);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Value )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ VARIANT __RPC_FAR *pValue);
        
        /* [propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Value )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [in] */ VARIANT NewValue);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Quality )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pQuality);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Timestamp )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ DATE __RPC_FAR *pTimeStamp);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ReadError )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pError);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_WriteError )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pError);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EUType )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pError);
        
        /* [propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EUInfo )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [retval][out] */ VARIANT __RPC_FAR *pError);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OPCRead )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [in] */ short Source,
            /* [out] */ VARIANT __RPC_FAR *pValue,
            /* [optional][out] */ VARIANT __RPC_FAR *pQuality,
            /* [optional][out] */ VARIANT __RPC_FAR *pTimestamp);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OPCWrite )( 
            IOPCItemDisp __RPC_FAR * This,
            /* [in] */ VARIANT NewValue);
        
        END_INTERFACE
    } IOPCItemDispVtbl;

    interface IOPCItemDisp
    {
        CONST_VTBL struct IOPCItemDispVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOPCItemDisp_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOPCItemDisp_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOPCItemDisp_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOPCItemDisp_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IOPCItemDisp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IOPCItemDisp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IOPCItemDisp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IOPCItemDisp_get_AccessPath(This,pAccessPath)	\
    (This)->lpVtbl -> get_AccessPath(This,pAccessPath)

#define IOPCItemDisp_get_AccessRights(This,pAccessRights)	\
    (This)->lpVtbl -> get_AccessRights(This,pAccessRights)

#define IOPCItemDisp_get_ActiveStatus(This,pActiveStatus)	\
    (This)->lpVtbl -> get_ActiveStatus(This,pActiveStatus)

#define IOPCItemDisp_put_ActiveStatus(This,ActiveStatus)	\
    (This)->lpVtbl -> put_ActiveStatus(This,ActiveStatus)

#define IOPCItemDisp_get_Blob(This,pBlob)	\
    (This)->lpVtbl -> get_Blob(This,pBlob)

#define IOPCItemDisp_get_ClientHandle(This,phClient)	\
    (This)->lpVtbl -> get_ClientHandle(This,phClient)

#define IOPCItemDisp_put_ClientHandle(This,Client)	\
    (This)->lpVtbl -> put_ClientHandle(This,Client)

#define IOPCItemDisp_get_ItemID(This,pItemID)	\
    (This)->lpVtbl -> get_ItemID(This,pItemID)

#define IOPCItemDisp_get_ServerHandle(This,pServer)	\
    (This)->lpVtbl -> get_ServerHandle(This,pServer)

#define IOPCItemDisp_get_RequestedDataType(This,pRequestedDataType)	\
    (This)->lpVtbl -> get_RequestedDataType(This,pRequestedDataType)

#define IOPCItemDisp_put_RequestedDataType(This,RequestedDataType)	\
    (This)->lpVtbl -> put_RequestedDataType(This,RequestedDataType)

#define IOPCItemDisp_get_Value(This,pValue)	\
    (This)->lpVtbl -> get_Value(This,pValue)

#define IOPCItemDisp_put_Value(This,NewValue)	\
    (This)->lpVtbl -> put_Value(This,NewValue)

#define IOPCItemDisp_get_Quality(This,pQuality)	\
    (This)->lpVtbl -> get_Quality(This,pQuality)

#define IOPCItemDisp_get_Timestamp(This,pTimeStamp)	\
    (This)->lpVtbl -> get_Timestamp(This,pTimeStamp)

#define IOPCItemDisp_get_ReadError(This,pError)	\
    (This)->lpVtbl -> get_ReadError(This,pError)

#define IOPCItemDisp_get_WriteError(This,pError)	\
    (This)->lpVtbl -> get_WriteError(This,pError)

#define IOPCItemDisp_get_EUType(This,pError)	\
    (This)->lpVtbl -> get_EUType(This,pError)

#define IOPCItemDisp_get_EUInfo(This,pError)	\
    (This)->lpVtbl -> get_EUInfo(This,pError)

#define IOPCItemDisp_OPCRead(This,Source,pValue,pQuality,pTimestamp)	\
    (This)->lpVtbl -> OPCRead(This,Source,pValue,pQuality,pTimestamp)

#define IOPCItemDisp_OPCWrite(This,NewValue)	\
    (This)->lpVtbl -> OPCWrite(This,NewValue)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_AccessPath_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pAccessPath);


void __RPC_STUB IOPCItemDisp_get_AccessPath_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_AccessRights_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pAccessRights);


void __RPC_STUB IOPCItemDisp_get_AccessRights_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_ActiveStatus_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pActiveStatus);


void __RPC_STUB IOPCItemDisp_get_ActiveStatus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_put_ActiveStatus_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL ActiveStatus);


void __RPC_STUB IOPCItemDisp_put_ActiveStatus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_Blob_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ VARIANT __RPC_FAR *pBlob);


void __RPC_STUB IOPCItemDisp_get_Blob_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_ClientHandle_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *phClient);


void __RPC_STUB IOPCItemDisp_get_ClientHandle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_put_ClientHandle_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [in] */ long Client);


void __RPC_STUB IOPCItemDisp_put_ClientHandle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_ItemID_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pItemID);


void __RPC_STUB IOPCItemDisp_get_ItemID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_ServerHandle_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pServer);


void __RPC_STUB IOPCItemDisp_get_ServerHandle_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_RequestedDataType_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pRequestedDataType);


void __RPC_STUB IOPCItemDisp_get_RequestedDataType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_put_RequestedDataType_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [in] */ short RequestedDataType);


void __RPC_STUB IOPCItemDisp_put_RequestedDataType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_Value_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ VARIANT __RPC_FAR *pValue);


void __RPC_STUB IOPCItemDisp_get_Value_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propput] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_put_Value_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [in] */ VARIANT NewValue);


void __RPC_STUB IOPCItemDisp_put_Value_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_Quality_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pQuality);


void __RPC_STUB IOPCItemDisp_get_Quality_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_Timestamp_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ DATE __RPC_FAR *pTimeStamp);


void __RPC_STUB IOPCItemDisp_get_Timestamp_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_ReadError_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pError);


void __RPC_STUB IOPCItemDisp_get_ReadError_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_WriteError_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pError);


void __RPC_STUB IOPCItemDisp_get_WriteError_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_EUType_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pError);


void __RPC_STUB IOPCItemDisp_get_EUType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [propget] */ HRESULT STDMETHODCALLTYPE IOPCItemDisp_get_EUInfo_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [retval][out] */ VARIANT __RPC_FAR *pError);


void __RPC_STUB IOPCItemDisp_get_EUInfo_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCItemDisp_OPCRead_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [in] */ short Source,
    /* [out] */ VARIANT __RPC_FAR *pValue,
    /* [optional][out] */ VARIANT __RPC_FAR *pQuality,
    /* [optional][out] */ VARIANT __RPC_FAR *pTimestamp);


void __RPC_STUB IOPCItemDisp_OPCRead_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT STDMETHODCALLTYPE IOPCItemDisp_OPCWrite_Proxy( 
    IOPCItemDisp __RPC_FAR * This,
    /* [in] */ VARIANT NewValue);


void __RPC_STUB IOPCItemDisp_OPCWrite_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOPCItemDisp_INTERFACE_DEFINED__ */



#ifndef __OPCSDKLib_LIBRARY_DEFINED__
#define __OPCSDKLib_LIBRARY_DEFINED__

/****************************************
 * Generated header for library: OPCSDKLib
 * at Tue Mar 02 14:11:12 1999
 * using MIDL 3.01.75
 ****************************************/
/* [helpstring][version][uuid] */ 



EXTERN_C const IID LIBID_OPCSDKLib;

#ifdef __cplusplus
EXTERN_C const CLSID CLSID_OPCGroup;

class DECLSPEC_UUID("94003F74-FE18-11D0-A265-0000E81E9085")
OPCGroup;
#endif

#ifdef __cplusplus
EXTERN_C const CLSID CLSID_OPCItem;

class DECLSPEC_UUID("94003F76-FE18-11D0-A265-0000E81E9085")
OPCItem;
#endif
#endif /* __OPCSDKLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long __RPC_FAR *, unsigned long            , BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long __RPC_FAR *, BSTR __RPC_FAR * ); 

unsigned long             __RPC_USER  VARIANT_UserSize(     unsigned long __RPC_FAR *, unsigned long            , VARIANT __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  VARIANT_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, VARIANT __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  VARIANT_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, VARIANT __RPC_FAR * ); 
void                      __RPC_USER  VARIANT_UserFree(     unsigned long __RPC_FAR *, VARIANT __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
