/*
 * Copyright (c) 2011-2022 Technosoftware GmbH. All rights reserved
 * Web: https://technosoftware.com 
 * 
 * The source code in this file is covered under a dual-license scenario:
 *   - Owner of a purchased license: SCLA 1.0
 *   - GPL V3: everybody else
 *
 * SCLA license terms accompanied with this source code.
 * See https://technosoftware.com/license/Source_Code_License_Agreement.pdf
 *
 * GNU General Public License as published by the Free Software Foundation;
 * version 3 of the License are accompanied with this source code.
 * See https://technosoftware.com/license/GPLv3License.txt
 *
 * This source code is distributed in the hope that it will be useful, but
 * WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY
 * or FITNESS FOR A PARTICULAR PURPOSE.
 */

#ifndef __OPCGROUP_H
#define __OPCGROUP_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "UtilityDefs.h"

class DaGroup;
class DaGenericGroup;
class DaGenericServer;
class DaBaseServer;


//-----------------------------------------------------------------------
// TYPEDEF
//-----------------------------------------------------------------------
typedef IConnectionPointImpl
                     <  DaGroup,
                        &IID_IOPCDataCallback,
                        CComUnkArray<1>         // Note: There must be a space after CComUnkArray<1> !!!
                     > IOPCGroupConnectionPointImpl;


/////////////////////////////////////////////////////////////////////////////
// DaGroup
class ATL_NO_VTABLE DaGroup :
   public CComObjectRootEx<CComMultiThreadModel>,
   public CComCoClass<DaGroup, &CLSID_OPCGroup>,

                        // OPC 1.0A Automation Interfaces
   public IDispatchImpl<IOPCItemMgtDisp, 
                        &IID_IOPCItemMgtDisp, 
                        &LIBID_OPCSDKLib>,
   public IDispatchImpl<IOPCGroupStateMgtDisp, 
                        &IID_IOPCGroupStateMgtDisp, 
                        &LIBID_OPCSDKLib>,
   public IDispatchImpl<IOPCSyncIODisp, 
                        &IID_IOPCSyncIODisp, 
                        &LIBID_OPCSDKLib>,
   public IDispatchImpl<IOPCAsyncIODisp, 
                        &IID_IOPCAsyncIODisp, 
                        &LIBID_OPCSDKLib>,
   public IDispatchImpl<IOPCPublicGroupStateMgtDisp, 
                        &IID_IOPCPublicGroupStateMgtDisp, 
                        &LIBID_OPCSDKLib>,

                        // OPC 1.0A Custom Interfaces
   //public IOPCGroupStateMgt,         is base class of IOPCGroupStateMgt2
   public IOPCPublicGroupStateMgt,
   //public IOPCSyncIO,                is base class of IOPCSyncIO2
   public IOPCAsyncIO,
   public IOPCItemMgt,
   public IDataObject,

                        // OPC 2.0 Custom Interfaces
   //public IOPCAsyncIO2,              is base class of IOPCAsyncIO3
   public IConnectionPointContainerImpl<DaGroup>,
   public IOPCGroupConnectionPointImpl,

                        // OPC 3.0 Custom Interfaces
   public IOPCGroupStateMgt2,
   public IOPCSyncIO2,
   public IOPCAsyncIO3,
   public IOPCItemDeadbandMgt
/*

   public IOPCItemSamplingMgt                   // [optional]
*/
{
public:
   DaGroup();
   ~DaGroup();

DECLARE_GET_CONTROLLING_UNKNOWN()

BEGIN_COM_MAP(DaGroup)
                        // OPC 1.0A Automation Interfaces
   COM_INTERFACE_ENTRY(IOPCItemMgtDisp)
   COM_INTERFACE_ENTRY(IOPCGroupStateMgtDisp)
   COM_INTERFACE_ENTRY(IOPCSyncIODisp)
   COM_INTERFACE_ENTRY(IOPCAsyncIODisp)
   COM_INTERFACE_ENTRY(IOPCPublicGroupStateMgtDisp)
   COM_INTERFACE_ENTRY2(IDispatch, IOPCItemMgtDisp)

                        // OPC 1.0A Custom Interfaces
   COM_INTERFACE_ENTRY(IOPCGroupStateMgt)
   COM_INTERFACE_ENTRY(IOPCPublicGroupStateMgt)
   COM_INTERFACE_ENTRY(IOPCSyncIO)
   COM_INTERFACE_ENTRY(IOPCAsyncIO)
   COM_INTERFACE_ENTRY(IOPCItemMgt)
   COM_INTERFACE_ENTRY(IDataObject)

                        // OPC 2.0 Custom Interfaces
   COM_INTERFACE_ENTRY(IOPCAsyncIO2)
   COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)

                        // OPC 3.0 Custom Interfaces
   COM_INTERFACE_ENTRY(IOPCGroupStateMgt2)
   COM_INTERFACE_ENTRY(IOPCSyncIO2)
   COM_INTERFACE_ENTRY(IOPCAsyncIO3)
   COM_INTERFACE_ENTRY(IOPCItemDeadbandMgt)
//   COM_INTERFACE_ENTRY(IOPCItemSamplingMgt)      // [optional]
   
   COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
END_COM_MAP()


BEGIN_CONNECTION_POINT_MAP(DaGroup)
   CONNECTION_POINT_ENTRY(IID_IOPCDataCallback)
END_CONNECTION_POINT_MAP()

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+  IOPCItemMgtDisp
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_Count(
        /*[out, retval]*/ long * pCount 
      );


      ////[ propget, 
      ////  restricted, 
      ////  id( DISPID_NEWENUM )
      ////]
      HRESULT STDMETHODCALLTYPE get__NewEnum(    
        /*[out, retval]*/ IUnknown ** ppUnk
      );


      //
      // Methods
      //

      HRESULT STDMETHODCALLTYPE Item(
        /*[in]         */ VARIANT      ItemSpecifier, 
        /*[out, retval]*/ IDispatch ** ppDisp 
      );

      HRESULT STDMETHODCALLTYPE AddItems(
        /*[in]               */ long      NumItems,
        /*[in]               */ VARIANT   ItemIDs,
        /*[in]               */ VARIANT   ActiveStates,
        /*[in]               */ VARIANT   ClientHandles,
        /*[out]              */ VARIANT * pServerHandles,
        /*[out]              */ VARIANT * errors,
        /*[out]              */ VARIANT * pItemObjects,
        /*[in, optional]     */ VARIANT   AccessPaths,
        /*[in, optional]     */ VARIANT   RequestedDataTypes,
        /*[in, out, optional]*/ VARIANT * pBlobs,
        /*[out, optional]    */ VARIANT * pCanonicalDataTypes,
        /*[out, optional]    */ VARIANT * pAccessRights
      );

      HRESULT STDMETHODCALLTYPE ValidateItems(
        /*[in]               */ long      NumItems,
        /*[in]               */ VARIANT   ItemIDs,
        /*[out]              */ VARIANT * errors,
        /*[in, optional]     */ VARIANT   AccessPaths,
        /*[in, optional]     */ VARIANT   RequestedDataTypes,
        /*[in, optional]     */ VARIANT BlobUpdate,
        /*[in, out, optional]*/ VARIANT * pBlobs,
        /*[out, optional]    */ VARIANT * pCanonicalDataTypes,
        /*[out, optional]    */ VARIANT * pAccessRights
      );

      HRESULT STDMETHODCALLTYPE RemoveItems( 
        /*[in] */ long      NumItems,
        /*[in] */ VARIANT   ServerHandles,
        /*[out]*/ VARIANT * errors,
        /*[in] */ VARIANT_BOOL   Force
        );


      HRESULT STDMETHODCALLTYPE SetActiveState(
        /*[in] */ long      NumItems,
        /*[in] */ VARIANT   ServerHandles,
        /*[in] */ VARIANT_BOOL   ActiveState,
        /*[out]*/ VARIANT * errors
        );


      HRESULT STDMETHODCALLTYPE SetClientHandles(
        /*[in] */ long      NumItems,
        /*[in] */ VARIANT   ServerHandles,
        /*[in] */ VARIANT   ClientHandles,
        /*[out]*/ VARIANT * errors
        );


      HRESULT STDMETHODCALLTYPE SetDatatypes(
        /*[in] */ long      NumItems,
        /*[in] */ VARIANT   ServerHandles,
        /*[in] */ VARIANT   RequestedDataTypes,    
        /*[out]*/ VARIANT * errors
      );

     
//+++++++++++++++++++++++++++++++++++++
//+ IOPCGroupStateMgtDisp
//+++++++++++++++++++++++++++++++++++++

     
      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_ActiveStatus(
        /*[out, retval]*/ VARIANT_BOOL * pActiveStatus
      );
        
      ////[ propput ]
      HRESULT STDMETHODCALLTYPE put_ActiveStatus(
        /*[in]*/ VARIANT_BOOL ActiveStatus
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_ClientGroupHandle(
        /*[out, retval]*/ long * phClientGroupHandle
      );

      ////[ propput ]
      HRESULT STDMETHODCALLTYPE put_ClientGroupHandle(
        /*[in]*/ long ClientGroupHandle
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_ServerGroupHandle(
        /*[out, retval]*/ long * phServerGroupHandle
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_Name(
        /*[out, retval]*/ BSTR * pName 
      );

      ////[ propput ]
      HRESULT STDMETHODCALLTYPE put_Name(
        /*[in]*/ BSTR Name 
      );
        
      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_UpdateRate(
        /*[out, retval]*/ long * pUpdateRate
      );

      ////[ propput ]
      HRESULT STDMETHODCALLTYPE put_UpdateRate(
        /*[in]*/ long UpdateRate
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_TimeBias(
        /*[out, retval]*/ long * pTimeBias
      );

      ////[ propput ]
      HRESULT STDMETHODCALLTYPE put_TimeBias(
        /*[in]*/ long TimeBias
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_PercentDeadBand(
        /*[out, retval]*/ float * pPercentDeadBand
      );

      ////[ propput ]
      HRESULT STDMETHODCALLTYPE put_PercentDeadBand(
        /*[in]*/ float PercentDeadBand
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_LCID(
        /*[out, retval]*/ long * pLCID
      );

      ////[ propput ]
      HRESULT STDMETHODCALLTYPE put_LCID(
        /*[in]*/ long LCID
      );

      //
      // Methods
      //

      HRESULT STDMETHODCALLTYPE CloneGroup(
        /*[in, optional]*/ VARIANT      Name,
        /*[out, retval] */ IDispatch ** ppDisp
      );
     
//+++++++++++++++++++++++++++++++++++++
//+ IOPCSyncIODisp
//+++++++++++++++++++++++++++++++++++++

      HRESULT STDMETHODCALLTYPE OPCRead(
        /*[in]           */ short     Source,
        /*[in]           */ long      NumItems,
        /*[in]           */ VARIANT   ServerHandles,
        /*[out]          */ VARIANT * pValues,
        /*[out, optional]*/ VARIANT * pQualities,
        /*[out, optional]*/ VARIANT * pTimeStamps,
        /*[out, optional]*/ VARIANT * errors
      );

      HRESULT STDMETHODCALLTYPE OPCWrite(
        /*[in]           */ long      NumItems,
        /*[in]           */ VARIANT   ServerHandles,
        /*[in]           */ VARIANT   Values,        
        /*[out, optional]*/ VARIANT * errors
      );
     
     
//+++++++++++++++++++++++++++++++++++++
//+ IOPCAsyncIODisp
//+++++++++++++++++++++++++++++++++++++

      HRESULT STDMETHODCALLTYPE AddCallbackReference(
        /*[in]         */ long        Context,
        /*[in]         */ IDispatch * pCallback,
        /*[out, retval]*/ long      * pConnection
      );

      HRESULT STDMETHODCALLTYPE DropCallbackReference(
        /*[in]*/ long Connection
      );

      HRESULT STDMETHODCALLTYPE OPCRead(
        /*[in]           */ short     Source,
        /*[in]           */ long      NumItems,
        /*[in]           */ VARIANT   ServerHandles,
        /*[out, optional]*/ VARIANT * errors,
        /*[out, retval]  */ long    * pTransactionID
      );

      HRESULT STDMETHODCALLTYPE OPCWrite(
        /*[in]           */ long      NumItems,
        /*[in]           */ VARIANT   ServerHandles,
        /*[in]           */ VARIANT   Values,        
        /*[out, optional]*/ VARIANT * errors,
        /*[out, retval]  */ long    * pTransactionID
      );

      HRESULT STDMETHODCALLTYPE Cancel(
        /*[in]*/ long TransactionID
      );

      HRESULT STDMETHODCALLTYPE Refresh(
        /*[in]         */ short   Source,
        /*[out, retval]*/ long  * pTransactionID
      );
     
     
//+++++++++++++++++++++++++++++++++++++
//+ IOPCPublicGroupStateMgtDisp
//+++++++++++++++++++++++++++++++++++++

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_State(
        /*[out, retval]*/ VARIANT_BOOL * pPublic
      );

      //
      // Methods
      //

      HRESULT STDMETHODCALLTYPE MoveToPublic_DISP(
      void
      );


// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
// [[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCGroupStateMgt[[[[[[[[[[[[[[[[[[[[[[[[[[[[
// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]


      HRESULT STDMETHODCALLTYPE GetState(
         /*[out]        */ DWORD     * pUpdateRate, 
         /*[out]        */ BOOL      * pActive, 
         /*[out, string]*/ LPWSTR    * ppName,
         /*[out]        */ LONG      * pTimeBias,
         /*[out]        */ float     * pPercentDeadband,
         /*[out]        */ DWORD     * pLCID,
         /*[out]        */ OPCHANDLE * phClientGroup,
         /*[out]        */ OPCHANDLE * phServerGroup
      );

      HRESULT STDMETHODCALLTYPE SetState( 
         /*[unique, in]*/ DWORD     * pRequestedUpdateRate, 
         /*[out]       */ DWORD     * pRevisedUpdateRate, 
         /*[unique, in]*/ BOOL      * pActive, 
         /*[unique, in]*/ LONG      * pTimeBias,
         /*[unique, in]*/ float     * pPercentDeadband,
         /*[unique, in]*/ DWORD     * pLCID,
         /*[unique, in]*/ OPCHANDLE * phClientGroup
      );

      HRESULT STDMETHODCALLTYPE SetName( 
         /*[in, string]*/ LPCWSTR szName
      );

      HRESULT STDMETHODCALLTYPE CloneGroup(
         /*[in, string]       */ LPCWSTR     szName,
         /*[in]               */ REFIID      riid,
         /*[out, iid_is(riid)]*/ LPUNKNOWN * ppUnk
      );


// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
// [[[[[[[[[[[[[[[[[[[[[[[ IOPCPublicGroupStateMgt[[[[[[[[[[[[[[[[[[[[[[[[[
// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]


      HRESULT STDMETHODCALLTYPE GetState(
         /*[out]*/ BOOL * pPublic
      );

      HRESULT STDMETHODCALLTYPE MoveToPublic(
         void
      );


// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
// [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCSyncIO [[[[[[[[[[[[[[[[[[[[[[[[[[[[[
// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]

      HRESULT STDMETHODCALLTYPE Read(
         /*[in]                       */ OPCDATASOURCE   dwSource,
         /*[in]                       */ DWORD           numItems, 
         /*[in, size_is(numItems)]  */ OPCHANDLE     * phServer, 
         /*[out, size_is(,numItems)]*/ OPCITEMSTATE ** ppItemValues,
         /*[out, size_is(,numItems)]*/ HRESULT      ** ppErrors
      );

      HRESULT STDMETHODCALLTYPE Write(
         /*[in]                       */ DWORD        numItems, 
         /*[in, size_is(numItems)]  */ OPCHANDLE  * phServer, 
         /*[in, size_is(numItems)]  */ VARIANT    * pItemValues, 
         /*[out, size_is(,numItems)]*/ HRESULT   ** ppErrors
      );


// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
// [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCAsyncIO [[[[[[[[[[[[[[[[[[[[[[[[[[[[
// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]


      HRESULT STDMETHODCALLTYPE Read(
         /*[in]                       */ DWORD           dwConnection,
         /*[in]                       */ OPCDATASOURCE   dwSource,
         /*[in]                       */ DWORD           numItems,
         /*[in, size_is(numItems)]  */ OPCHANDLE     * phServer,
         /*[out]                      */ DWORD         * pTransactionID,
         /*[out, size_is(,numItems)]*/ HRESULT      ** ppErrors
      );

      HRESULT STDMETHODCALLTYPE Write(
         /*[in]                       */ DWORD       dwConnection,
         /*[in]                       */ DWORD       numItems, 
         /*[in, size_is(numItems)]  */ OPCHANDLE * phServer,
         /*[in, size_is(numItems)]  */ VARIANT   * pItemValues, 
         /*[out]                      */ DWORD     * pTransactionID,
         /*[out, size_is(,numItems)]*/ HRESULT ** ppErrors
      );

      HRESULT STDMETHODCALLTYPE Refresh(
         /*[in] */ DWORD           dwConnection,
         /*[in] */ OPCDATASOURCE   dwSource,
         /*[out]*/ DWORD         * pTransactionID
      );

      HRESULT STDMETHODCALLTYPE Cancel(
         /*[in]*/ DWORD dwTransactionID
      );


// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
// [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCItemMgt [[[[[[[[[[[[[[[[[[[[[[[[[[[[[
// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]


      HRESULT STDMETHODCALLTYPE AddItems( 
         /*[in]                       */ DWORD            numItems,
         /*[in, size_is(numItems)]  */ OPCITEMDEF     * pItemArray,
         /*[out, size_is(,numItems)]*/ OPCITEMRESULT ** ppAddResults,
         /*[out, size_is(,numItems)]*/ HRESULT       ** ppErrors
      );

      HRESULT STDMETHODCALLTYPE ValidateItems( 
         /*[in]                       */ DWORD             numItems,
         /*[in, size_is(numItems)]  */ OPCITEMDEF      * pItemArray,
         /*[in]                       */ BOOL              bBlobUpdate,
         /*[out, size_is(,numItems)]*/ OPCITEMRESULT  ** ppValidationResults,
         /*[out, size_is(,numItems)]*/ HRESULT        ** ppErrors
      );

      HRESULT STDMETHODCALLTYPE RemoveItems( 
         /*[in]                       */ DWORD        numItems,
         /*[in, size_is(numItems)]  */ OPCHANDLE  * phServer,
         /*[out, size_is(,numItems)]*/ HRESULT   ** ppErrors
      );

      HRESULT STDMETHODCALLTYPE SetActiveState(
         /*[in]                       */ DWORD        numItems,
         /*[in, size_is(numItems)]  */ OPCHANDLE  * phServer,
         /*[in]                       */ BOOL         bActive, 
         /*[out, size_is(,numItems)]*/ HRESULT   ** ppErrors
      );

      HRESULT STDMETHODCALLTYPE SetClientHandles(
         /*[in]                       */ DWORD        numItems,
         /*[in, size_is(numItems)]  */ OPCHANDLE  * phServer,
         /*[in, size_is(numItems)]  */ OPCHANDLE  * phClient,
         /*[out, size_is(,numItems)]*/ HRESULT   ** ppErrors
      );

      HRESULT STDMETHODCALLTYPE SetDatatypes(
         /*[in]                       */ DWORD        numItems,
         /*[in, size_is(numItems)]  */ OPCHANDLE  * phServer,
         /*[in, size_is(numItems)]  */ VARTYPE    * pRequestedDatatypes,
         /*[out, size_is(,numItems)]*/ HRESULT   ** ppErrors
      );
 
      HRESULT STDMETHODCALLTYPE CreateEnumerator(
         /*[in]               */ REFIID      riid,
         /*[out, iid_is(riid)]*/ LPUNKNOWN * ppUnk
      );


// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
// [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[ IDataObject[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]

   HRESULT STDMETHODCALLTYPE GetData( 
      /* [unique][in] */ FORMATETC __RPC_FAR *pformatetcIn,
      /* [out] */ STGMEDIUM __RPC_FAR *pmedium
      );

        
   HRESULT STDMETHODCALLTYPE GetDataHere( 
      /* [unique][in] */ FORMATETC __RPC_FAR *pformatetc,
      /* [out][in] */ STGMEDIUM __RPC_FAR *pmedium
      );

        
   HRESULT STDMETHODCALLTYPE QueryGetData( 
      /* [unique][in] */ FORMATETC __RPC_FAR *pformatetc
      );

        
   HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc( 
      /* [unique][in] */ FORMATETC __RPC_FAR *pformatectIn,
      /* [out] */ FORMATETC __RPC_FAR *pformatetcOut
      );

        
   HRESULT STDMETHODCALLTYPE SetData( 
      /* [unique][in] */ FORMATETC __RPC_FAR *pformatetc,
      /* [unique][in] */ STGMEDIUM __RPC_FAR *pmedium,
      /* [in] */ BOOL fRelease
      );

        
   HRESULT STDMETHODCALLTYPE EnumFormatEtc( 
      /* [in] */ DWORD dwDirection,
      /* [out] */ IEnumFORMATETC __RPC_FAR *__RPC_FAR *ppenumFormatEtc
      );

        
   HRESULT STDMETHODCALLTYPE DAdvise( 
      /* [in] */ FORMATETC __RPC_FAR *pformatetc,
      /* [in] */ DWORD advf,
      /* [unique][in] */ IAdviseSink __RPC_FAR *pAdvSink,
      /* [out] */ DWORD __RPC_FAR *pdwConnection
      );

        
   HRESULT STDMETHODCALLTYPE DUnadvise( 
      /* [in] */ DWORD dwConnection
      );

        
   HRESULT STDMETHODCALLTYPE EnumDAdvise( 
      /* [out] */ IEnumSTATDATA __RPC_FAR *__RPC_FAR *ppenumAdvise
      );


   ///////////////////////////////////////////////////////////////////////////
   //////////////////////////// IOPCAsyncIO2 /////////////////////////////////
   ///////////////////////////////////////////////////////////////////////////

   STDMETHODIMP Read(
                  /* [in] */                    DWORD             count,
                  /* [size_is][in] */           OPCHANDLE      *  serverHandle,
                  /* [in] */                    DWORD             dwTransactionID,
                  /* [out] */                   DWORD          *  pdwCancelID,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

   STDMETHODIMP Write(
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][in] */           VARIANT        *  pItemValues,
                  /* [in] */                    DWORD             dwTransactionID,
                  /* [out] */                   DWORD          *  pdwCancelID,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

   STDMETHODIMP Refresh2(
                  /* [in] */                    OPCDATASOURCE     dwSource,
                  /* [in] */                    DWORD             dwTransactionID,
                  /* [out] */                   DWORD          *  pdwCancelID
                  );

   STDMETHODIMP Cancel2(
                  /* [in] */                    DWORD             dwCancelID
                  );

   STDMETHODIMP SetEnable(
                  /* [in] */                    BOOL              bEnable
                  );

   STDMETHODIMP GetEnable(
                  /* [out] */                   BOOL           *  pbEnable
                  );


   //=========================================================================
   // OPC 3.0 Custom Interfaces
   //=========================================================================
                                                                           
   ///////////////////////////////////////////////////////////////////////////
   //////////////////////////// IOPCGroupStateMgt2 ///////////////////////////
   ///////////////////////////////////////////////////////////////////////////
   
   STDMETHODIMP SetKeepAlive(
                  /* [in] */                    DWORD             dwKeepAliveTime,
                  /* [out] */                   DWORD          *  pdwRevisedKeepAliveTime
                  );
        
   STDMETHODIMP GetKeepAlive(
                  /* [out] */                   DWORD          *  pdwKeepAliveTime
                  );


   ///////////////////////////////////////////////////////////////////////////
   //////////////////////////// IOPCSyncIO2 //////////////////////////////////
   ///////////////////////////////////////////////////////////////////////////

   STDMETHODIMP ReadMaxAge(
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][in] */           DWORD          *  pdwMaxAge,
                  /* [size_is][size_is][out] */ VARIANT        ** ppvValues,
                  /* [size_is][size_is][out] */ WORD           ** ppwQualities,
                  /* [size_is][size_is][out] */ FILETIME       ** ppftTimeStamps,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

   STDMETHODIMP WriteVQT(
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][in] */           OPCITEMVQT     *  pItemVQT,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );


   ///////////////////////////////////////////////////////////////////////////
   //////////////////////////// IOPCAsyncIO3 /////////////////////////////////
   ///////////////////////////////////////////////////////////////////////////

   STDMETHODIMP ReadMaxAge(
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][in] */           DWORD          *  pdwMaxAge,
                  /* [in] */                    DWORD             dwTransactionID,
                  /* [out] */                   DWORD          *  pdwCancelID,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

   STDMETHODIMP WriteVQT(
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][in] */           OPCITEMVQT     *  pItemVQT,
                  /* [in] */                    DWORD             dwTransactionID,
                  /* [out] */                   DWORD          *  pdwCancelID,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

   STDMETHODIMP RefreshMaxAge(
                  /* [in] */                    DWORD             dwMaxAge,
                  /* [in] */                    DWORD             dwTransactionID,
                  /* [out] */                   DWORD          *  pdwCancelID
                  );


   ///////////////////////////////////////////////////////////////////////////
   //////////////////////////// IOPCItemDeadbandMgt //////////////////////////
   ///////////////////////////////////////////////////////////////////////////

   STDMETHODIMP SetItemDeadband(
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][in] */           FLOAT          *  pPercentDeadband,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

   STDMETHODIMP GetItemDeadband( 
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][size_is][out] */ FLOAT          ** ppPercentDeadband,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

   STDMETHODIMP ClearItemDeadband(
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

#ifdef NOT_YET_IMPLEMENTED
   ///////////////////////////////////////////////////////////////////////////
   //////////////////////////// IOPCItemSamplingMgt //////////////////////////
   ///////////////////////////////////////////////////////////////////////////

   STDMETHODIMP SetItemSamplingRate(
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][in] */           DWORD          *  pdwRequestedSamplingRate,
                  /* [size_is][size_is][out] */ DWORD          ** ppdwRevisedSamplingRate,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

   STDMETHODIMP GetItemSamplingRate(
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][size_is][out] */ DWORD          ** ppdwSamplingRate,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

   STDMETHODIMP ClearItemSamplingRate(
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

   STDMETHODIMP SetItemBufferEnable(
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][in] */           BOOL           *  pbEnable,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

   STDMETHODIMP GetItemBufferEnable(
                  /* [in] */                    DWORD             dwCount,
                  /* [size_is][in] */           OPCHANDLE      *  phServer,
                  /* [size_is][size_is][out] */ BOOL           ** ppbEnable,
                  /* [size_is][size_is][out] */ HRESULT        ** ppErrors
                  );

#endif //NOT_YET_IMPLEMENTED


   //======================================================================


// Overridables      
                                                // IConnectionPoint funtions
   STDMETHODIMP Advise( IUnknown* pUnkSink, DWORD* pdwCookie );
   STDMETHODIMP Unadvise( DWORD dwCookie );

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+  ...
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


   HRESULT FinalConstruct()
   {
      return CoCreateFreeThreadedMarshaler(
         GetControllingUnknown(), &m_pUnkMarshaler.p);
   }

   void FinalRelease()
   {
      m_pUnkMarshaler.Release();
   }

   CComPtr<IUnknown> m_pUnkMarshaler;

// IOPCGroup
protected:

      // each COM group (private and public) belongs to a COM server 
   DaGenericServer *m_pServer;

   DaBaseServer *m_pServerHandler;

               /** assigned by the generic part 
                  used to reach the group (DaGenericGroup)
                  assigned to this COM Group **/
   long  m_ServerGroupHandle;

   // IOPCDataCallback interface access
   // ---
      // Cookie used by Global Interface Table related functions to access
      // the IOPCDataCallback interface from other apartments.
      // This member is used only if the Global Interface Table is used for
      // marshaling of callback interfaces. Therefore the flag
      // _Module.m_fMarshalCallbacks is TRUE.
   private:
   DWORD m_dwCookieGITDataCb;

   public:
   HRESULT GetCallbackInterface( IOPCDataCallback** ppCallback );
   // ---

private:
      // Generic Group associated with this COM Group
   DaGenericGroup *m_pGroup;

public:
      // since a COM Group is generated by COM the default destructor is called
      // initialisation must be done with Create
   HRESULT Create(
            DaGenericServer *pServer,
            long           ServerGroupHandle,
            DaGenericGroup  *pGroup
         );

      // access to generic group associated to this COM group
   HRESULT GetGenericGroup( DaGenericGroup **group );
   HRESULT ReleaseGenericGroup();

private:
   HRESULT GetItemAttrList( DaGenericGroup *group, OPCITEMATTRIBUTES **ItemList, long *ItemCount );
   HRESULT FreeItemAttrList( OPCITEMATTRIBUTES *ItemList, long ItemCount );

   HRESULT GetCOMItemList( DaGenericGroup *group, LPUNKNOWN **ItemList, long *ItemCount );
   HRESULT FreeCOMItemList( LPUNKNOWN *ItemList, long ItemCount );

public:
      // Invoke IOPCDataCallback::OnDataChange
   HRESULT FireOnDataChange( DWORD dwNumOfItems, OPCITEMSTATE* pItemStates, HRESULT* errors );

   // Move this function to the private part if
   // AddCallrPhyvalItems() moved into the Generic Server part.

      // Called by AddItems() and AddCallrPhyValItems()
   HRESULT STDMETHODCALLTYPE AddItemsInternal( 
   /*[in]                       */ DWORD           numItems,
   /*[in, size_is(numItems)]  */ OPCITEMDEF     *pItemArray,
   /*[out, size_is(,numItems)]*/ OPCITEMRESULT **ppAddResults,
   /*[out, size_is(,numItems)]*/ HRESULT       **ppErrors,
                                   BOOL            fPhyValItem );

      // Implementation for functions ICallrItemMgt::ReserveItems and ICallrItemMgt::ReleaseItems
   HRESULT STDMETHODCALLTYPE ReserveOrReleaseItems( 
      /* [in] */                    BOOL fReserveItems,
      /* [in] */                    DWORD numItems,
      /* [size_is][in] */           OPCHANDLE __RPC_FAR *phServer,
      /* [size_is][size_is][out] */ HRESULT __RPC_FAR *__RPC_FAR *ppErrors );

      // Called by IOPCAsyncIO2::Read() and IOPCAsyncIO3::ReadMaxAge()
   HRESULT ReadAsync(
      /* [in] */                    DWORD             dwCount,
      /* [size_is][in] */           OPCHANDLE      *  phServer,
      /* [size_is][in] */           DWORD          *  pdwMaxAge,
      /* [in] */                    DWORD             dwTransactionID,
      /* [out] */                   DWORD          *  pdwCancelID,
      /* [size_is][size_is][out] */ HRESULT        ** ppErrors );

      // Called by IOPCAsyncIO2::Write() and IOPCAsyncIO3::WriteVQT()
   HRESULT WriteAsync(
      /* [in] */                    DWORD             dwCount,
      /* [size_is][in] */           OPCHANDLE      *  phServer,
      /* [size_is][in] */           OPCITEMVQT     *  pItemVQT,
      /* [in] */                    DWORD             dwTransactionID,
      /* [out] */                   DWORD          *  pdwCancelID,
      /* [size_is][size_is][out] */ HRESULT        ** ppErrors );

      // Called by IOPCAsyncIO2::Refresh2() and IOPCAsyncIO3::RefreshMaxAge()
   HRESULT Refresh2OrRefreshMaxAge(
      /* [in] */                    OPCDATASOURCE  *  pdwSource,
      /* [in] */                    DWORD          *  pdwMaxAge,
      /* [in] */                    DWORD             dwTransactionID,
      /* [out] */                   DWORD          *  pdwCancelID );

      // Called by IOPCSyncIO::Write() and IOPCSyncIO2::WriteVQT()
   HRESULT WriteSync(
      /* [in] */                    DWORD             dwCount,
      /* [size_is][in] */           OPCHANDLE      *  phServer,
      /* [size_is][in] */           OPCITEMVQT     *  pItemVQT,
      /* [size_is][size_is][out] */ HRESULT        ** ppErrors );

      // Called by IOPCItemDeadbandMgt::SetItemDeadband(), IOPCItemDeadbandMgt::GetItemDeadband()
      // and IOPCItemDeadbandMgt::ClearItemDeadband()
   HRESULT ItemDeadband(
      /* [in] */                    DWORD             dwCount,
      /* [size_is][in] */           OPCHANDLE      *  phServer,
      /* [size_is][size_is][out] */ HRESULT        ** ppErrors,
      /* [size_is][in] */           FLOAT          *  pPercentDeadbandIn = NULL,
      /* [size_is][in] */           FLOAT          *  pPercentDeadbandOut = NULL );
};
//DOM-IGNORE-END

#endif //__OPCGROUP_H
