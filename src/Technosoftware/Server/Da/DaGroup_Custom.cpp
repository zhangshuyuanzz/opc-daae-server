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

//DOM-IGNORE-BEGIN

#include "stdafx.h"
#include "DaGroup.h"
#include "DaGenericGroup.h"
#include "Logger.h"

////////////////////////////////////////////////////////////////////////////////////////////
//      Dispatch Interface
////////////////////////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCAsyncIODisp.AddCallbackReference
// ------------------------------------
// Store the Callback address for the specified request type.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::AddCallbackReference(
  /*[in]         */ long        Context,
  /*[in]         */ IDispatch * pCallback,
  /*[out, retval]*/ long      * pConnection )
{
         HRESULT                 res;
         DaGenericGroup           *group;
         UINT                    StreamFormat;
         LPWSTR                  CallbackMethodName;

   res = GetGenericGroup( &group );
   if( FAILED(res) ) {
      LOGFMTE("IOPCAsyncIODisp.AddCallbackReference: Internal error: No generic group." );
      goto AddCallbackExit0;
   }
   USES_CONVERSION;
   LOGFMTI( "IOPCAsyncIODisp.AddCallbackReference to group %s", W2A(group->m_Name));

   if( pConnection != NULL ) {
      *pConnection = 0;
   }

   if( ( pConnection == NULL ) || ( pCallback == NULL ) ) {

         LOGFMTE( "AddCallbackReference() failed with invalid argument(s):" );
         if( pConnection == NULL ) {
            LOGFMTE( "      pConnection is NULL" );
         }
         if( pCallback == NULL ) {
            LOGFMTE( "      pCallback is NULL" );
         }

      res = E_INVALIDARG;
      goto AddCallbackExit1a;
   }

   EnterCriticalSection( &group->m_CallbackCritSec );

   StreamFormat = (UINT)Context;

   if ((StreamFormat == group->m_StreamDataTime) || (StreamFormat == group->m_StreamData) ||
       (StreamFormat == 1) || (StreamFormat == 2)) {
         // both format do activate the DataTime Callback (OnDataChange)
         // so you cannot have both StreamData and StreamDataTime
      if (group->m_DataTimeCallbackDisp) {      // group already has a sink for this format 
         res = CONNECT_E_ADVISELIMIT;
         goto AddCallbackExit1;
      }

     // register the callback interface in the global interface table
     res = core_generic_main.m_pGIT->RegisterInterfaceInGlobal( pCallback, IID_IDispatch, (LPDWORD)&group->m_DataTimeCallbackDisp );
     if (FAILED( res )) {
        pCallback->Release();               // note :   register increments the refcount
        goto AddCallbackExit1;              //          even if the function failed     
     }

      CallbackMethodName = L"OnDataChange";
      res = pCallback->GetIDsOfNames(
                           IID_NULL,                              // dummy (not used)
                           &CallbackMethodName,                   // array[1] of names
                           1,                                     // size of array is 1
                           group->m_dwLCID,
                           &group->m_DataTimeCallbackMethodID );  // Method Id

      if (FAILED( res )) {
         group->m_DataTimeCallbackDisp = NULL;  // interface does not support OnDataChange!
         goto AddCallbackExit2;
      }
      *pConnection = STREAM_DATATIME_CONNECTION_DISP;
   }
   else if ((StreamFormat == group->m_StreamWrite) || (StreamFormat == 3)) {

      if (group->m_WriteCallbackDisp) {
         res = CONNECT_E_ADVISELIMIT;           // group already has a sink for this format 
         goto AddCallbackExit1;
      }

     // register the callback interface in the global interface table
     res = core_generic_main.m_pGIT->RegisterInterfaceInGlobal( pCallback, IID_IDispatch, (LPDWORD)&group->m_WriteCallbackDisp );
     if (FAILED( res )) {
        pCallback->Release();               // note :   register increments the refcount
        goto AddCallbackExit1;              //          even if the function failed     
     }

      CallbackMethodName = L"OnAsyncWriteComplete";
      res = pCallback->GetIDsOfNames( 
                           IID_NULL,                              // dummy (not used)
                           &CallbackMethodName,                   // array[1] of names
                           1,                                     // size of array is 1
                           group->m_dwLCID,
                           &group->m_WriteCallbackMethodID );     // Method Id

      if (FAILED( res )) {
         group->m_WriteCallbackDisp = NULL;     // interface does not support OnDataChange!
         goto AddCallbackExit2;
      }
      *pConnection = STREAM_WRITE_CONNECTION_DISP;
   }
   else {
      res = E_INVALIDARG;
      goto AddCallbackExit1;
   }

   LeaveCriticalSection( &group->m_CallbackCritSec );
   ReleaseGenericGroup( );
   return S_OK;

AddCallbackExit2:
   pCallback->Release();

AddCallbackExit1:
   LeaveCriticalSection( &group->m_CallbackCritSec );

AddCallbackExit1a:
   ReleaseGenericGroup();

AddCallbackExit0:
   return res;
}



//=========================================================================
// IOPCAsyncIODisp.DropCallbackReference
// -------------------------------------
// Release the Callback for the specified request type.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::DropCallbackReference(
  /*[in]*/ long Connection  )
{
      DaGenericGroup     *group;
      DaGenericItem*     pGItem = NULL;
      HRESULT           res;

   res = GetGenericGroup( &group );
   if (FAILED( res )) {
      LOGFMTE("IOPCAsyncIODisp.DropCallbackReference: Internal error: No generic group." );
      goto DropCallbackExit0;
   }
   USES_CONVERSION;
   LOGFMTI( "IOPCAsyncIODisp.DropCallbackReference for group %s", W2A(group->m_Name));

   EnterCriticalSection( &group->m_CallbackCritSec );

                           // Figure out what type of connection it is
   if (Connection == STREAM_DATATIME_CONNECTION_DISP) {

      if (group->m_DataTimeCallbackDisp == NULL) {
         res = OLE_E_NOCONNECTION;
         goto DropCallbackExit1;
      }

      res = core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( (DWORD)group->m_DataTimeCallbackDisp );
      group->m_DataTimeCallbackDisp = NULL;
      if (FAILED( res )) {
         goto DropCallbackExit1;
      }
   } 
   else if(Connection == STREAM_WRITE_CONNECTION_DISP) {

      if (group->m_WriteCallbackDisp == NULL) {
         res = OLE_E_NOCONNECTION;
         goto DropCallbackExit1;
      }

      res = core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( (DWORD)group->m_WriteCallbackDisp );
      group->m_WriteCallbackDisp = NULL;
      if (FAILED( res )) {
         goto DropCallbackExit1;
      }
   }
   else {                  // Bad stream format
      res = OLE_E_NOCONNECTION;
      goto DropCallbackExit1;
   }

   group->ResetLastReadOfAllGenericItems();     // Version 1.3
                                                // Clear the last read value of all generic items.
                                                // For this reason the values of all group items are sent again if
                                                // the next callback function will be installed.
   res = S_OK;

DropCallbackExit1:
   LeaveCriticalSection( &group->m_CallbackCritSec );
   ReleaseGenericGroup();

DropCallbackExit0:
   return res;
}
      


////////////////////////////////////////////////////////////////////////////////////////////
//      Custom Interface
////////////////////////////////////////////////////////////////////////////////////////////


//=========================================================================
// IDataObject.DAdvise
// -------------------
// Store the Callback address for the specified request type.
// 
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::DAdvise( 
   /* [in] */         FORMATETC __RPC_FAR   *pfe,
   /* [in] */         DWORD                  adv,
   /* [unique][in] */ IAdviseSink __RPC_FAR *pAdvSink,
   /* [out] */        DWORD __RPC_FAR       *pdwConnection   )
{
      DaGenericGroup     *group;
      IUnknown**        pp;
      HRESULT           res;

   res = GetGenericGroup( &group );
   if( FAILED(res) ) {
      LOGFMTE("IOPCAsyncIO::DAdvise: Internal error: No generic group." );
      goto DadviseExit0;
   }
   USES_CONVERSION;
   LOGFMTI( "IOPCAsyncIO::DAdvise for group %s", W2A(group->m_Name));

   if (pdwConnection) {  
      *pdwConnection = 0;
   }

   if( (pAdvSink == NULL) || (pfe == NULL) ) {  

         LOGFMTE( "DAdvise() failed with invalid argument(s) :" );
         if( pdwConnection == NULL ) {  
            LOGFMTE( "      pdwConnection is NULL" );
         }
         if( pAdvSink == NULL ) {
            LOGFMTE( "      pAdvSink is NULL" );
         }
         if( pfe == NULL ) {
            LOGFMTE( "      pfe is NULL" );
         }

      res = E_INVALIDARG;
      goto DadviseExit1a;
   }

   EnterCriticalSection( &group->m_CallbackCritSec );

   ///
   // Check if there is an existing IOPCDataCallback connection
   // 
   CComObject<DaGroup>* pCOMGroup;
   group->m_pServer->m_COMGroupList.GetElem( group->m_hServerGroupHandle, &pCOMGroup );
   if (pCOMGroup == NULL) {
      res = E_FAIL;
      goto DadviseExit1;
   }

   pCOMGroup->Lock();                        // Lock the connection point list
   pp = pCOMGroup->m_vec.begin();            // Check if there is a registered callback function
   pCOMGroup->Unlock();
   if (*pp) {
      LOGFMTE( "Error: There is an existing IOPCDataCallback connection" );
      res = CONNECT_E_ADVISELIMIT;
      goto DadviseExit1;
   }

   if (pfe->cfFormat ==  group->m_StreamData) {
      LOGFMTI( "   Format Type : OPCSTMFORMATDATA" );

         // check if there already is an automation Callback installed or 
         // an advise Sink for this stream format
      if (group->m_DataCallback) {
         res = CONNECT_E_ADVISELIMIT;           // group already has a sink for this format 
         goto DadviseExit1;
      }

     // register the callback interface in the global interface table
     res = core_generic_main.m_pGIT->RegisterInterfaceInGlobal( pAdvSink, IID_IAdviseSink, (LPDWORD)&group->m_DataCallback );
     if (FAILED( res )) {
        pAdvSink->Release();                // note :   register increments the refcount
        goto DadviseExit1;                  //          even if the function failed
     }
      *pdwConnection = STREAM_DATA_CONNECTION;
   }
   else if (pfe->cfFormat == group->m_StreamDataTime)  {
      LOGFMTI( "   Format Type : OPCSTMFORMATDATATIME" );

         // check if there already is an automation Callback installed or 
         // an advise Sink for this stream format
      if (group->m_DataTimeCallback) {
         res = CONNECT_E_ADVISELIMIT;           // group already has a sink for this format 
         goto DadviseExit1;
      }

     // register the callback interface in the global interface table
     res = core_generic_main.m_pGIT->RegisterInterfaceInGlobal( pAdvSink, IID_IAdviseSink, (LPDWORD)&group->m_DataTimeCallback );
     if (FAILED( res )) {
        pAdvSink->Release();                // note :   register increments the refcount
        goto DadviseExit1;                  //          even if the function failed
     }
      *pdwConnection = STREAM_DATATIME_CONNECTION;
   }
   else if (pfe->cfFormat == group->m_StreamWrite) {
      LOGFMTI( "   Format Type : OPCSTMFORMATWRITECOMPLETE" );

      if (group->m_WriteCallback) {
         res = CONNECT_E_ADVISELIMIT;           // group already has a sink for this format 
         goto DadviseExit1;
      }

     // register the callback interface in the global interface table
     res = core_generic_main.m_pGIT->RegisterInterfaceInGlobal( pAdvSink, IID_IAdviseSink, (LPDWORD)&group->m_WriteCallback );
     if (FAILED( res )) {
        pAdvSink->Release();                // note :   register increments the refcount
        goto DadviseExit1;                  //          even if the function failed
     }

      *pdwConnection = STREAM_WRITE_CONNECTION;
   }
   else {         // Bad stream format
      res = DV_E_FORMATETC;
      goto DadviseExit1;
   }

   LeaveCriticalSection( &group->m_CallbackCritSec );

   ReleaseGenericGroup() ;
   return S_OK;

DadviseExit1:
   LeaveCriticalSection( &group->m_CallbackCritSec );

DadviseExit1a:
   ReleaseGenericGroup();

DadviseExit0:
   return res;
}




        
//=========================================================================
// IDataObject.DUnadvise
// ---------------------
// Release the Callback for the specified request type.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::DUnadvise( 
   /* [in] */        DWORD dwConnection  )
{
      DaGenericGroup*    group;
      DaGenericItem*     pGItem = NULL;
      HRESULT           res;

   res = GetGenericGroup( &group );
   if( FAILED(res) ) {
      LOGFMTE("IOPCAsyncIO::DUnadvise: Internal error: No generic group." );
      goto DUnadviseExit0;
   }
   USES_CONVERSION;
   LOGFMTI( "IOPCAsyncIO::DUnadvise for group %s", W2A(group->m_Name));

   EnterCriticalSection( &group->m_CallbackCritSec );

                        // Figure out what type of connection it is
   if (dwConnection == STREAM_DATA_CONNECTION) {
      LOGFMTI( "   Format Type : OPCSTMFORMATDATA" );
      if( group->m_DataCallback == NULL ) {
         res = OLE_E_NOCONNECTION;
         goto DUnadviseExit1;
      }

      res = core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( (DWORD)group->m_DataCallback );
      group->m_DataCallback = NULL;
      if (FAILED( res )) {
         goto DUnadviseExit1;
      }
   } 
   else if (dwConnection == STREAM_DATATIME_CONNECTION) {
      LOGFMTI( "   Format Type : OPCSTMFORMATDATATIME" );
      if (group->m_DataTimeCallback == NULL) {
         res = OLE_E_NOCONNECTION;
         goto DUnadviseExit1;
      }

      res = core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( (DWORD)group->m_DataTimeCallback );
      group->m_DataTimeCallback = NULL;
      if (FAILED( res )) {
         goto DUnadviseExit1;
      }
   } 
   else if (dwConnection == STREAM_WRITE_CONNECTION) {
      LOGFMTI( "   Format Type : OPCSTMFORMATWRITECOMPLETE" );
      if( group->m_WriteCallback == NULL ) {
         res = OLE_E_NOCONNECTION;
         goto DUnadviseExit1;
      }

      res = core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( (DWORD)group->m_WriteCallback );
      group->m_WriteCallback = NULL;
      if (FAILED( res )) {
         goto DUnadviseExit1;
      }
   }
   else {               // Bad stream format
      res = OLE_E_NOCONNECTION;
      goto DUnadviseExit1;
   }

   group->ResetLastReadOfAllGenericItems();     // Version 1.3
                                                // Clear the last read value of all generic items.
                                                // For this reason the values of all group items are sent again if
                                                // the next callback function will be installed.
   res = S_OK;

DUnadviseExit1:
   LeaveCriticalSection( &group->m_CallbackCritSec );
   ReleaseGenericGroup();

DUnadviseExit0:
   return res;
}
        

        
//=========================================================================
// IDataObject.QueryGetData
// Since IDataObject deals with STREAMS only, this method need not be supported.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::QueryGetData( 
   /* [unique][in] */ FORMATETC __RPC_FAR *pformatetc  )
{
   return E_NOTIMPL;
}



        
//=========================================================================
// IDataObject.SetData
// Since IDataObject deals with STREAMS only, this method need not be supported.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::SetData( 
   /* [unique][in] */ FORMATETC __RPC_FAR *pformatetc,
   /* [unique][in] */ STGMEDIUM __RPC_FAR *pmedium,
   /* [in] */ BOOL fRelease  )
{
   return E_NOTIMPL;
}



        
//=========================================================================
// IDataObject.EnumFormatEtc
// Since IDataObject deals with STREAMS only, this method need not be supported.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::EnumFormatEtc( 
   /* [in] */ DWORD dwDirection,
   /* [out] */ IEnumFORMATETC __RPC_FAR *__RPC_FAR *ppenumFormatEtc  )
{
   return E_NOTIMPL;
}



        
//=========================================================================
// IDataObject.EnumDAdvise
// Since IDataObject deals with STREAMS only, this method need not be supported.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::EnumDAdvise( 
   /* [out] */ IEnumSTATDATA __RPC_FAR *__RPC_FAR *ppenumAdvise )
{
   return E_NOTIMPL;
}





//=========================================================================
// IDataObject.GetData
// Since IDataObject deals with STREAMS only, this method need not be supported.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::GetData( 
   /* [unique][in] */ FORMATETC __RPC_FAR *pformatetcIn,
   /* [out] */        STGMEDIUM __RPC_FAR *pmedium  )
{
   return E_NOTIMPL;
}



        
//=========================================================================
// IDataObject.GetDataHere
// Since IDataObject deals with STREAMS only, this method need not be supported.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::GetDataHere( 
   /* [unique][in] */ FORMATETC __RPC_FAR *pformatetc,
   /* [out][in] */    STGMEDIUM __RPC_FAR *pmedium  )
{
   return E_NOTIMPL;
}



        
//=========================================================================
// IDataObject.GetCanonicalFormatEtc
// Since IDataObject deals with STREAMS only, this method need not be supported.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::GetCanonicalFormatEtc( 
   /* [unique][in] */ FORMATETC __RPC_FAR *pformatectIn,
   /* [out] */        FORMATETC __RPC_FAR *pformatetcOut   )
{
   return E_NOTIMPL;
}

//DOM-IGNORE-END
