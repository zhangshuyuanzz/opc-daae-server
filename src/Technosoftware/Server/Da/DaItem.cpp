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
#include "DaItem.h"
#include "DaGenericGroup.h"
#include "UtilityFuncs.h"
#include "Logger.h"

/////////////////////////////////////////////////////////////////////////////
// DaItem

// !!!???
// in theory all the methods that access
// a GenericItem should be synchronized !!!
// it's possible that a thread reading and writing
// a property will collide and destroy data!

//================================
// constructor
//================================
DaItem::DaItem() 
{
   m_pUnkMarshaler = NULL;
   m_LastReadError = 0;
   m_LastWriteError = 0;
   m_pItem = NULL;
   m_pGroup = NULL;
   m_ServerHandle = 0;
   //!!!??? m_Accessing threads counter
}



//===========================================================================================
// Initializes a group
//===========================================================================================
HRESULT DaItem::Create( DaGenericGroup *pGroup, long ServerHandle, DaGenericItem *pItem )
{
   _ASSERTE( ( (pGroup != NULL) && (ServerHandle != 0) && (pItem != NULL) ) );

      // attach to the item
   EnterCriticalSection( &pGroup->m_ItemsCritSec );
   pItem->Attach();
   LeaveCriticalSection( &pGroup->m_ItemsCritSec );

      // as long as item exists also the group exists!
   m_pGroup = pGroup;
   m_ServerHandle = ServerHandle;
   m_pItem = pItem;
   return S_OK;
}


//================================
// destructor
//================================
DaItem::~DaItem() 
{
      // delete the COM item from the Group COM list!
      // eventually the GenericItem must be released!
   EnterCriticalSection( &(m_pGroup->m_ItemsCritSec) );
   m_pGroup->m_oaCOMItems.PutElem( m_ServerHandle, NULL );

      // detach from the generic item
   m_pItem->Detach();

   LeaveCriticalSection( &(m_pGroup->m_ItemsCritSec) );
}


//+++++++++++++++++++++++++++++++++++++
//+ IOPCItemDisp
//+++++++++++++++++++++++++++++++++++++

//===========================================================================================
// IOPCItemDisp.get_AccessPath
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_AccessPath(
  /*[out, retval]*/ BSTR * pAccessPath   )
{
      LPWSTR         AccessPath;
      HRESULT        res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.get_AccessPath");

   if ( pAccessPath == NULL ) {
      LOGFMTE( "get_AccessPath() failed with invalid argument(s):  pAccessPath is NULL" );
      return E_INVALIDARG;
   }
   *pAccessPath = NULL;

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   res = item->get_AccessPath( &AccessPath );
   if ( SUCCEEDED( res ) ) {
      if ( AccessPath == NULL ) {
         *pAccessPath = NULL;
      } else {
            // create a COM string
         *pAccessPath = SysAllocString( AccessPath );
         if (*pAccessPath == NULL)  {
            res = E_OUTOFMEMORY;
         }
      }
      WSTRFree( AccessPath, NULL);
   }

   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_AccessRights
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_AccessRights(
  /*[out, retval]*/ long * pAccessRights
)
{
      DWORD          DaAccessRights;
      HRESULT        res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.get_AccessRights");

   if ( pAccessRights == NULL ) {

      LOGFMTE( "get_AccessRights() failed with invalid argument(s): pAccessRights is NULL" );
      return E_INVALIDARG;
   }
   *pAccessRights = 0;

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   res = item->get_AccessRights( &DaAccessRights );
   if ( SUCCEEDED( res ) ) {
      *pAccessRights = (long)DaAccessRights;
   }

   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_ActiveStatus
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_ActiveStatus(
//        /*[out, retval]*/ boolean * pActiveStatus
  /*[out, retval]*/ VARIANT_BOOL * pActiveStatus
)
{
      BOOL Active;
      HRESULT res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.get_ActiveStatus");

   if ( pActiveStatus == NULL ) {

      LOGFMTE( "get_ActiveStatus() failed with invalid argument(s): pActiveStatus is NULL" );
      return E_INVALIDARG;
   }
   *pActiveStatus = FALSE;

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   Active = item->get_Active();
   if ( SUCCEEDED( res ) ) {
      if ( Active == FALSE ) {
         *pActiveStatus = VARIANT_FALSE;
      } else {
         *pActiveStatus = VARIANT_TRUE;
      }
   }

   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.put_ActiveStatus
//    [ propput ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::put_ActiveStatus(
//        /*[in]*/ boolean ActiveStatus
  /*[in]*/ VARIANT_BOOL ActiveStatus
)
{
      HRESULT res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.put_ActiveStatus");

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   if ( ActiveStatus == VARIANT_FALSE ) {
      res = item->set_Active( FALSE );
   } else {
      res = item->set_Active( TRUE );
   }

   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_Blob
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_Blob(
  /*[out, retval]*/ VARIANT * pBlob
)
{
      HRESULT        res;
      DaGenericItem   *item;
      DWORD          BlobSize;
      BYTE           *pBytesBlob;

   LOGFMTI("IOPCItemDisp.get_Blob");

   if ( pBlob == NULL ) {

      LOGFMTE( "get_Blob() failed with invalid argument(s): pBlob is NULL" );
      return E_INVALIDARG;
   }

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   item->get_Blob( &pBytesBlob, &BlobSize );

   ConvertBlobBytesToVariant(
            BlobSize,
            pBytesBlob, 
            &pBlob, 
            pIMalloc
         );


   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_ClientHandle
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_ClientHandle(
  /*[out, retval]*/ long * phClient   )
{
      HRESULT res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.get_ClientHandle");

   if ( phClient == NULL ) {

      LOGFMTE( "get_ClientHandle() failed with invalid argument(s): phClient is NULL" );
      return E_INVALIDARG;
   }
   *phClient = 0;

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }
   *phClient = item->get_ClientHandle();
   ReleaseGenericItem();
   return res;
}



//===========================================================================================
// IOPCItemDisp.put_ClientHandle
//    [ propput ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::put_ClientHandle(
  /*[in]*/ long Client )
{
      HRESULT        res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.put_ClientHandle");

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }
   item->set_ClientHandle( Client );
   ReleaseGenericItem();
   return res;
}



//===========================================================================================
// IOPCItemDisp.get_ItemID
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_ItemID(
  /*[out, retval]*/ BSTR * pItemID
)
{
      HRESULT        res;
      LPWSTR         ItemID;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.get_ItemID");

   if ( pItemID == NULL ) {

      LOGFMTE( "get_ItemID() failed with invalid argument(s): pItemID is NULL" );
      return E_INVALIDARG;
   }
   *pItemID = NULL;

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   res = item->get_ItemIDCopy( &ItemID );
   if ( SUCCEEDED( res ) ) {
      if ( ItemID == NULL ) {
         *pItemID = NULL;
      } else {
            // create a COM string
         *pItemID = SysAllocString( ItemID );
         if (*pItemID == NULL)  {
            res = E_OUTOFMEMORY;
         }
         WSTRFree( ItemID, NULL );
      }
   }

   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_ServerHandle
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_ServerHandle(
  /*[out, retval]*/ long * pServer
)
{
      HRESULT res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.get_ServerHandle");

	if ( pServer == NULL ) {

      LOGFMTE( "get_ServerHandle() failed with invalid argument(s): pServer is NULL" );
      return E_INVALIDARG;
   }
   *pServer = 0;

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   *pServer = m_ServerHandle;

   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_RequestedDataType
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_RequestedDataType(
  /*[out, retval]*/ short * pRequestedDataType   )
{
      HRESULT res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.get_RequestedDataType");

   if ( pRequestedDataType == NULL ) {

      LOGFMTE( "get_RequestedDataType() failed with invalid argument(s): pRequestedDataType is NULL" );
      return E_INVALIDARG;
   }
   *pRequestedDataType = 0;

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }
   *pRequestedDataType = item->get_RequestedDataType();
   ReleaseGenericItem();
   return res;
}



//===========================================================================================
// IOPCItemDisp.put_RequestedDataType
//    [ propput ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::put_RequestedDataType(
  /*[in]*/ short RequestedDataType
)
{
      HRESULT res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.put_RequestedDataType");

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   res =  item->set_RequestedDataType( RequestedDataType );

   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_Value
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_Value(
  /*[out, retval]*/ VARIANT * pValue
)
{
      VARIANT tQuality, tTimestamp;
      HRESULT res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.get_Value");

   if ( pValue == NULL ) {

      LOGFMTE( "get_Value() failed with invalid argument(s): pValue is NULL" );
      return E_INVALIDARG;
   }

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   res = OPCRead( OPC_DS_CACHE, pValue, &tQuality, &tTimestamp);
   if ( FAILED( res ) ) {
      m_LastWriteError = res;
      goto ItemGetValueExit0;
   }
   //save quality and timestamp?

ItemGetValueExit0:
   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.put_Value
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::put_Value(
  /*[in]*/ VARIANT NewValue
)
{
      HRESULT        res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.put_Value");

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   res = OPCWrite( NewValue );

   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_Quality
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_Quality(
  /*[out, retval]*/ short * pQuality
)
{
      HRESULT        res;
      DaGenericItem   *item;
      VARIANT        tValue, tQuality, tTimestamp;

   LOGFMTI("IOPCItemDisp.get_Quality");

   if ( pQuality == NULL ) {

      LOGFMTE( "get_Quality() failed with invalid argument(s): pQuality is NULL" );
      return E_INVALIDARG;
   }

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   res = OPCRead( OPC_DS_CACHE, &tValue, &tQuality, &tTimestamp);
   if ( FAILED( res ) ) {
      m_LastReadError = res;
      goto GetQualityExit0;
   }

   *pQuality = V_I2( &tQuality );

GetQualityExit0:
   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_Timestamp
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_Timestamp(
  /*[out, retval]*/ DATE * pTimeStamp
)
{
      HRESULT        res;
      DaGenericItem   *item;
      VARIANT        tValue, tQuality, tTimestamp;


   LOGFMTI("IOPCItemDisp.get_Timestamp");

   if ( pTimeStamp == NULL ) {

      LOGFMTE( "get_Timestamp() failed with invalid argument(s): pTimeStamp is NULL" );
      return E_INVALIDARG;
   }

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   res = OPCRead( OPC_DS_CACHE, &tValue, &tQuality, &tTimestamp);
   if ( FAILED( res ) ) {
      m_LastReadError = res;
      goto GetTimestampExit0;
   }

   *pTimeStamp = V_DATE( &tTimestamp );

GetTimestampExit0:
   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_ReadError
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_ReadError(
  /*[out, retval]*/ long * pError
)
{
      HRESULT        res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.get_ReadError");

   if ( pError == NULL ) {

      LOGFMTE( "get_ReadError() failed with invalid argument(s): pError is NULL" );
      return E_INVALIDARG;
   }

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   *pError = m_LastReadError;

   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_WriteError
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_WriteError(
  /*[out, retval]*/ long * pError
)
{
      HRESULT        res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.get_WriteError");

   if ( pError == NULL ) {

      LOGFMTE( "get_WriteError() failed with invalid argument(s): pError is NULL" );
      return E_INVALIDARG;
   }

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   *pError = m_LastWriteError;

   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_EUType
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_EUType(
  /*[out, retval]*/ short * pError
)
{
      OPCEUTYPE      EUType;
      VARIANT        EUInfo;
      HRESULT        res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.get_EUType");

   if ( pError == NULL ) {

      LOGFMTE( "get_EUType() failed with invalid argument(s): pError is NULL" );
      return E_INVALIDARG;
   }

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }


   res = item->get_EUData( &EUType, &EUInfo );
   *pError = (short)EUType;

   ReleaseGenericItem();

   return res;
}

//===========================================================================================
// IOPCItemDisp.get_EUInfo
//    [ propget ]
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::get_EUInfo(
  /*[out, retval]*/ VARIANT * pError
)
{
      OPCEUTYPE      EUType;
      HRESULT        res;
      DaGenericItem   *item;

   LOGFMTI("IOPCItemDisp.get_EUInfo");

   if ( pError == NULL ) {

      LOGFMTE( "get_EUInfo() failed with invalid argument(s): pError is NULL" );
      return E_INVALIDARG;
   }

   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   res = item->get_EUData( &EUType, pError );

   ReleaseGenericItem();

   return res;
}

      //
      // Methods
      //

//===========================================================================================
// IOPCItemDisp.OPCRead
// read from cache
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::OPCRead(
  /*[in]           */ short     Source,
  /*[out]          */ VARIANT * pValue,
  /*[out, optional]*/ VARIANT * pQuality,
  /*[out, optional]*/ VARIANT * pTimestamp
)
{
      OPCITEMSTATE   tItemValue;
      HRESULT        tError = S_OK;
      HRESULT        res;
      DaGenericItem   *GItem;
      DaDeviceItem    *DItem;
      DWORD          AccessRight;
      #ifdef _CALLR_ICALLRITEMMGT
         ITEMDEFEXT        ExtItemDef;
      #endif

   LOGFMTI("IOPCItemDisp.OPCRead");

   if ((Source != OPC_DS_CACHE) && (Source != OPC_DS_DEVICE) ) {
      LOGFMTE( "OPCRead() failed with error: 'Source' argument is invalid" );
      return E_INVALIDARG;
   }

   if ( pValue != NULL ) {
      VariantInit( pValue );
   }

   if ( (pValue == NULL) || (pQuality == NULL) || (pTimestamp == NULL) ) {

         LOGFMTE( "OPCRead() failed with invalid argument(s):" );
         if ( pValue == NULL ) {
            LOGFMTE( "      pValue is NULL" );
         }
         if ( pQuality == NULL ) {
            LOGFMTE( "      pQuality is NULL" );
         }
         if ( pTimestamp == NULL ) {
            LOGFMTE( "      pTimestamp is NULL" );
         }

      return E_INVALIDARG;
   }

   res = GetGenericItem( &GItem );
   if ( FAILED( res ) ) {
      return res;
   }
   
   res = GItem->get_AccessRights( &AccessRight );
   if ( FAILED(res) ) {
      goto OPCReadExit1;
   }
   if ( (AccessRight & OPC_READABLE) == 0 ) {
      res = OPC_E_BADRIGHTS;
      goto OPCReadExit1;
   }

   res = GItem->AttachDeviceItem( &DItem );
   if ( FAILED(res) ) {
      goto OPCReadExit1;
   }

   V_VT( &tItemValue.vDataValue ) = GItem->get_RequestedDataType();

   res = m_pGroup->InternalRead( 
               Source,
               1, 
               &DItem, 
               &tItemValue,
               &tError
               );

   // Note : res is always S_OK or S_FALSE
   if( FAILED( tError ) ) {
      res = tError;
      goto OPCReadExit2;
   }

   VariantClear( pValue );
   res = VariantCopy( pValue , &tItemValue.vDataValue ); 
   VariantClear( &tItemValue.vDataValue );      // No longer used
   if (FAILED( res )) {
      goto OPCReadExit2;
   }

   if ( (pQuality != NULL) && (V_VT( pQuality ) != VT_NULL) ) {
      VariantClear( pQuality );
      V_VT( pQuality ) = VT_I2;
      V_I2( pQuality ) = tItemValue.wQuality;
   }

   if ( (pTimestamp != NULL) && (V_VT( pTimestamp ) != VT_NULL) ) {
      VariantClear( pTimestamp );
      V_VT( pTimestamp ) = VT_DATE;
      FileTimeToDATE( &tItemValue.ftTimeStamp, V_DATE( pTimestamp ) );
   }
   res = S_OK;


OPCReadExit2:
   DItem->Detach();
   
OPCReadExit1:
   ReleaseGenericItem();
   m_LastReadError = res;

   return res;
}


//===========================================================================================
// IOPCItemDisp.OPCWrite
// write to device
//===========================================================================================
HRESULT STDMETHODCALLTYPE DaItem::OPCWrite(
  /*[in]*/ VARIANT NewValue
)
{
      HRESULT        tError = S_OK;
      HRESULT        res;
      DaGenericItem  *item;
      DaDeviceItem   *DItem;
      DWORD          AccessRight;


   res = GetGenericItem( &item );
   if ( FAILED( res ) ) {
      return res;
   }

   res = item->get_AccessRights( &AccessRight );
   if ( FAILED(res) ) {
      goto OPCWriteExit1;
   }
   if ( (AccessRight & OPC_WRITEABLE) == 0 ) {
      res = OPC_E_BADRIGHTS;
      goto OPCWriteExit1;
   }

   res = item->AttachDeviceItem( &DItem );
   if ( FAILED(res) ) {
      goto OPCWriteExit1;
   }

   OPCITEMVQT  ItemVQT;
   memset( &ItemVQT, 0, sizeof (OPCITEMVQT));
   res = VariantCopy( &ItemVQT.vDataValue, &NewValue );
   if (FAILED( res )) goto OPCWriteExit3;

   res = m_pGroup->InternalWriteVQT(
               1, 
               &DItem, 
               &ItemVQT,
               &tError
               );

   VariantClear( &ItemVQT.vDataValue );

   // Note : res is always S_OK or S_FALSE
   res = tError;

OPCWriteExit3:

   if (DItem) {
      DItem->Detach();
   }

OPCWriteExit1:
   ReleaseGenericItem();

   m_LastWriteError = res;

   return res;
}


//////////////////////////////////////////////////////////////


//===========================================================================================
//===========================================================================================
HRESULT DaItem::GetGenericItem( DaGenericItem **ppItem )
{
   return m_pGroup->GetGenericItem( m_ServerHandle, ppItem );
}


//===========================================================================================
//===========================================================================================
HRESULT DaItem::ReleaseGenericItem( )
{
   return m_pGroup->ReleaseGenericItem( m_ServerHandle );
}


//DOM-IGNORE-END



