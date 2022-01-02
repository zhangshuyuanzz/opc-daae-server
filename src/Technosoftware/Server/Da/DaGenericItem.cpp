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

#include "DaGenericItem.h"
#include "DaGenericGroup.h"

#include "UtilityFuncs.h"
#include "enumclass.h"
#include "variantcompare.h"

      // ===============================================================
      //             Management Functions
      // ===============================================================

//=====================================================================================
// Constructor
//=====================================================================================
DaGenericItem::DaGenericItem( )
{
   m_Created = FALSE;
   m_RefCount           = 0 ;
   m_Active             = FALSE;
   m_ToKill             = FALSE;
   m_pGroup             = NULL;
   m_DeviceItem         = NULL;
   m_LastReadQuality    = OPC_QUALITY_BAD;

   memset( &m_ExtItemDef, 0, sizeof (ITEMDEFEXT) );

      // for access to members of this group (mostly m_RefCount and m_ToKill)
   InitializeCriticalSection( &m_CritSec );
   VariantInit( &m_LastReadValue ) ;
}



//=====================================================================================
// Initializer
//=====================================================================================
HRESULT DaGenericItem::Create( BOOL            Active,
                              OPCHANDLE       Client,
                              VARTYPE         ReqDataType,
                              DaGenericGroup   *pGroup,
                              DaDeviceItem     *pDeviceItem,
                              long            *pServerHandle,
                              BOOL            fPhyvalItem )
{
      HRESULT res;

   _ASSERTE( ( pGroup != NULL ) );
   _ASSERTE( ( pGroup != NULL ) );
   _ASSERTE( ( pDeviceItem != NULL ) );

   if ( m_Created == TRUE ) {
         // already created
      return E_FAIL;
   }

   m_ExtItemDef.m_fPhyvalItem = fPhyvalItem;

   m_pGroup = pGroup;
   m_RefCount           = 0 ;                
   m_Active             = Active;
   m_ToKill             = FALSE ;
   m_ClientHandle       = Client ;
   m_RequestedDataType  = ReqDataType ;
   m_DeviceItem         = pDeviceItem ;


   m_DeviceItem->Attach() ;                  // Inc DeviceItem RefCount

      // attach to group
   pGroup->Attach();

      // add item to group
   EnterCriticalSection( &m_pGroup->m_ItemsCritSec );
      // get new handle
   m_ServerHandle = m_pGroup->m_oaItems.New();
      
   res = m_pGroup->m_oaItems.PutElem( m_ServerHandle, this );
   if ( FAILED( res ) ) {
      goto CreateExit1;
   }
   LeaveCriticalSection( &m_pGroup->m_ItemsCritSec );
   *pServerHandle = m_ServerHandle;

   m_Created = TRUE;
                              // Notify the attached DeviceItem if there
                              // is a new active item.
   if (m_Active && pGroup->GetActiveState()) {
      AttachActiveCountOfDeviceItem();
   }
   return S_OK;

CreateExit1:
   LeaveCriticalSection( &m_pGroup->m_ItemsCritSec );
   m_pGroup->Detach();
   m_DeviceItem->Detach();
   return res;
}


//=====================================================================================
// Clone Initializer
//=====================================================================================
HRESULT DaGenericItem::Create( 
                              DaGenericGroup *pGroup,
                              DaGenericItem *pCloned,
                              long *pServerHandle )
{
      HRESULT res;

   _ASSERTE( pCloned != NULL );
   _ASSERTE( pServerHandle != NULL );
   _ASSERTE( pGroup != NULL );
   _ASSERTE( pCloned->m_DeviceItem != NULL );


   if ( m_Created == TRUE ) {
         // already created
      return E_FAIL;
   }

   m_RefCount           = 0 ;                
   m_ToKill             = FALSE ;
   m_Active             = pCloned->m_Active;
   m_ClientHandle       = pCloned->m_ClientHandle ;
   m_RequestedDataType  = pCloned->m_RequestedDataType ;

   res = pCloned->get_ExtItemDef( &m_ExtItemDef );
   if ( FAILED(res) ) {
      return res;
   }

   m_DeviceItem         = pCloned->m_DeviceItem ;
   m_DeviceItem->Attach() ;                  // Inc DeviceItem RefCount
                              // !!!??? device item cannot be deleted!

   m_pGroup             = pGroup;
   m_pGroup->Attach();                       // attach to group


                              // add item to group 
   EnterCriticalSection( &m_pGroup->m_ItemsCritSec );
   m_ServerHandle = m_pGroup->m_oaItems.New();
   res = m_pGroup->m_oaItems.PutElem( m_ServerHandle, this );
   if ( FAILED( res ) ) {
      goto CreateCloneExit1;
   }
   LeaveCriticalSection( &(m_pGroup->m_ItemsCritSec) );
   *pServerHandle = m_ServerHandle;

   m_Created = TRUE;

                              // Notify the attached DeviceItem if there
                              // is a new active item.
   if (m_Active && pGroup->GetActiveState()) {
      AttachActiveCountOfDeviceItem();
   }
      return S_OK;

CreateCloneExit1:
   LeaveCriticalSection( &m_pGroup->m_ItemsCritSec );
   m_pGroup->Detach();
   m_DeviceItem->Detach();
   return res;
}


   
//=====================================================================================
//  Destructor
//=====================================================================================
DaGenericItem::~DaGenericItem()
{

   if ( m_Created == TRUE ) {
         // here there are only shut-down 
         // proceedings reguarding a successfully created instance

         // Notify the attached DeviceItem that there is one less active DeviceItem.
      if (m_Active && m_pGroup->GetActiveState()) {
         DetachActiveCountOfDeviceItem();
      }
         
         // detach from group
      EnterCriticalSection( &m_pGroup->m_ItemsCritSec );
      m_pGroup->m_oaItems.PutElem( m_ServerHandle, NULL );
      LeaveCriticalSection( &m_pGroup->m_ItemsCritSec );

      m_pGroup->Detach();

         // Dec DeviceItem RefCount
      m_DeviceItem->Detach() ;                  
   }
   VariantClear( &m_LastReadValue ) ;

      // kill local data
   DeleteCriticalSection( &m_CritSec );


}



//=====================================================================================
//  Attach to class by incrementing the RefCount
//  Returns the current RefCount or -1 if the ToKill flag is set.
//=====================================================================================
int  DaGenericItem::Attach( void ) 
{
      long i;

   _ASSERTE( (m_Created == TRUE)  );

   EnterCriticalSection( &m_CritSec );

   if( m_ToKill ) {
      LeaveCriticalSection( &m_CritSec );
      return -1 ;
   }
   i = m_RefCount ++;
   LeaveCriticalSection( &m_CritSec );
   return i;
}




//=====================================================================================
//  Detach from class by decrementing the RefCount.
//  Kill the class instance if request pending.
//  Returns the current RefCount or -1 if it was killed.
//=====================================================================================
int  DaGenericItem::Detach( void )
{
      long i;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );
   _ASSERTE( m_RefCount > 0 );
   m_RefCount --;

   if( ( m_RefCount == 0 )                      // not referenced
   &&  ( m_ToKill == TRUE ) ){                  // kill request
      LeaveCriticalSection( &m_CritSec );
      delete this;                              // remove from memory
      return -1 ;
   }
   i = m_RefCount ;
   LeaveCriticalSection( &m_CritSec );
   return i;
}


//=====================================================================================
//  Kill the class instance.
//  If the RefCount is > 0 then only the kill request flag is set.
//  Returns the current RefCount or -1 if it was killed.
//=====================================================================================
int  DaGenericItem::Kill( void )
{
      long i;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   if( m_RefCount ) {            // still referenced
      m_ToKill = TRUE ;          // set kill request flag
   }
   else {
      LeaveCriticalSection( &m_CritSec );
      delete this ;              // Detach from DeviceItem is in Destructor
      return -1 ;
   }  
   i = m_RefCount ;
   LeaveCriticalSection( &m_CritSec );
   return i;
}


//=====================================================================================
// tells whether class instance was killed or not!
//=====================================================================================
BOOL DaGenericItem::Killed()
{
      BOOL b;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );
   b = m_ToKill;
   LeaveCriticalSection( &m_CritSec );
   return b;
}




//=====================================================================================
//  Attach to connected DeviceItem class by incrementing the RefCount
//  Returns the current RefCount and sets the class pointer variable
//  or -1 if the DeviceItem is not connected or the ToKill flag is set.
//=====================================================================================
int  DaGenericItem::AttachDeviceItem( DaDeviceItem **DItem )
{
      long rc;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   *DItem = m_DeviceItem ;                         // DeviceItem Ptr
   if( m_DeviceItem == NULL ) {   
      rc = -1 ;                                    // no connection
   } else {
      rc = m_DeviceItem->Attach() ;                // return RefCount
   }
   LeaveCriticalSection( &m_CritSec );
   return rc;
}




//=====================================================================================
//  Detach from class by decrementing the RefCount.
//  Kill the class instance if request pending.
//  Returns the current RefCount or -1 if it was killed.
//=====================================================================================
int  DaGenericItem::DetachDeviceItem( void )
{
      long rc;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   if( m_DeviceItem == NULL ) {
      rc = E_FAIL;
   } else {
      rc = m_DeviceItem->Detach() ;     // return RefCount
   }
   LeaveCriticalSection( &m_CritSec );

   return rc;
}






   // ===============================================================
   // function to get/change members and members 
   // ===============================================================

//=====================================================================================
// Get Active State
// ----------------
// if you want the return value to be sure
// enter the instance's CritSec before calling !
//=====================================================================================
BOOL DaGenericItem::get_Active( void )
{
      BOOL t;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   if( Killed() ) {   
      t = FALSE ;
   } else {
      t = m_Active ;
   }
   LeaveCriticalSection( &m_CritSec );
   return t;
}



//=====================================================================================
// Sets the Active State Flag and handles the active counter of the
// attached Device Item.
//=====================================================================================
HRESULT DaGenericItem::set_Active( BOOL Active )
{
   if (Killed()) {
      return E_FAIL;
   }

   HRESULT hres = S_OK;

   EnterCriticalSection( &m_CritSec );

   if (m_Active != Active) {                    // Notify the attached Device Item only
      m_Active = Active;                        // if the state has changed because there is 
                                                // a counter in the DeviceItem.
      if (m_pGroup->GetActiveState()) {         
                                                // Do not modify the counter if the group is inactive.
         if (Active) {
            hres = AttachActiveCountOfDeviceItem();
            ResetLastRead();                    // Force a subscription callback.
         }
         else {
            hres = DetachActiveCountOfDeviceItem();
         }
      }
   }

   LeaveCriticalSection( &m_CritSec );
   return hres;
}



//=====================================================================================
// Compares the last read value and quality with the specified values.
//
// Parameters:
//    IN
//       fltPercentDeadband
//       vCompValue
//       wCompQuality
//    OUT
//       fChanged
//
//=====================================================================================
HRESULT DaGenericItem::CompareLastRead( float fltPercentDeadband,
                                       const VARIANT& vCompValue, WORD wCompQuality,
                                       BOOL& fChanged )
{
   _ASSERTE( m_Created );
   _ASSERTE( m_DeviceItem );

   if (Killed()) {
      return E_FAIL;
   }

   // Handle Item specific Deadband
   FLOAT fltItemDeadband;
   if (SUCCEEDED( m_DeviceItem->GetItemDeadband( &fltItemDeadband ) )) {
      // Use the Item specific Deadband and ignore the Group specific Deadband
      fltPercentDeadband = fltItemDeadband;
   }

   HRESULT  hres = S_OK;

   fChanged = FALSE;

   EnterCriticalSection( &m_CritSec );

   if (m_LastReadQuality != wCompQuality) {
      fChanged = TRUE;
   }
   else {

      hres = CompareVariant(  *m_DeviceItem, fltPercentDeadband,
                              m_LastReadValue, vCompValue, fChanged );
   }
   LeaveCriticalSection( &m_CritSec );
   return hres;
}



//=====================================================================================
// Assigns new values to the members m_LastReadValue and m_LastReadQuality
//=====================================================================================
HRESULT DaGenericItem::UpdateLastRead( VARIANT vValue, WORD wQuality )
{
   _ASSERTE( m_Created );

   if (Killed()) {
      return E_FAIL;
   }

   EnterCriticalSection( &m_CritSec );
   HRESULT hres = VariantCopy( &m_LastReadValue, &vValue );
   if (SUCCEEDED( hres )) {
      m_LastReadQuality = wQuality;
   }
   LeaveCriticalSection( &m_CritSec );
   return hres;
}



//=====================================================================================
// Resets the members m_LastReadValue and m_LastReadQuality
//=====================================================================================
void DaGenericItem::ResetLastRead( void )
{
   _ASSERTE( m_Created );

   EnterCriticalSection( &m_CritSec );
   VariantClear( &m_LastReadValue );
   m_LastReadQuality = OPC_QUALITY_BAD;
   LeaveCriticalSection( &m_CritSec );
}



//=====================================================================================
// Get the data type requestd by the client
//=====================================================================================
VARTYPE DaGenericItem::get_RequestedDataType( void )
{
      VARTYPE rd;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );
   rd = m_RequestedDataType ;
   LeaveCriticalSection( &m_CritSec );
   return rd;
}



//=====================================================================================
// Sets Requested Data type in the generic item
// If the conversion from or to Canonical - Requested is not possible an 
// error is returned
//=====================================================================================
HRESULT DaGenericItem::set_RequestedDataType( VARTYPE RequestedDataType )
{
      OPCITEMDEF     ItemDef;
      OPCITEMRESULT  ItemResult;
      HRESULT        hrError = S_OK;
      HRESULT        hres;

   _ASSERTE( m_Created == TRUE );

   ItemResult.hServer               = 0;
   ItemResult.vtCanonicalDataType   = VT_EMPTY;
   ItemResult.wReserved             = 0;
   ItemResult.dwAccessRights        = 0;
   ItemResult.dwBlobSize            = 0;
   ItemResult.pBlob                 = NULL;

   EnterCriticalSection( &m_CritSec );

   hres = get_AccessPath( &ItemDef.szAccessPath );
   if (FAILED( hres )) {
      ItemDef.szAccessPath = WSTRClone( L"" );
   }
   hres = get_ItemIDCopy( &ItemDef.szItemID );
   if (FAILED( hres ) || ItemDef.szItemID == NULL) {
      ItemDef.szItemID = WSTRClone( L"" );
   }
   ItemDef.bActive = get_Active();
   ItemDef.hClient = get_ClientHandle();
   hres = get_Blob( &ItemDef.pBlob, &ItemDef.dwBlobSize );
   if (FAILED( hres )) {
      ItemDef.pBlob      = NULL;
      ItemDef.dwBlobSize = 0;
   }
   ItemDef.vtRequestedDataType = RequestedDataType;

      // call Validate item to see if requested data type would be accepted
   hres = m_pGroup->m_pServerHandler->OnValidateItems(
                                    OPC_VALIDATEREQ_ITEMRESULTS,
                                    FALSE,      // No Blob
                                    1,
                                    &ItemDef,
                                    NULL,
                                    &ItemResult,
                                    &hrError);
   FreeOPCITEMDEF( &ItemDef );

   if (SUCCEEDED( hres )) {
      hres = hrError;
   }
   if (SUCCEEDED( hres )) {
      m_RequestedDataType = RequestedDataType;  // Accepted data type
   }
   LeaveCriticalSection( &m_CritSec );
   return hres;
}




//=====================================================================================
// Get the client handle
//=====================================================================================
unsigned long  DaGenericItem::get_ClientHandle( void )
{
      DWORD ch;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );
   ch = m_ClientHandle ;
   LeaveCriticalSection( &m_CritSec );
   return ch;
}



//=====================================================================================
// Set the client handle
//=====================================================================================
void DaGenericItem::set_ClientHandle( unsigned long ClientHandle )
{
   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );
   m_ClientHandle = ClientHandle ;
   LeaveCriticalSection( &m_CritSec );
}




//=====================================================================================
// Get additional data of a generic item
//=====================================================================================
HRESULT DaGenericItem::get_ExtItemDef( ITEMDEFEXT * pExtItemDef )
{
   EnterCriticalSection( &m_CritSec );
   memcpy( pExtItemDef, &m_ExtItemDef, sizeof (ITEMDEFEXT) );
   LeaveCriticalSection( &m_CritSec );
   return S_OK;
}




//=====================================================================================
// Set additional data of a generic item
//=====================================================================================
HRESULT DaGenericItem::set_ExtItemDef( ITEMDEFEXT * pExtItemDef )
{
   EnterCriticalSection( &m_CritSec );
   memcpy( &m_ExtItemDef, pExtItemDef, sizeof (ITEMDEFEXT) );
   LeaveCriticalSection( &m_CritSec );
   return S_OK;
}




// ===============================================================
// function to get/change members and members of the
// connected DaDeviceItem class
// ===============================================================

//=====================================================================================
// Return an item ID (Pointer)
//    Returns the pointer !!! so do not free !
//
//    Do not use this function if you implement a CALL-R Server.
//=====================================================================================
HRESULT DaGenericItem::get_ItemIDPtr( LPWSTR *ItemID )
{
      HRESULT res;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   if( m_DeviceItem == NULL ) {
      LeaveCriticalSection( &m_CritSec );
      return E_FAIL;
   }
   res = m_DeviceItem->get_ItemIDPtr( ItemID );
   LeaveCriticalSection( &m_CritSec );
   return res;
}



//=====================================================================================
// Return an item ID (Copy)
// ------------------------
//    Returns a copy of the Item ID.
//    The access to the Item ID is protected with a critical section.
//=====================================================================================
HRESULT DaGenericItem::get_ItemIDCopy( LPWSTR *ItemID )
{
      HRESULT res;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   if( m_DeviceItem == NULL ) {
      LeaveCriticalSection( &m_CritSec );
      return E_FAIL;
   }
   res = m_DeviceItem->get_ItemIDCopy( ItemID );
   LeaveCriticalSection( &m_CritSec );
   return res;
}



//=====================================================================================
// !!!??? this function has not much sense as long items
// cannot be renamed!!!!???
//=====================================================================================
HRESULT DaGenericItem::set_ItemID( LPWSTR ItemID )
{
      HRESULT res;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   if( m_DeviceItem == NULL ) {
      LeaveCriticalSection( &m_CritSec );
      return E_FAIL;
   }
   res = m_DeviceItem->set_ItemID( ItemID );

   LeaveCriticalSection( &m_CritSec );

   return res;
}




//=====================================================================================
// get the actual access path
// --------------------------
//=====================================================================================
HRESULT DaGenericItem::get_AccessPath( LPWSTR *AccessPath )
{
      HRESULT res;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   if( m_DeviceItem == NULL ) {
      LeaveCriticalSection( &m_CritSec );
      return E_FAIL;
   }
   res = m_DeviceItem->get_AccessPath( AccessPath  );
   LeaveCriticalSection( &m_CritSec );
   return res;
}



//=====================================================================================
// set the actual access path
// --------------------------
//=====================================================================================
HRESULT DaGenericItem::set_AccessPath( LPWSTR AccessPath )
{
      HRESULT res;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   if( m_DeviceItem == NULL ) {
      LeaveCriticalSection( &m_CritSec );
      return E_FAIL;
   }
   res = m_DeviceItem->set_AccessPath( AccessPath  );
   LeaveCriticalSection( &m_CritSec );
   return res;
}




//=====================================================================================
//=====================================================================================
HRESULT DaGenericItem::get_Blob( BYTE **pBlob, DWORD *BlobSize )
{
      HRESULT res;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   if( m_DeviceItem == NULL ) {
      LeaveCriticalSection( &m_CritSec );
      return E_FAIL;
   }
   res = m_DeviceItem->get_Blob( pBlob, BlobSize  );
   LeaveCriticalSection( &m_CritSec );
      return res;
}



//=====================================================================================
//=====================================================================================
HRESULT DaGenericItem::set_Blob( BYTE *pBlob, DWORD BlobSize )
{
      HRESULT res;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   if( m_DeviceItem == NULL ) {
      LeaveCriticalSection( &m_CritSec );
      return E_FAIL;
   }
   res = m_DeviceItem->set_Blob( pBlob, BlobSize  );
   LeaveCriticalSection( &m_CritSec );
   return res;
}




//=====================================================================================
//=====================================================================================
VARTYPE DaGenericItem::get_CanonicalDataType( void )
{
      VARTYPE t;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   if( m_DeviceItem == NULL ) {
      LeaveCriticalSection( &m_CritSec );
      return VT_EMPTY;
   }
   t = m_DeviceItem->get_CanonicalDataType();
   LeaveCriticalSection( &m_CritSec );
   return t;
}



//=====================================================================================
//=====================================================================================
HRESULT DaGenericItem::get_AccessRights( DWORD *DaAccessRights )
{
      HRESULT res;

   _ASSERTE( ( m_Created == TRUE ) );

   EnterCriticalSection( &m_CritSec );

   if( m_DeviceItem == NULL ) {
      LeaveCriticalSection( &m_CritSec );
      return E_FAIL;
   }
   res = m_DeviceItem->get_AccessRights( DaAccessRights );
   LeaveCriticalSection( &m_CritSec );
   return res;
}




//=====================================================================================
//=====================================================================================
HRESULT DaGenericItem::get_EUData( OPCEUTYPE *EUType, VARIANT *EUInfo )
{
      HRESULT res;

   _ASSERTE( ( m_Created == TRUE ) );
   EnterCriticalSection( &m_CritSec );

   if( m_DeviceItem == NULL ) {
      LeaveCriticalSection( &m_CritSec );
      return E_FAIL;
   }
   res = m_DeviceItem->get_EUData( EUType, EUInfo );
   LeaveCriticalSection( &m_CritSec );
   return res;
}



//=====================================================================================
// get all item attributes.
//
//    Since Call-R functions can change item attributes entire attribute
//    record must be protected with critical section.
//=====================================================================================
HRESULT DaGenericItem::get_Attr( OPCITEMATTRIBUTES * pAttr, OPCHANDLE hServer )
{
                                       // Protect all item attributes (for CALL-R)
   EnterCriticalSection( &m_DeviceItem->m_CritSecAllAttrs );

   get_AccessPath( &pAttr->szAccessPath );
   get_ItemIDCopy( &pAttr->szItemID );
   pAttr->bActive = get_Active();
   pAttr->hClient = get_ClientHandle();
   pAttr->hServer = hServer;
   get_AccessRights( &pAttr->dwAccessRights );
   get_Blob( &pAttr->pBlob, &pAttr->dwBlobSize );
   pAttr->vtRequestedDataType = get_RequestedDataType();
   pAttr->vtCanonicalDataType = get_CanonicalDataType();
   get_EUData( &pAttr->dwEUType, &pAttr->vEUInfo );

                                       // Release item attributes protection (for CALL-R)
   LeaveCriticalSection( &m_DeviceItem->m_CritSecAllAttrs );
   return S_OK;
}




//=====================================================================================
// Attach the Active Counter of the attached DeviceItem.
//=====================================================================================
HRESULT DaGenericItem::AttachActiveCountOfDeviceItem( void )
{
   _ASSERTE( m_Created == TRUE );
   _ASSERTE( m_DeviceItem != NULL );

   if (Killed()) {
      return E_FAIL ;
   }
   return m_DeviceItem->AttachActiveCount();
}




//=====================================================================================
// Detach the Active Counter of the attached DeviceItem.
//=====================================================================================
HRESULT DaGenericItem::DetachActiveCountOfDeviceItem( void )
{
   _ASSERTE( m_Created == TRUE );
   _ASSERTE( m_DeviceItem != NULL );

   if (Killed()) {
      return E_FAIL ;
   }
   return m_DeviceItem->DetachActiveCount();
}
//DOM-IGNORE-END
