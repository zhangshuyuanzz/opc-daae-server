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
#include <comdef.h>                             // for _variant_t
#include "DaDeviceItem.h"
#include "UtilityFuncs.h"
#include "variantconversion.h"
#include "DaBaseServer.h"

//=========================================================================
// Constructor
//=========================================================================
DaDeviceItem::DaDeviceItem( void )
{
   m_ItemID             = NULL;
   m_AccessPath         = NULL;
   m_Active             = FALSE;
   m_ToKill             = FALSE;
   m_RefCount           = 0;
   m_AccessRights       = OPC_READABLE; 
   m_Quality            = OPC_QUALITY_BAD;   // init to BAD
   m_BlobSize           = 0;
   m_pBlob              = NULL;
   m_EUType             = OPC_NOENUM;
   m_dwActiveCount      = 0;
   m_dAnalogEURange     = 0;
   m_fltPercentDeadband = -1;

   VariantInit( &m_EUInfo );
   VariantInit( &m_Value  );

   CoFileTimeNow( &m_TimeStamp );      // now

   InitializeCriticalSection( &m_CritSecAllAttrs );
   InitializeCriticalSection( &m_CritSec );
}



//=========================================================================
// Initializer
//=========================================================================
HRESULT DaDeviceItem::Create(  LPWSTR                  szItemID,
                              DWORD                   dwAccessRights,
                              LPVARIANT               pvValue,
                              BOOL                    fActive,       /* = TRUE */
                              DWORD                   dwBlobSize,    /* = 0 */
                              LPBYTE                  pBlob,         /* = NULL */
                              LPWSTR                  szAccessPath,  /* = NULL */
                              OPCEUTYPE               eEUType,       /* = OPC_NOENUM */
                              LPVARIANT               pvEUInfo       /* = NULL */
                           )
{
   _ASSERTE( V_ISBYREF( pvValue ) == FALSE );   // BYREF not supported as canonical data type
   #ifdef _DEBUG
      if (eEUType != OPC_NOENUM) {
         _ASSERTE( pvEUInfo );                  // Must not be NULL if EUType is not OPC_NOENUM
      }
      if (dwBlobSize) {
         _ASSERTE( pBlob );                     // Must not be NULL if Blob Size is not 0
      }
   #endif
   #ifdef _NO_DEADBAND_HANDLING                 // If there are Items with Analog EU Info then
      _ASSERTE( eEUType != OPC_ANALOG );        // the compiler option _NO_DEADBAND_HANDLING
   #endif                                       // must not be set.
                                                

   HRESULT hres;

   hres = set_ItemID( szItemID );
   if (FAILED( hres )) {
      return hres;
   }
   hres = set_AccessRights( dwAccessRights );
   hres = set_Active( fActive );
   hres = set_Blob( pBlob, dwBlobSize );
   hres = set_AccessPath( szAccessPath );

   if (pvEUInfo) {
      hres = set_EUData( eEUType, pvEUInfo );
   }
   else {
      _variant_t vEUDummy;
      hres = set_EUData( eEUType, &vEUDummy );
   }
   hres = VariantCopy( &m_Value, pvValue );

   if (FAILED( hres )) {
      return hres;
   }
   Attach();                                     // Now the Device Item is created
   return S_OK;
}


//=========================================================================
// Destructor
//=========================================================================
DaDeviceItem::~DaDeviceItem()
{
   if (m_ItemID) {
      delete m_ItemID;
   }
   if (m_AccessPath) {
      delete m_AccessPath;
   }
   VariantClear( &m_Value );
   VariantClear( &m_EUInfo );
   DeleteCriticalSection( &m_CritSec );
   DeleteCriticalSection( &m_CritSecAllAttrs );
}


//=========================================================================
// get_RefCount
//=========================================================================
DWORD DaDeviceItem::get_RefCount(void)
{

	return m_RefCount;
}

//=========================================================================
// Get the Item's Id String
// returns the pointer!!! so do not free!
//
// Do not use this function if you implement a CALL-R Server.
//
//=========================================================================
HRESULT DaDeviceItem::get_ItemIDPtr( LPWSTR *ItemID )
{
   *ItemID = m_ItemID ;
   return S_OK;
}



//=========================================================================
// Get the Item's Id String
// ------------------------
//    Returns a copy of the Item ID.
//    The access to the Item ID is protected with a critical section.
//
//       pIMalloc          Allocates string with new() if NULL; otherwise
//                         with pIMalloc->Alloc().
//
//=========================================================================
HRESULT DaDeviceItem::get_ItemIDCopy( LPWSTR *ItemID, IMalloc *pIMalloc /*=NULL*/ )
{
   if (ItemID == NULL) {
      return E_INVALIDARG;
   }

   HRESULT hres = S_OK;
   EnterCriticalSection( &m_CritSec );

   if (m_ItemID) {
      *ItemID = WSTRClone( m_ItemID, pIMalloc );
      if (*ItemID == NULL) {
         hres = E_OUTOFMEMORY;
      }
   }
   else {
      *ItemID = NULL;
   }
   LeaveCriticalSection( &m_CritSec );
   return hres;
}



//=========================================================================
// Set the Item's Id String
// Currently only used when the item is created.
// Makes an internal copy of the passed parameter
//=========================================================================
HRESULT DaDeviceItem::set_ItemID( LPWSTR ItemID )
{
   EnterCriticalSection( &m_CritSec );

   if ( m_ItemID != NULL ) {
      delete m_ItemID;
      m_ItemID = NULL;
   }

   if ( ItemID != NULL ) {
      m_ItemID = WSTRClone( ItemID, NULL );
      if ( m_ItemID == NULL ) {
         LeaveCriticalSection( &m_CritSec );
         return E_OUTOFMEMORY;
      }
   } else {
      m_ItemID = NULL;
   }

   LeaveCriticalSection( &m_CritSec );
   return S_OK;
}




//=========================================================================
//  Attach to class by incrementing the RefCount
//  Returns the current RefCount
//  or -1 if the ToKill flag is set.
//=========================================================================
int  DaDeviceItem::Attach( void ) 
{
      long rc;

   EnterCriticalSection( &m_CritSec );
   if( m_ToKill ) {
      rc = -1 ;
   } else {
      rc = ++m_RefCount ;
   }
   LeaveCriticalSection( &m_CritSec );
   return rc;
}




//=========================================================================
//  Detach from class by decrementing the RefCount.
//  Kill the class instance if request pending.
//  Returns the current RefCount or -1 if it was killed.
//=========================================================================
int DaDeviceItem::Detach( void )
{
      long rc;

   EnterCriticalSection( &m_CritSec );
   if (m_RefCount > 0) {
      --m_RefCount;
   }
   if ((m_RefCount == 0) &&                     // not referenced
       (m_ToKill == TRUE)) {                    // kill request
      LeaveCriticalSection( &m_CritSec );
      delete this;                              // remove from memory
      return -1;
   }
   else {
      rc = m_RefCount;
   }
   LeaveCriticalSection( &m_CritSec );
   return rc;
}



//=========================================================================
//  Kill the class instance.
//  If the RefCount is > 0 then only the kill request flag is set.
//  Returns the current RefCount or -1 if it was killed.
//  If WithDeatch == TRUE then RefCound is first decremented.
//=========================================================================
int  DaDeviceItem::Kill( BOOL WithDetach )
{
      long rc;

   EnterCriticalSection( &m_CritSec );

      //
   if( WithDetach ) {
      if( m_RefCount > 0 )    
         --m_RefCount ;
   }
   if( m_RefCount ) {            // still referenced
      m_ToKill = TRUE ;          // set kill request flag
   }
   else {
      LeaveCriticalSection( &m_CritSec );
      delete this ;
      return -1 ;
   }  
   rc = m_RefCount ;

   LeaveCriticalSection( &m_CritSec );

   return rc;
}


//=========================================================================
// Killed()
// tells whether the class instance was killed (with Kill()) or not
//=========================================================================
BOOL DaDeviceItem::Killed( void )
{
   return m_ToKill;
}



//=========================================================================
// get_Active
//=========================================================================
BOOL   DaDeviceItem::get_Active( void )
{
   if (Killed())
      return FALSE;

   return m_Active;
}



//=========================================================================
// set_Active
//=========================================================================
HRESULT DaDeviceItem::set_Active( BOOL Active )
{
   if (Killed()) {
      return E_FAIL ;
   }

   m_Active = Active;
   return S_OK;
}




//=========================================================================
// AttachActiveCount
//=========================================================================
HRESULT DaDeviceItem::AttachActiveCount( void )
{
   if (Killed()) {
      return E_FAIL;
   }
   m_dwActiveCount++;

   return S_OK;
}




//=========================================================================
// DetachActiveCount
//=========================================================================
HRESULT DaDeviceItem::DetachActiveCount( void )
{
   if (Killed()) {
      return E_FAIL;
   }

   EnterCriticalSection( &m_CritSec );

   if (m_dwActiveCount) {
      m_dwActiveCount--;
   }
   LeaveCriticalSection( &m_CritSec );
   return S_OK;
}




//=========================================================================
// get_ActiveCount
//=========================================================================
DWORD DaDeviceItem::get_ActiveCount( void )
{
   if (Killed()) {
      return 0;
   }
   return m_dwActiveCount;
}




//=========================================================================
// Get the Item's Engineering Unit (EU) Information
// ------------------------------------------------
// The type may be:  OPC:NOENUM (0) :  No EU Info available
//                   OPC_ANALOG (1) :  Analog, EUInfo contains a SAFEARRAY
//                                     of two doubles ( LOW / HI EU range )
//                   OPC_ENUMERATED (2)
//                                  :  Enumerated, EUInfo contains a SAFEARRAY of BSTR
//                                     which contains a list of strings
//                                     ( e.g. "OPEN","CLOSE","IN TRANSIT")
//                                     corresponding to sequential numeric
//                                     values (0,1,2, .. )
// The generic part of the server does only pass the EUInfo to the client.
//=========================================================================
HRESULT DaDeviceItem::get_EUData( OPCEUTYPE *pEUType, VARIANT *pEUInfo )
{
   EnterCriticalSection( &m_CritSec );

   *pEUType = m_EUType;                   // returns the currently set info
   VariantInit( pEUInfo );                // AM, 11-may-98
   HRESULT hres = VariantCopy( pEUInfo, &m_EUInfo );

   LeaveCriticalSection( &m_CritSec );

   return hres;
}     




//=========================================================================
// Set the Item's Engineering Unit (EU) Information
// ------------------------------------------------
// The type may be:  OPC:NOENUM (0) :  No EU Info available
//                   OPC_ANALOG (1) :  Analog, EUInfo contains a SAFEARRAY
//                                     of two doubles ( LOW / HI EU range )
//                   OPC_ENUMERATED (2)
//                                  :  Enumerated, EUInfo contains a SAFEARRAY of BSTR
//                                     which contains a list of strings
//                                     ( e.g. "OPEN","CLOSE","IN TRANSIT")
//                                     corresponding to sequential numeric
//                                     values (0,1,2, .. )
// The generic part of the server does only pass the EUInfo from the client.
//=========================================================================
HRESULT DaDeviceItem::set_EUData( OPCEUTYPE EUType, VARIANT *pEUInfo )
{
   HRESULT  hres;
   VARIANT  varOld;


   #ifdef _DEBUG     // Plausibility checks

      switch (EUType) {
         case OPC_NOENUM:        _ASSERTE( V_VT( pEUInfo ) == VT_EMPTY );              break;
         case OPC_ANALOG:        _ASSERTE( V_VT( pEUInfo ) == (VT_ARRAY | VT_R8) );    break;
         case OPC_ENUMERATED:    _ASSERTE( V_VT( pEUInfo ) == (VT_ARRAY | VT_BSTR) );  break;
         default :               _ASSERT( FALSE );
      }

   #endif            // Plausibility checks

   VariantInit( &varOld );
   EnterCriticalSection( &m_CritSec );

                                       // First save the current value so we can restore it
                                       // if the new value cannot be set.
   hres = VariantCopy( &varOld, &m_EUInfo );

   if (SUCCEEDED( hres )) {
                                       // Clear current value
      hres = VariantClear( &m_EUInfo );
   }
   if (SUCCEEDED( hres )) {
                                       // Set new value
      hres = VariantCopy( &m_EUInfo, pEUInfo );
      if (FAILED( hres )) {
                                       // Restore to old value if new value cannot be set.
                                       // Keep result returned by copy function.
         VariantCopy( &m_EUInfo, &varOld );
      }
      else {
         m_EUType = EUType;            // stores the passed EUInfo

         if (EUType == OPC_ANALOG) {   // store specified range into 'm_dAnalogRange'.
                                       // For this reason the range must not be calcuated at
                                       // every update cycle.
            double   dLow = 0;
            double   dHi  = 0;
            long     lElIndex = 0;

            hres = SafeArrayGetElement( V_ARRAY( pEUInfo ), &lElIndex, &dLow );     // LOW EU range
            if (SUCCEEDED( hres )) {
               lElIndex++;
               hres = SafeArrayGetElement( V_ARRAY( pEUInfo ), &lElIndex, &dHi );   // HI  EU range
            }
            if (SUCCEEDED( hres )) {
               m_dAnalogEURange = dHi - dLow;
            }
            else {
                                       // Restore to old value if EU info cannotbe read.
                                       // Keep result returned by EU access function.
               VariantCopy( &m_EUInfo, &varOld );
            }
         } // EU Type is Analog
      }
   }

   LeaveCriticalSection( &m_CritSec );
   VariantClear( &varOld );            // Clear temporary variant

   return hres;
}




//=========================================================================
// get_AnalogEURange
//=========================================================================
HRESULT DaDeviceItem::get_AnalogEURange( double* pdAnalogEURange )
{
   if (Killed()) {
      return E_FAIL;
   }

   EnterCriticalSection( &m_CritSec );

   HRESULT hres = S_OK;

   if (m_EUType != OPC_ANALOG) {
      hres = E_FAIL;
   }
   else {
      *pdAnalogEURange = m_dAnalogEURange;
   }
   LeaveCriticalSection( &m_CritSec );

   return hres;
}



//=========================================================================
// Get the Item's CanonicalDataType
// --------------------------------
// This is the data type used in the cache. The hardware access driver
// works with this data type. Read/Write functions may convert the type
// to the type requested/provided by the client.
//=========================================================================
VARTYPE DaDeviceItem::get_CanonicalDataType( void )
{
   return V_VT(&m_Value);       // return the current VARIANT type
}



//=========================================================================
// Set the Item's CanonicalDataType
// --------------------------------
// This is the data type used in the cache. The hardware access driver
// works with this data type. Read/Write functions may convert the type
// to the type requested/provided by the client.
//=========================================================================
HRESULT DaDeviceItem::set_CanonicalDataType( VARTYPE CanonicalDataType )
{
   // check if the change is possible 
   // and if so, then convert with VariantChangeType()
   return E_FAIL ;   // Type changes not supported by this server implementation
}




//=========================================================================
// Get the Item's current DaAccessRights
// -----------------------------------
// OPC supported Access rights are:  OPC_READABLE, OPC_WRITEABLE
//=========================================================================
HRESULT DaDeviceItem::get_AccessRights( DWORD *pAccessRights )
{
   *pAccessRights = m_AccessRights;
   return S_OK;
}



//=========================================================================
// Set the Item's DaAccessRights
// ---------------------------
// OPC supported Access rights are:  OPC_READABLE, OPC_WRITEABLE
//=========================================================================
HRESULT DaDeviceItem::set_AccessRights( DWORD DaAccessRights )
{
   m_AccessRights = DaAccessRights;
   return S_OK;
}



//=========================================================================
// get_OPCITEMRESULT
// -----------------
// Returns the Item Attributes as OPCITERESULT.
// Note: The OPCITEMRESULT member hServer is always 0 because available
// only via GenericItem.
//=========================================================================
HRESULT DaDeviceItem::get_OPCITEMRESULT( BOOL blobUpdate, OPCITEMRESULT* pItemResult )
{
   HRESULT hres = S_OK;

   EnterCriticalSection( &m_CritSecAllAttrs );
   EnterCriticalSection( &m_CritSec );

   pItemResult->hServer             = 0;
   pItemResult->vtCanonicalDataType = V_VT( &m_Value );
   pItemResult->wReserved           = 0;
   pItemResult->dwAccessRights      = m_AccessRights;

   if (blobUpdate) {
      hres = get_Blob( &pItemResult->pBlob, &pItemResult->dwBlobSize );
   }
   else {
      pItemResult->pBlob            = NULL;
      pItemResult->dwBlobSize       = 0;
   }

   LeaveCriticalSection( &m_CritSec );
   LeaveCriticalSection( &m_CritSecAllAttrs );

   if (FAILED( hres )) {
      pItemResult->hServer             = 0;
      pItemResult->vtCanonicalDataType = VT_EMPTY;
      pItemResult->wReserved           = 0;
      pItemResult->dwAccessRights      = 0;
      pItemResult->dwBlobSize          = 0;
      if (pItemResult->pBlob) {
         pIMalloc->Free( pItemResult->pBlob );
         pItemResult->pBlob = NULL;
      }
   }
   return hres;
}



//=========================================================================
// Read current item value, quality and time stamp from the cache
// --------------------------------------------------------------
// (The values stored in the DaDeviceItem member variables.)
// Refreshing the cache is not part of this function. 
// This is done in DaBaseServer::RefreshInputCache() in module
// DaServer.cpp.
// 
// If the client calls read through functions (Source == Device) then the
// generic server part first calles the
// DaBaseServer::RefreshInputCache() method.
//
// Value.vt contains the requested data type to which the returned value
// should be converted from the cache value (which has canonical data type)
//   
//=========================================================================
HRESULT DaDeviceItem::get_ItemValue( LPVARIANT   pvValue,
                                    LPWORD      pwQuality,
                                    LPFILETIME  pftTimeStamp )
{
   _ASSERTE( pvValue );                         // Must not be NULL
   _ASSERTE( pwQuality );                       // Must not be NULL
   _ASSERTE( pftTimeStamp );                    // Must not be NULL

   HRESULT  hr;                                 // Store the requested data type.
   VARTYPE  vtRequestedDataType = V_VT( pvValue );

   VariantInit( pvValue );                      // Initialze the destination variant. Only the
                                                // 'vt' data member with requested data type was valid.
   EnterCriticalSection( &m_CritSec );

   hr = VariantFromVariant(   pvValue,
                              vtRequestedDataType, &m_Value );

   if (SUCCEEDED( hr )) {
      *pwQuality     = m_Quality;
      *pftTimeStamp  = m_TimeStamp;
   }
   LeaveCriticalSection( &m_CritSec );
   return hr;
}



//=========================================================================
// Cache Update: Value, Quality and TimeStamp
// ------------------------------------------
// (The values stored in the DaDeviceItem member variables.)
// Doing the actual device write is not part of this method.
// This is done in DaBaseServer::RefreshOutputDevices() in
// module DaServer.cpp.
// Depending on the application this method may need to force it.
//
// This function requires a value with the canonical data type.
//=========================================================================
HRESULT DaDeviceItem::set_ItemValue( LPVARIANT   pvValue,
                                    WORD        wQuality,      /* = OPC_QUALITY_GOOD | OPC_LIMIT_OK */
                                    LPFILETIME  pftTimeStamp   /* = NULL */ )
{
   //_ASSERTE( pvValue );                         // Must not be NULL
   //_ASSERTE( get_CanonicalDataType() == V_VT( pvValue ) );

   if (pvValue == NULL || get_CanonicalDataType() != V_VT(pvValue))
   {
	   return S_FALSE;
   }

   FILETIME ftTimeStamp;
   HRESULT  hres;

   if (!pftTimeStamp) {
      hres = CoFileTimeNow( &ftTimeStamp );
      if (FAILED( hres )) return hres;
   }
   else {
      ftTimeStamp = *pftTimeStamp;
   }

   EnterCriticalSection( &m_CritSec );
                                                // Note : VariantCopy() frees the destination variant if all succeeded.
   hres = VariantCopy( &m_Value, pvValue );     // Set the new value
   if (SUCCEEDED( hres )) {
      m_Quality   = wQuality;
      m_TimeStamp = ftTimeStamp;
   }

   LeaveCriticalSection( &m_CritSec );
   return hres;
}



//=========================================================================
// Cache Update: Quality and TimeStamp
// -----------------------------------
// Changes only the quality and the time stamp and let the item
// value unchanged.
//=========================================================================
HRESULT DaDeviceItem::set_ItemQuality(  WORD        wQuality,
                                       LPFILETIME  pftTimeStamp /* = NULL */ )
{
   if (!pftTimeStamp) {
      FILETIME ftTimeStamp;
                                                // Get current time
      HRESULT hres = CoFileTimeNow( &ftTimeStamp );
      if (FAILED( hres )) return hres;

      EnterCriticalSection( &m_CritSec );
      m_Quality   = wQuality;
      m_TimeStamp = ftTimeStamp;
      LeaveCriticalSection( &m_CritSec );
   }
   else {
      EnterCriticalSection( &m_CritSec );
      m_Quality   = wQuality;
      m_TimeStamp = *pftTimeStamp;
      LeaveCriticalSection( &m_CritSec );
   }
   return S_OK;
}



//=========================================================================
// Cache Update: Value, Quality and TimeStamp
// ------------------------------------------
//    Value    
//       The value is only updated if OPCITEMVQT::vDataValue is not of
//       type VT_EMPTY.
//
//    Quality
//       The Quality is upadated with OPCITEMVQT::wQuality if
//       OPCITEMVQT::bQualitySpecified is TRUE,; otherwise the new
//       Quality is OPC_QUALITY_GOOD | OPC_LIMIT_OK.
//
//    TimeStamp
//       The TimeStamp is upadated with OPCITEMVQT::ftTimeStamp if
//       OPCITEMVQT::bTimeStampSpecified is TRUE; otherwise the current
//       time is used for the new TimeStamp.
//=========================================================================
HRESULT DaDeviceItem::set_ItemVQT( OPCITEMVQT* pItemVQT )
{
   _ASSERTE( pItemVQT  );                       // Must not be NULL

   WORD     wQuality;
   FILETIME ftTimeStamp;
   HRESULT  hr = S_OK;

   if (pItemVQT->bTimeStampSpecified)
      ftTimeStamp = pItemVQT->ftTimeStamp;
   else {
      hr = CoFileTimeNow( &ftTimeStamp );
      if (FAILED( hr )) return hr;
   }

   if (pItemVQT->bQualitySpecified)
      wQuality = pItemVQT->wQuality;
   else
      wQuality = OPC_QUALITY_GOOD | OPC_LIMIT_OK;
  
   if (V_VT( &pItemVQT->vDataValue ) == VT_EMPTY) {
      EnterCriticalSection( &m_CritSec );
      m_Quality   = wQuality;
      m_TimeStamp = ftTimeStamp;
      LeaveCriticalSection( &m_CritSec );
   }
   else {
      _ASSERTE( get_CanonicalDataType() == V_VT( &pItemVQT->vDataValue ) );
      EnterCriticalSection( &m_CritSec );
                                                // Note : VariantCopy() frees the destination variant if all succeeded.
                                                // Set the new value
      hr = VariantCopy( &m_Value, &pItemVQT->vDataValue );     
      if (SUCCEEDED( hr )) {
         m_Quality   = wQuality;
         m_TimeStamp = ftTimeStamp;
      }
      LeaveCriticalSection( &m_CritSec );
   }
   return hr;   
}



//=========================================================================
// Converts a variant value from one type to the canonical data type
// -----------------------------------------------------------------
// This function checks the type of the converted value because it may
// happen that VariantChangeType() returns SUCCEEDED but the  conversion
// failed and the value type is VT_EMPTY.
// Also DISP_E_xxx error codes are converted to OPC error codes.
//=========================================================================
HRESULT DaDeviceItem::ChangeValueToCanonicalType( LPVARIANT pvValue )
{
   if (V_VT( pvValue ) == VT_EMPTY) {
      return OPC_E_BADTYPE;
   }

   VARIANT vTmpNew;
   VariantInit( &vTmpNew );

   HRESULT hr = VariantFromVariant( &vTmpNew,
                                    get_CanonicalDataType(), pvValue );

   if (SUCCEEDED( hr )) {
      // shallow copy, faster than VariantCopy()
      VariantClear( pvValue );
      memcpy( pvValue, &vTmpNew, sizeof (VARIANT) );
   }
   return hr;
}



//=========================================================================
// Compares the current TimeStamp with the specified value
//=========================================================================
BOOL DaDeviceItem::IsTimeStampOlderThan( LPFILETIME pftTimeStamp )
{
   _ASSERTE( pftTimeStamp );                    // Must not be NULL
   
   EnterCriticalSection( &m_CritSec );
   LONG lRes = CompareFileTime( &m_TimeStamp, pftTimeStamp );
   LeaveCriticalSection( &m_CritSec );

   return (lRes == -1) ? TRUE : FALSE;
}

   
   
//=========================================================================
// get_PropertyValue
// -----------------
//    This function handles the properties with the IDs OPC_PROP_HIEU,
//    OPC_PROP_LOEU and all from the ID Set 1 'OPC Specific Properties'.
//    Other IDs must be handled in the application specific server part.
//=========================================================================
HRESULT DaDeviceItem::get_PropertyValue( DaBaseServer* const pServerHandler,
                                        DWORD dwPropID,
                                        LPVARIANT pvPropData )
{
   _ASSERTE( pServerHandler );
   _ASSERTE( pvPropData );

   HRESULT hres = S_OK;

   // 22-jan-2002 MT
   // Change citicalsections to prevent dead-locks.
   // If OnRefreshInputCache() is called within CriticalSection,
   // then a dead-lock can occur if the refresh-thread is using
   // the same DaDeviceItem at the same time.
   // EnterCriticalSection( &criticalSection_ );       // 22-jan-2002

   switch (dwPropID) {

      case OPC_PROPERTY_DATATYPE :              // Canonical Data Type
         EnterCriticalSection( &m_CritSec );    // 22-jan-2002 MT
         V_VT( pvPropData ) = VT_I2;
         V_I2( pvPropData ) = V_VT( &m_Value );
         LeaveCriticalSection( &m_CritSec );    // 22-jan-2002 MT
         break;

      case OPC_PROPERTY_VALUE :
      case OPC_PROPERTY_QUALITY :
      case OPC_PROPERTY_TIMESTAMP :
         {
                                                // Refresh Cache from Device
            DaDeviceItem* pDevItem = this;
            pServerHandler->OnRefreshInputCache( OPC_REFRESH_CLIENT, 1, &pDevItem, &hres );
            if (SUCCEEDED( hres )) {            // Use the individual item error as return code
                                                // Cache refresh succeeded

               pServerHandler->readWriteLock_.BeginReading();
               EnterCriticalSection( &m_CritSec ); // 22-jan-2002 MT
               switch (dwPropID) {

                  case OPC_PROPERTY_VALUE :
                     hres = VariantCopy( pvPropData, &m_Value );
                     break;

                  case OPC_PROPERTY_QUALITY :
                     V_VT( pvPropData ) = VT_I2;
                     V_I2( pvPropData ) = m_Quality;
                     break;

                  case OPC_PROPERTY_TIMESTAMP :
                     V_VT( pvPropData ) = VT_DATE;
                     hres = FileTimeToDATE( &m_TimeStamp, V_DATE( pvPropData ) );
                     break;

                  default :
                     _ASSERTE( 0 );
                     break;
               }
                                                
               LeaveCriticalSection( &m_CritSec ); // 22-jan-2002 MT
               pServerHandler->readWriteLock_.EndReading();
            }
         }
         break;

      case OPC_PROPERTY_ACCESS_RIGHTS :
         EnterCriticalSection( &m_CritSec );    // 22-jan-2002 MT
         V_VT( pvPropData ) = VT_I4;
         V_I4( pvPropData ) = m_AccessRights;
         LeaveCriticalSection( &m_CritSec );    // 22-jan-2002 MT
         break;

      case OPC_PROPERTY_SCAN_RATE :
         EnterCriticalSection( &m_CritSec );    // 22-jan-2002 MT
         V_VT( pvPropData ) = VT_R4;
         V_R4( pvPropData ) = (float)pServerHandler->GetBaseUpdateRate();
         LeaveCriticalSection( &m_CritSec );    // 22-jan.2002 MT
         break;

      case OPC_PROPERTY_EU_TYPE :
         EnterCriticalSection( &m_CritSec );
         V_VT( pvPropData ) = VT_I4;
         V_I4( pvPropData ) = m_EUType;
         LeaveCriticalSection( &m_CritSec );
         break;

      case OPC_PROPERTY_EU_INFO :
         OPCEUTYPE dwEUType;
         hres = get_EUData( &dwEUType, pvPropData );
         if (SUCCEEDED( hres ) && dwEUType != OPC_ENUMERATED) {
            VariantClear( pvPropData );
         }
         break;

      #ifndef _NO_DEADBAND_HANDLING
      case OPC_PROPERTY_HIGH_EU :
      case OPC_PROPERTY_LOW_EU :
         EnterCriticalSection( &m_CritSec );    // 22-jan-2002 MT
         if (m_EUType == OPC_ANALOG) {
            long lElIndex = (dwPropID == OPC_PROPERTY_HIGH_EU) ? 1 : 0;
            V_VT( pvPropData ) = VT_R8;
            hres = SafeArrayGetElement( V_ARRAY( &m_EUInfo ), &lElIndex, &V_R8( pvPropData )  );
         }
         LeaveCriticalSection( &m_CritSec );    // 22-jan-2002 MT
         break;
      #endif

      default :                                 // It is a vendor specific property.
         hres = E_FAIL;
         break;

   } // switch Property ID

   // 22-jan-2002 MT
   // See comment at function entry
   // LeaveCriticalSection( &criticalSection_ );

   if (FAILED( hres )) {
      VariantClear( pvPropData );
   }
   return hres;
}



//=========================================================================
//  Return the currently defined AccessPath.
//  ----------------------------------------
//  It is application specific how the AccessPath is used.
//  The generic part of the OPC server does not handle it in any other 
//  way then passing it from this method to the client.
//
//  In this server implementation there is only one accesspath for each 
//  device item and this will be used when reading or writing
//
//  Use  new/delete  or  WSTRClone(...,NULL)/WSTRFree(...,NULL)  to
//  allocate and deallocate memory
//=========================================================================
HRESULT DaDeviceItem::get_AccessPath( LPWSTR *AccessPath )
{
   if ( AccessPath == NULL ) {
      return E_INVALIDARG;
   }

   EnterCriticalSection( &m_CritSec );

   if ( m_AccessPath != NULL ) {
      *AccessPath = WSTRClone( m_AccessPath, NULL );
      if ( *AccessPath == NULL ) {
         LeaveCriticalSection( &m_CritSec );
         return E_OUTOFMEMORY;
      }
   } else {
      *AccessPath = NULL;
   }

   LeaveCriticalSection( &m_CritSec );

   return S_OK;
}



//=========================================================================
//  Set a new AccessPath.
//  ---------------------
//  It is application specific how the AccessPath is used.
//  The generic part of the OPC server does not handle it in any other 
//  way then passing it from the client to this method.
//  Depending on the application it may be necessary to check if
//  the AccessPath is valid.
//
//  Use  new/delete  or  WSTRClone(...,NULL)/WSTRFree(...,NULL)  to
//  allocate and deallocate memory
//=========================================================================
HRESULT DaDeviceItem::set_AccessPath( LPWSTR AccessPath )
{
      
   //  check if the path is valid may be required

   EnterCriticalSection( &m_CritSec );

   if( m_AccessPath != NULL ) {        // already a path defined
      delete m_AccessPath;
   }

   if( AccessPath == NULL) {           // no path definition
      m_AccessPath = NULL;             // clear member variable
   } else {                              // Path definition passed
      m_AccessPath = WSTRClone( AccessPath, NULL );
      if ( m_AccessPath == NULL ) {
         LeaveCriticalSection( &m_CritSec );
         return E_OUTOFMEMORY;         // error
      }
   }

   LeaveCriticalSection( &m_CritSec );

   return S_OK;
}



//=========================================================================
// Return the currently defined Blob
// ---------------------------------
//  It is application specific how the Blob is used.
//  The generic part of the OPC server does not handle it in any other 
//  way then passing it from this method to the client.
//  Outputs:
//    It is expected that a copy of the "internal" blob is returned.
//    Memory for the copy must be allocated from the
//    Global COM Memory Manager.
//    That is the returned blob belongs to the caller.
//    This is because we need to synchronize access at device item level.
//=========================================================================
HRESULT DaDeviceItem::get_Blob( BYTE **pBlob, DWORD *BlobSize )
{
   *pBlob    = NULL;          // NULL returned
   *BlobSize = 0 ;            // since this server does not use/support Blobs
   return S_OK;
}



//=========================================================================
//  Set a new Blob
//  --------------
//  It is application specific how the Blob is used.
//  The generic part of the OPC server does not handle it in any other 
//  way then passing it from this method to the client.
//
//  Inputs:
//    It is expected that a copy is done of the passed Blob
//    That is, the input Blob still belongs to the caller
//=========================================================================
HRESULT DaDeviceItem::set_Blob( BYTE *pBlob, DWORD BlobSize )
{
          // ignored since this server does not use/support Blobs
   return E_NOTIMPL; 
}



//=========================================================================
// Sets the PercentDeadband of this item
//=========================================================================
HRESULT DaDeviceItem::SetItemDeadband( FLOAT fltPercentDeadband )
{
   if (fltPercentDeadband < 0.0 || fltPercentDeadband > 100.0) {
      return E_INVALIDARG;
   }

   HRESULT hr = S_OK;
   EnterCriticalSection( &m_CritSec );

   if  (m_EUType != OPC_ANALOG) {
      hr = OPC_E_DEADBANDNOTSUPPORTED;
   }
   else {
      m_fltPercentDeadband = fltPercentDeadband;
   }

   LeaveCriticalSection( &m_CritSec );
   return hr;
}



//=========================================================================
// Gets the PercentDeadband of this item
//=========================================================================
HRESULT DaDeviceItem::GetItemDeadband( FLOAT* pfltPercentDeadband )
{
   _ASSERTE( pfltPercentDeadband );             // Must not be NULL

   HRESULT hr = S_OK;
   EnterCriticalSection( &m_CritSec );

   if  (m_EUType != OPC_ANALOG) {
      hr = OPC_E_DEADBANDNOTSUPPORTED;
   }
   else if (m_fltPercentDeadband < 0.0) {
      hr = OPC_E_DEADBANDNOTSET;
   }
   else {
      *pfltPercentDeadband = m_fltPercentDeadband;
   }
   
   LeaveCriticalSection( &m_CritSec );
   return hr;
}



//=========================================================================
// Clears the PercentDeadband of this item
//=========================================================================
HRESULT DaDeviceItem::ClearItemDeadband()
{
   HRESULT hr = S_OK;
   EnterCriticalSection( &m_CritSec );

   if  (m_EUType != OPC_ANALOG) {
      hr = OPC_E_DEADBANDNOTSUPPORTED;
   }
   else if (m_fltPercentDeadband < 0.0) {
      hr = OPC_E_DEADBANDNOTSET;
   }
   else {
      m_fltPercentDeadband = -1;
   }
   
   LeaveCriticalSection( &m_CritSec );
   return hr;
}

//DOM-IGNORE-END
