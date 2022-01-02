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

#ifdef   _OPC_SRV_AE                            // Alarms & Events Server

//DOM-IGNORE-BEGIN
//-----------------------------------------------------------------------
// INLCUDE
//-----------------------------------------------------------------------
#include "stdafx.h"
#include "UtilityFuncs.h"
#include "AeSource.h"
#include "AeCategory.h"
#include "AeCondition.h"
#include "AeConditionDefinition.h"
#include "AeAttribute.h"
#include "AeEvent.h"

//-------------------------------------------------------------------------
// CODE AeEvent
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AeEvent::AeEvent()
{
   AddRef();

   szConditionName      = NULL;
   szSubconditionName   = NULL;
   szActorID            = NULL;
   szSource             = NULL;
   szMessage            = NULL;
   dwNumEventAttrs      = 0;
   pEventAttributes     = NULL;
}



//=========================================================================
// Initializer
// -----------
//    This function creates a simple or tracking event and must be
//    called after construction.
//    Condition-related events are initialized with the function
//    AeCondition::CreateEventInstance() which is a fried function 
//    of this class.
//
// Parameters:
//    pCat                    The Category associated withthe event.
//    pSource                 The Source associated withthe event.
//
//    szParMessage            Text which describes the event.
//
//    dwParSeverity           The urgency of the event.
//
//    szParActorID            If NULL-Pointer, a simple event will be
//                            generated; otherwise a tracking event.
//
//    pvAttrValues            Array of attribute values. The number of
//                            values and the order of the types must be
//                            identical with the attributes of the
//                            specified category.
//
//    pft                     Time that the event occured. If NULL-Pointer
//                            then the current time will be used.
//
//=========================================================================
HRESULT AeEvent::Create( AeCategory* pCat, AeSource *pSource,
                          LPCWSTR szParMessage, DWORD dwParSeverity, LPCWSTR szParActorID,
                          LPVARIANT pvAttrValues, LPFILETIME pft )
{
   _ASSERTE( pCat );
   _ASSERTE( pSource );
   _ASSERTE( szParMessage );

   HRESULT  hres = S_OK;
   DWORD*   pdwAttrIDs = NULL;

   try {

      // initialize members which are different between simple and tracking events
      if (szParActorID) {
         szActorID = WSTRClone( szParActorID );
         dwEventType = OPC_TRACKING_EVENT;
      }
      else {
         szActorID = WSTRClone( L"" );
         dwEventType = OPC_SIMPLE_EVENT;
      }
      if (dwEventType != pCat->EventType())
         throw E_FAIL;                          // Invalid Category specified for this type of event


      // initialize members which are not required for simple and tracking events
      wChangeMask          = 0;
      wNewState            = 0;
      szConditionName      = WSTRClone( L"" );
      szSubconditionName   = WSTRClone( L"" );
      wQuality             = 0;
      wReserved            = 0;
      bAckRequired         = 0;
      dwCookie             = 0;
      memset( &ftActiveTime, 0 , sizeof (ftActiveTime) );

      // set the occurence time of the event
      if (pft) {
         ftTime = *pft;
      }
      else {
         hres = CoFileTimeNow( &ftTime );
         if (FAILED( hres )) throw hres;
      }
      szSource             = pSource->Name().Copy();
      szMessage            = WSTRClone( szParMessage );
      dwEventCategory      = pCat->CatID();
      dwSeverity           = dwParSeverity;

      if ((szConditionName == NULL) || (szSubconditionName == NULL) ||
          (szActorID == NULL) || (szSource == NULL) || (szMessage == NULL)) {
         throw E_OUTOFMEMORY;
      }

      // Add all attribute IDs and values
      // The IDs are from the specified category and the current values from the user.
      // The values must be in the same order and have the same types as specified by
      // the category.

      DWORD dwCount;

      hres = pCat->GetAttributeIDs( &dwCount, &pdwAttrIDs );
      if (FAILED( hres )) throw hres;

      hres = m_mapAttrValues.Create( dwCount );
      if (FAILED( hres )) throw hres;

      // Initialize the vendor specific attributes
      DWORD i;
      for (i=0; i < dwCount; i++) {

         switch (pdwAttrIDs[i]) {

            case ATTRID_ACKCOMMENT:             // Default attribute 'ACK COMMENT'
                                                // There is no acknowledge comment for simple and tracking events.
            #if ((defined(CE_TSUNSOL2)    || CE_OSUNSOL2)      || \
                 (defined(CE_TSUNSOL7_32) || CE_OSUNSOL7_32)   || \
                 (defined(CE_TSUNSOL7_64) || CE_OSUNSOL7_64)   || \
                 (defined(CE_TSUNSOL8_64) || CE_OSUNSOL8_64))
               //
               // Sun Solaris on SPARC hardware.
               //
               {
                  _variant_t vTmp( "" );
                  hres = m_mapAttrValues.SetAtIndex( i, pdwAttrIDs[i], &vTmp );
               }
            #else
               //
               // Not Sun Solaris.
               //
               hres = m_mapAttrValues.SetAtIndex( i, pdwAttrIDs[i], &_variant_t( "" ) );
            #endif
               if (FAILED( hres )) throw hres;
               break;

            case ATTRID_AREAS:                  // Default attribute 'AREAS'
               {
                  VARIANT vAreas;
                  hres = pSource->GetAreaNames( &vAreas );
                  if (FAILED( hres )) throw hres;

                  hres = m_mapAttrValues.SetAtIndex( i, pdwAttrIDs[i], &vAreas );
                  VariantClear( &vAreas );
                  if (FAILED( hres )) throw hres;
               }
               break;

            default:
               hres = m_mapAttrValues.SetAtIndex( i, pdwAttrIDs[i], &pvAttrValues[i - pCat->NumOfInternalAttrs()] );
               if (FAILED( hres )) throw hres;
               break;
         }
      }
   }
   catch( HRESULT hresEx ) {
      Cleanup();
      hres = hresEx;                            // catch own exceptions
   }
   catch(...) {                                 // catch all other exception
      Cleanup();                                // e.g. if the attribute value array is invalid
      hres = E_FAIL;
   }

   if (pdwAttrIDs)
      delete [] pdwAttrIDs;

   return hres;
}



//-------------------------------------------------------------------------
// IMPLEMENTATION
//-------------------------------------------------------------------------

//=========================================================================
// Cleanup
// -------
//    All resources used by this this instance are released.
//    This function is called by the destructor or if
//    not all succeeded during initialization.
//=========================================================================
void AeEvent::Cleanup()
{
   if (szConditionName) {
      delete szConditionName;
      szConditionName = NULL;
   }
   if (szSubconditionName) {
      delete szSubconditionName;
      szSubconditionName = NULL;
   }
   if (szActorID) {
      delete szActorID;
      szActorID = NULL;
   }
   if (szSource) {
      delete szSource;
      szSource = NULL;
   }
   if (szMessage) {
      delete szMessage;
      szMessage = NULL;
   }
   m_mapAttrValues.~AttributeValueMap();            // remove all added values
}



//-------------------------------------------------------------------------
// CODE AeEventArray
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AeEventArray::AeEventArray()
{
   m_dwMax = 0;
   m_dwNum = 0;
   m_parOnEvent = NULL;
}



//=========================================================================
// Initializer
// -----------
//    Must be called after construction.
//    Parameter dwMax specifies the maximum number of elements
//    which can be added to the array.
//=========================================================================
HRESULT AeEventArray::Create( DWORD dwMax )
{
   if (m_parOnEvent)
      return E_FAIL;                            // already created

   m_parOnEvent = new AeEvent* [dwMax];
   if (!m_parOnEvent) return E_OUTOFMEMORY;

   m_dwMax = dwMax;
   return S_OK;
}



//=========================================================================
// Destructor
//=========================================================================
AeEventArray::~AeEventArray()
{
   if (m_parOnEvent) {
      while (m_dwNum--) {
         m_parOnEvent[m_dwNum]->Release();
      }
      delete [] m_parOnEvent;
   }
}



//-------------------------------------------------------------------------
// OPERATION
//-------------------------------------------------------------------------

//=========================================================================
// Add
// ---
//    Adds an new element to the array. The maximum number of elements is
//    specified by the Create() function.
//=========================================================================
HRESULT AeEventArray::Add( AeEvent* pEvent )
{
   _ASSERTE( m_dwNum < m_dwMax );
   m_parOnEvent[m_dwNum++] = pEvent;
   return S_OK;
}



//-------------------------------------------------------------------------
// CODE AeSubscribedEvent
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AeSubscribedEvent::AeSubscribedEvent()
{
   memset( (ONEVENTSTRUCT *)this, 0, sizeof (ONEVENTSTRUCT) );
}



//=========================================================================
// Initializer
// -----------
//    Must be called after construction.
//    Creates a shallow copy of the specified AeEvent instance.
//    The event attributes are shallow copies of the attribute 
//    values associated with the specified IDs.
//=========================================================================
HRESULT AeSubscribedEvent::Create( AeEvent* pOnEvent, DWORD dwNumOfAttrIDs, DWORD dwAttrIDs[] )
{
   DWORD       i;
   LPVARIANT   pv;

   *(ONEVENTSTRUCT *)this = *pOnEvent;       // shallow copy 

                                             // shallow copy of selected attributes
   pEventAttributes = new VARIANT [dwNumOfAttrIDs];
   if (!pEventAttributes) return E_OUTOFMEMORY;

   dwNumEventAttrs = dwNumOfAttrIDs;

   for (i=0; i < dwNumEventAttrs; i++) {
      pv = pOnEvent->LookupAttributeValue( dwAttrIDs[i] );
      if (pv)  pEventAttributes[i] = *pv;
      else     VariantInit( &pEventAttributes[i] );
   }
   return S_OK;
}



//=========================================================================
// Destructor
//=========================================================================
AeSubscribedEvent::~AeSubscribedEvent()
{
   if (pEventAttributes) {
      delete [] pEventAttributes;
   }
}
//DOM-IGNORE-END
#endif