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
//-------------------------------------------------------------------------
// INLCUDE
//-------------------------------------------------------------------------
#include "stdafx.h"
#include "FixOutArray.h"
#include "AeCondition.h"
#include "AeEventArea.h"
#include "AeEvent.h"

//-------------------------------------------------------------------------
// CODE
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AeSource::AeSource()
{
}



//=========================================================================
// Initializer
// -----------
//    Must be called after construction.
//
//    If fIsFullyQualified is TRUE then parameter szName specifies a
//    fully qualified source name; otherwise szName specifies only
//    a partial source name and this function creates the fully qualified
//    souce name with the fully qualified name of the specified
//    area object.
//=========================================================================
HRESULT AeSource::Create( EventArea* pArea, LPCWSTR szName, BOOL fIsFullyQualified )
{
   _ASSERTE( pArea );

   HRESULT hres;

   // generate the fully qualified source name
   if (pArea->IsRoot() || fIsFullyQualified) {
      hres = m_wsName.SetString( szName );
      m_pszPartialName = m_wsName;              // Initialize pointer to the partial area name
   }
   else {
                                                // get the fully qualified name of the area
                                                // and append the name of the source
      LPWSTR szQualifiedName = new WCHAR [wcslen( pArea->Name() ) + wcslen( szName ) + 2 ];
                                                // +2 for EOS and the delimiter character
      if (!szQualifiedName) {
         hres = E_OUTOFMEMORY;
      }
      else {
         wcscpy( szQualifiedName, pArea->Name() );
         wcscat( szQualifiedName, pArea->Delimiter() );
         wcscat( szQualifiedName, szName );
         hres = m_wsName.SetString( szQualifiedName );
         delete [] szQualifiedName;             // Temporary area name no longer used
                                                // Initialize pointer to the partial area name
         m_pszPartialName = m_wsName + wcslen( pArea->Name() ) + 1;
      }
   }

   // attach this new source instance to the area
   if (SUCCEEDED( hres )) {
      hres = pArea->AttachSource( this );
   }
   if (SUCCEEDED( hres )) {
      m_csAreaRefs.Lock();                      // successfully attached
      hres = m_arAreaRefs.Add( pArea ) ? S_OK : E_OUTOFMEMORY;
      if (FAILED( hres )) {
         pArea->DetachSource( this );
      }
      m_csAreaRefs.Unlock();
   }
   return hres;
}



//=========================================================================
// Destructor
//=========================================================================
AeSource::~AeSource()
{
   m_csAreaRefs.Lock();

   int i=m_arAreaRefs.GetSize();
   while (i--) {
      m_arAreaRefs[i]->DetachSource( this );
   }
   m_arAreaRefs.RemoveAll();

   m_csAreaRefs.Unlock();
}



//-------------------------------------------------------------------------
// OPERATIONS
//-------------------------------------------------------------------------

//=========================================================================
// AddToAdditionalArea
// -------------------
//    Adds this source instance to an additional area.
//=========================================================================
HRESULT AeSource::AddToAdditionalArea( EventArea* pArea )
{
   _ASSERTE( pArea );

   HRESULT hres;

   m_csAreaRefs.Lock();                      // successfully attached

   if (m_arAreaRefs.GetSize() < 1) {
      hres = E_FAIL;                         // not yet created successfully
   }
   else {
      // attach this source instance to an additional area
      hres = pArea->AttachSource( this );

      if (SUCCEEDED( hres )) {
         hres = m_arAreaRefs.Add( pArea ) ? S_OK : E_OUTOFMEMORY;
         if (FAILED( hres )) {
            pArea->DetachSource( this );
         }
      }
   }

   m_csAreaRefs.Unlock();
   return hres;
}



//=========================================================================
// AttachCondition
// ---------------
//    Attaches a link to a condition to this source.
//=========================================================================
HRESULT AeSource::AttachCondition( AeCondition* pCond )
{
   _ASSERTE( pCond != NULL );                   // Must not be NULL

   HRESULT hres;
   m_csCondRefs.Lock();
   hres = m_arCondRefs.Add( pCond ) ? S_OK : E_OUTOFMEMORY;
   m_csCondRefs.Unlock();
   return hres;
}



//=========================================================================
// DetachCondition
// ---------------
//    Detaches a link to a condition from this source.
//=========================================================================
HRESULT AeSource::DetachCondition( AeCondition* pCond )
{
   _ASSERTE( pCond != NULL );                   // Must not be NULL

   HRESULT hres;
   m_csCondRefs.Lock();
   hres = m_arCondRefs.Remove( pCond ) ? S_OK : E_FAIL;
   m_csCondRefs.Unlock();
   return hres;
}



//=========================================================================
// QueryConditionNames
// -------------------
//    Gets the names of all conditions attached to this source.
//=========================================================================
HRESULT AeSource::QueryConditionNames(
                                    DWORD    *  pdwCount,
                                    LPWSTR   ** ppszConditionNames )
{
   HRESULT                 hres = S_OK;
   CFixOutArray< LPWSTR >  aName;

   m_csCondRefs.Lock();
   *pdwCount = m_arCondRefs.GetSize();

   try {                                        // Allocate and initialize the result arrays
      aName.Init( *pdwCount, ppszConditionNames );
                                                // Fill the result arrays with the
      for (DWORD i=0; i < *pdwCount; i++) {     // name of all condition definitions
         aName[i] = m_arCondRefs[i]->Name().CopyCOM();
         if (aName[i] == NULL) throw E_OUTOFMEMORY;
      }
   }
   catch (HRESULT hresEx) {
      *pdwCount = 0;
      aName.Cleanup();
      hres = hresEx;
   }
   m_csCondRefs.Unlock();

   return hres;
}



//=========================================================================
// GetConditionState
// -----------------
//    Gets the stat of the condition inclusive the specified attaributes.
//=========================================================================
HRESULT AeSource::GetConditionState(
   /* [in] */                    LPWSTR         szConditionName,
   /* [in] */                    DWORD          dwNumEventAttrs,
   /* [size_is][in] */           DWORD       *  pdwAttributeIDs,
   /* [out] */                   OPCCONDITIONSTATE
                                             ** ppConditionState )
{
   HRESULT hres = E_INVALIDARG;                 // error code if condition not exist

   m_csCondRefs.Lock();

   AeCondition* pCond = LookupCondition( szConditionName );
   if (pCond) {
                                                // condition found
         hres = pCond->GetState( dwNumEventAttrs, pdwAttributeIDs,
                                 ppConditionState );
   }

   m_csCondRefs.Unlock();
   return hres;
}



//=========================================================================
// AckCondition
// ------------
//    Acknowledges the condition specified by name.
//=========================================================================
HRESULT AeSource::AckCondition(
   /* [string][in] */            LPWSTR         szAcknowledgerID,
   /* [string][in] */            LPWSTR         szComment,
   /* [string][in] */            LPWSTR         szConditionName,
   /* [in] */                    FILETIME       ftActiveTime,
   /* [out] */                   AeEvent**     ppEvent )
{
   HRESULT hres = E_INVALIDARG;                 // Error code if condition not exist

   m_csCondRefs.Lock();

   AeCondition* pCond = LookupCondition( szConditionName );
   if (pCond) {                                 // Condition found
      hres = pCond->Acknowledge( szAcknowledgerID, szComment, ftActiveTime, ppEvent );
   }

   m_csCondRefs.Unlock();
   return hres;
}



//=========================================================================
// EnableConditions
// ----------------
//    Enables or disables all conditions which belongs to this source.
//
// Parameters:
//    IN
//       fEnable              if TRUE the conditions are enabled;
//                            otherwise disabled.
//    OUT
//       parEvents            Array with created event instances for
//                            each condition with changed enable state.
//
// Return Code:
//    S_OK                    All succeeded
//    S_FALSE                 Not succeeded for all conditions.
//    E_XXX                   Error occured.
//=========================================================================
HRESULT AeSource::EnableConditions(
   /* [in] */                    BOOL           fEnable,
   /* [out] */                   AeEventArray* parEvents )
{
   AeEvent*   pEvent;
   HRESULT     hres = S_OK;
   HRESULT     hresRet = S_OK;
   DWORD       dwSize, i;

   m_csCondRefs.Lock();
   dwSize = m_arCondRefs.GetSize();

   hresRet = parEvents->Create( dwSize );
   if (FAILED( hresRet )) {
      m_csCondRefs.Unlock();
      return hresRet;
   }
                                                // Handle all conditions which belongs to this source.
   for (i=0; i < dwSize; i++) {
      hres = m_arCondRefs[i]->Enable( fEnable, &pEvent );
      if (hres == S_OK) {                       // If hres is S_FALSE then the condition is already in the in
         _ASSERTE( pEvent );                    // specified state and pEvent is NULL.
         hres = parEvents->Add( pEvent );
      }
      if (FAILED( hres )) {
         hresRet = S_FALSE;
      }
   }
   m_csCondRefs.Unlock();
   return hresRet;
}



//=========================================================================
// GetAreaNames
// ------------
//    Creates a variant with all names of the areas to which this
//    source belongs (without root area).
//=========================================================================
HRESULT AeSource::GetAreaNames( LPVARIANT pvAreas )
{
   HRESULT hres = S_OK;
   long    i;

   VariantInit( pvAreas );
   V_VT( pvAreas ) = VT_ARRAY | VT_BSTR;

   m_csAreaRefs.Lock();

   long lElem = m_arAreaRefs.GetSize();
   for (i=0; i < lElem; i++) {                  // Checks if the source belongs to the root area.
      if (m_arAreaRefs[i]->IsRoot()) {          // Do not use the name of the root area
         lElem--;                               // beacuse the root instance is used only internally.
         break;
      }
   }
                                                // Create a safe array of BSTRs
   if (!(V_ARRAY( pvAreas ) = SafeArrayCreateVector( VT_BSTR, 0, lElem ))) {
      hres = E_OUTOFMEMORY;
   }
   else {
      BSTR        bstrArea;
      EventArea* pArea;
      long ndx = 0;
                                                // Fill the array of BSTRs with the area names
      lElem = m_arAreaRefs.GetSize();
      for (i=0; i < lElem; i++) {
         pArea = m_arAreaRefs[i];
         if (!pArea->IsRoot()) {                // Do not add the name of the root area to the list
                                                // beacuse the root instance is used only internally.
            bstrArea = pArea->Name().CopyBSTR();
            if (bstrArea) {
               hres = SafeArrayPutElement( V_ARRAY( pvAreas ), &ndx, bstrArea );
               SysFreeString( bstrArea );
            }
            else {
               hres = E_OUTOFMEMORY;
            }
            if (FAILED( hres )) {
               break;
            }
            ndx++;
         }
      }
   }
   m_csAreaRefs.Unlock();

   if (FAILED( hres )) {
      VariantClear( pvAreas );
   }

   return hres;
}


//-------------------------------------------------------------------------
// IMPLEMENTATION
//-------------------------------------------------------------------------

//=========================================================================
// LookupCondition
// ---------------
//    Gets the condition with the specified condition name.
//=========================================================================
AeCondition* AeSource::LookupCondition( LPCWSTR szName )
{
   AeCondition* pCond = NULL;
   m_csCondRefs.Lock();

   for (int i=0; i < m_arCondRefs.GetSize(); i++) {
      if (wcscmp( m_arCondRefs[i]->Name(), szName ) == 0) {
         pCond = m_arCondRefs[i];
         break;
      }
   }
   m_csCondRefs.Unlock();
   return pCond;
}
//DOM-IGNORE-END


#endif