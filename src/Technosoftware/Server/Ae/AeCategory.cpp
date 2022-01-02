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
//-------------------------------------------------------------------------
// INLCUDE
//-------------------------------------------------------------------------
#include "stdafx.h"
#include "FixOutArray.h"
#include "AeAttribute.h"
#include "AeConditionDefinition.h"
#include "AeCategory.h"

//-------------------------------------------------------------------------
// CODE
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AeCategory::AeCategory()
{
   m_dwID         = 0;
   m_dwEventType  = 0;
   m_dwNumOfInternalAttrs = 0;
}



//=========================================================================
// Initializer
// -----------
//    Must be called after construction.
//    Also the internal default attributes 'ACK COMMENT' (for condition
//    event types only) and 'AREAS' (for all event types) are added.
//=========================================================================
HRESULT AeCategory::Create( DWORD dwID, LPCWSTR szDescr, DWORD dwEventType )
{
   HRESULT           hres;
   AeAttribute*  pAttr = NULL;

   m_csAttrMap.Lock();
   try {
      // Initialize class members
      m_dwID         = dwID;
      m_dwEventType  = dwEventType;
      hres = m_wsDescr.SetString( szDescr );
      if (FAILED( hres )) throw hres;

      // Add internal default attributes

      if (dwEventType == OPC_CONDITION_EVENT) { // ACK COMMENT only for categories related to type 'Condition Event'
         pAttr = new AeAttribute;              
         if (!pAttr) throw E_OUTOFMEMORY;

         hres = pAttr->Create( ATTRID_ACKCOMMENT, L"Ack Comment", VT_BSTR );
         if (FAILED( hres )) throw hres;
   
         hres = AddEventAttribute( pAttr );     // Add the attribute to this category
         if (FAILED( hres )) throw hres;
         m_dwNumOfInternalAttrs++;
      }
                                                
      pAttr = new AeAttribute;              // AREAS
      if (!pAttr) throw E_OUTOFMEMORY;

      hres = pAttr->Create( ATTRID_AREAS, L"Areas", VT_ARRAY | VT_BSTR );
      if (FAILED( hres )) throw hres;
   
      hres = AddEventAttribute( pAttr );        // Add the attribute to this category
      if (FAILED( hres )) throw hres;
      m_dwNumOfInternalAttrs++;
   }
   catch (HRESULT hresEx) {
      if (pAttr)
         delete pAttr;
      hres = hresEx;
   }
   m_csAttrMap.Unlock();
   return hres;
}



//=========================================================================
// Destructor
//=========================================================================
AeCategory::~AeCategory()
{
   int i;

   // Cleanup Attribute Array
   AeAttribute* pAttr;
   m_csAttrMap.Lock();
   for (i=0; i<m_mapAttributes.GetSize(); i++) {
      pAttr = m_mapAttributes.m_aVal[i];
      if (pAttr) delete pAttr;
   }
   m_mapAttributes.RemoveAll();
   m_csAttrMap.Unlock();
}


//-------------------------------------------------------------------------
// OPERATIONS
//-------------------------------------------------------------------------

//=========================================================================
// AddEventAttribute
// -----------------
//    Adds a vendor specific attribute to the category.
//=========================================================================
HRESULT AeCategory::AddEventAttribute( AeAttribute* pAttr )
{
                                                // Add the new attribute to the list
   m_csAttrMap.Lock();
   HRESULT hres = m_mapAttributes.Add( pAttr->AttrID(), pAttr ) ? S_OK : E_OUTOFMEMORY;
   m_csAttrMap.Unlock();

   return hres;
}



//=========================================================================
// ExistEventAttribute
// -------------------
//    Checks if there is an event attribute with the specified identifier.
//=========================================================================
BOOL AeCategory::ExistEventAttribute( DWORD dwAttrID )
{
   m_csAttrMap.Lock();
   BOOL fExist = m_mapAttributes.Lookup( dwAttrID ) ? TRUE : FALSE;
   m_csAttrMap.Unlock();

   return fExist;
}



//=========================================================================
// ExistEventAttribute
// -------------------
//    Checks if there is an event attribute with the specified description.
//=========================================================================
BOOL AeCategory::ExistEventAttribute( LPCWSTR szAttrDescr )
{
   BOOL  fExist = FALSE;

   m_csAttrMap.Lock();
   for (int i=0; i<m_mapAttributes.GetSize(); i++) {
      if (wcscmp( m_mapAttributes.m_aVal[i]->Descr(), szAttrDescr ) == 0) {
         fExist = TRUE;
         break;
      }
   }
   m_csAttrMap.Unlock();

   return fExist;
}



//=========================================================================
// QueryEventAttributes
// --------------------
//    Queries all vendor specific attributes attached to this category.
//=========================================================================
HRESULT AeCategory::QueryEventAttributes(
                                    DWORD    *  pdwCount,
                                    DWORD    ** ppdwAttrIDs,
                                    LPWSTR   ** ppszAttrDescs,
                                    VARTYPE  ** ppvtAttrTypes )
{
   HRESULT           hres = S_OK;

   CFixOutArray< DWORD >   aID;
   CFixOutArray< VARTYPE > aVT;
   CFixOutArray< LPWSTR >  aDescr;

   m_csAttrMap.Lock();
   *pdwCount = m_mapAttributes.GetSize();

   try {
      aID.Init( *pdwCount, ppdwAttrIDs );       // Allocate and initialize the result arrays
      aVT.Init( *pdwCount, ppvtAttrTypes );
      aDescr.Init( *pdwCount, ppszAttrDescs );

      AeAttribute* pAttr;

      for (DWORD i=0; i < *pdwCount; i++) {     // Fill the result arrays with the
         pAttr = m_mapAttributes.m_aVal[i];     // values of all event attribute 
         aID[i]      = pAttr->AttrID();
         aVT[i]      = pAttr->VarType();
         aDescr[i]   = pAttr->Descr().CopyCOM();
         if (aDescr[i] == NULL) throw E_OUTOFMEMORY;
      }
   }
   catch (HRESULT hresEx) {
      *pdwCount = 0;
      aID.Cleanup();
      aVT.Cleanup();
      aDescr.Cleanup();
      hres = hresEx;
   }
   m_csAttrMap.Unlock();

   return hres;
}



//=========================================================================
// GetAttributeIDs
// ---------------
//    Returns an array with the IDs of all attributes of this category.
//    The caller of the function must free the array.
//=========================================================================
HRESULT AeCategory::GetAttributeIDs( DWORD* pdwCount, DWORD** ppdwAttrIDs )
{
   DWORD             i;
   HRESULT           hres = S_OK;

   m_csAttrMap.Lock();

   *pdwCount = m_mapAttributes.GetSize();
   *ppdwAttrIDs = new DWORD [*pdwCount];
   if (*ppdwAttrIDs == NULL) {
      m_csAttrMap.Unlock();
      *pdwCount = 0;
      return E_OUTOFMEMORY;
   }

   for (i=0; i < *pdwCount; i++) {              // Fill the result arrays with the IDs
      (*ppdwAttrIDs)[i] = m_mapAttributes.m_aKey[i];
   }

   m_csAttrMap.Unlock();

   return hres;
}



//=========================================================================
// QueryConditionNames
// -------------------
//    Gets the names of all condition definitions attached to this
//    category.
//=========================================================================
HRESULT AeCategory::QueryConditionNames(
                                    DWORD    *  pdwCount,
                                    LPWSTR   ** ppszConditionNames )
{
   HRESULT                 hres = S_OK;
   CFixOutArray< LPWSTR >  aName;

   m_csCondDefRefs.Lock();
   *pdwCount = m_arCondDefRefs.GetSize();

   try {                                        // Allocate and initialize the result arrays
      aName.Init( *pdwCount, ppszConditionNames );
                                                // Fill the result arrays with the
      for (DWORD i=0; i < *pdwCount; i++) {     // name of all condition definitions
         aName[i] = m_arCondDefRefs[i]->Name().CopyCOM();
         if (aName[i] == NULL) throw E_OUTOFMEMORY;
      }
   }
   catch (HRESULT hresEx) {
      *pdwCount = 0;
      aName.Cleanup();
      hres = hresEx;
   }
   m_csCondDefRefs.Unlock();

   return hres;
}



//=========================================================================
// AttachConditionDef
// ------------------
//    Attaches a link to a condition definition to this category.
//=========================================================================
HRESULT AeCategory::AttachConditionDef( AeConditionDefiniton* pCondDef )
{
   _ASSERTE( pCondDef != NULL );                // Must not be NULL

   HRESULT hres;
   m_csCondDefRefs.Lock();
   hres = m_arCondDefRefs.Add( pCondDef ) ? S_OK : E_OUTOFMEMORY;
   m_csCondDefRefs.Unlock();
   return hres;
}



//=========================================================================
// DetachConditionDef
// ------------------
//    Detaches a link to a condition definition from this category.
//=========================================================================
HRESULT AeCategory::DetachConditionDef( AeConditionDefiniton* pCondDef )
{
   _ASSERTE( pCondDef != NULL );                // Must not be NULL

   HRESULT hres;
   m_csCondDefRefs.Lock();
   hres = m_arCondDefRefs.Remove( pCondDef ) ? S_OK : E_FAIL;
   m_csCondDefRefs.Unlock();
   return hres;
}
//DOM-IGNORE-END






