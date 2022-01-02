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
#include "UtilityFuncs.h"
#include "UtilityDefs.h"
#include "AeEvent.h"
#include "AeCategory.h"
#include "AeConditionDefinition.h"

//-------------------------------------------------------------------------
// CODE AeSubConditionDefiniton
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AeSubConditionDefiniton::AeSubConditionDefiniton()
{
   m_dwSeverity = 0;
   m_fAckRequired = FALSE;
}



//=========================================================================
// Initializer
// -----------
//    Must be called after construction.
//=========================================================================
HRESULT AeSubConditionDefiniton::Create( LPCWSTR szName, LPCWSTR szDef,
                                       DWORD dwSeverity, LPCWSTR szDescr,
                                       BOOL fAckRequired )
{
   HRESULT hres = S_OK;

   try {
      m_dwSeverity = dwSeverity;
      m_fAckRequired = fAckRequired;

      hres = m_wsName.SetString( szName );
      if (FAILED( hres ))  throw hres;

      hres = m_wsDescr.SetString( szDescr );
      if (FAILED( hres ))  throw hres;

      hres = m_wsDef.SetString( szDef );
      if (FAILED( hres ))  throw hres;
   }
   catch(HRESULT hresEx) {
      hres = hresEx;
   }
   return hres;
}



//=========================================================================
// Destructor
//=========================================================================
AeSubConditionDefiniton::~AeSubConditionDefiniton()
{
}



//-------------------------------------------------------------------------
// CODE AeConditionDefiniton
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AeConditionDefiniton::AeConditionDefiniton()
{
   m_pCategory = NULL;
   m_dwCondDefID  = 0;
}



//=========================================================================
// Initializer Single State Condition
// ----------------------------------
//    The condition is initialized as single state condition.
//=========================================================================
HRESULT AeConditionDefiniton::Create( AeCategory* pCategory, DWORD dwCondDefID,
                                    LPCWSTR szName, LPCWSTR szDef,
                                    DWORD dwSeverity, LPCWSTR szDescr,
                                    BOOL fAckRequired )
{
   m_csMem.Lock();

   HRESULT hres = Create( pCategory, dwCondDefID, szName );
   if (SUCCEEDED( hres )) {
      hres = AddSubCondDef( 0, szName, szDef, dwSeverity, szDescr, fAckRequired );
   }

   m_csMem.Unlock();
   return hres;
}



//=========================================================================
// Initializer Multi State Condition
// ---------------------------------
//    The condition is initialized as multi state condition.
//=========================================================================
HRESULT AeConditionDefiniton::Create( AeCategory* pCategory, DWORD dwCondDefID, LPCWSTR szName )
{
   HRESULT hres;

   m_csMem.Lock();

   m_pCategory = pCategory;
   m_dwCondDefID = dwCondDefID;

   hres = m_wsName.SetString( szName );
   if (FAILED( hres ))
      return hres;

   m_csMem.Unlock();
   return hres;
}



//=========================================================================
// AddSubCondDef
// -------------
//    Adds a sub condition if initialized as multi-state condition object.
//=========================================================================
HRESULT AeConditionDefiniton::AddSubCondDef( DWORD dwSubCondDefID, LPCWSTR szName, LPCWSTR szDef, DWORD dwSeverity, LPCWSTR szDescr, BOOL fAckRequired )
{
   HRESULT                 hres;
   AeSubConditionDefiniton*  pSubCond = NULL;
   m_csMem.Lock();
   try {
                                                // ID must be unique
      if (m_mapSubCond.Lookup( dwSubCondDefID ))
         throw E_OBJECT_IN_LIST;

      DWORD dwDummyID;                          // Name must be unique
      hres = GetSubCondDefID( szName, &dwDummyID );
      if (SUCCEEDED( hres )) throw E_OBJECT_IN_LIST;

      pSubCond = new AeSubConditionDefiniton;
      if (!pSubCond) throw E_OUTOFMEMORY;
      
      hres = pSubCond->Create( szName, szDef, dwSeverity, szDescr, fAckRequired );
      if (FAILED( hres )) throw hres;

      hres = m_mapSubCond.Add( dwSubCondDefID, pSubCond ) ? S_OK : E_OUTOFMEMORY;
      if (FAILED( hres )) throw hres;
   }
   catch(HRESULT hresEx) {
      if (pSubCond)
         delete pSubCond;
      hres = hresEx;
   }
   m_csMem.Unlock();
   return hres;
}



//=========================================================================
// Destructor
//=========================================================================
AeConditionDefiniton::~AeConditionDefiniton()
{
   // Cleanup Sub Condition Map
   AeSubConditionDefiniton* pSubCond;
   m_csMem.Lock();
   for (int i=0; i < m_mapSubCond.GetSize(); i++) {
      pSubCond = m_mapSubCond.m_aVal[i];
      if (pSubCond) delete pSubCond;
   }
   m_csMem.Unlock();
}


//-------------------------------------------------------------------------
// OPERATIONS
//-------------------------------------------------------------------------

//=========================================================================
// GetSubConditionInfo
// -------------------
//    Returns information of all subconditions of this condition
//    definition. Memory is allocated from the Global COM Memory Manager.
//=========================================================================
HRESULT AeConditionDefiniton::GetSubConditionInfo(
                        LPDWORD pdwNumSCs,
                        LPWSTR** pszSCNames, LPWSTR** pszSCDefinitions,
                        LPDWORD* pdwSCSeverities, LPWSTR** pszSCDescriptions )
{
   DWORD    i;
   HRESULT  hres = S_OK;
   LPWSTR*  pNames  = NULL;
   LPWSTR*  pDefs   = NULL;
   LPDWORD  pSevs   = NULL;
   LPWSTR*  pDescrs = NULL;

   m_csMem.Lock();
   try {

      *pdwNumSCs = m_mapSubCond.GetSize();

      if (!(pNames = ComAlloc<LPWSTR>( *pdwNumSCs )))
         throw E_OUTOFMEMORY;
      memset( pNames, 0, sizeof (LPWSTR) * (*pdwNumSCs) );

      if (!(pDefs = ComAlloc<LPWSTR>( *pdwNumSCs )))
         throw E_OUTOFMEMORY;
      memset( pDefs, 0, sizeof (LPWSTR) * (*pdwNumSCs) );

      if (!(pDescrs = ComAlloc<LPWSTR>( *pdwNumSCs )))
         throw E_OUTOFMEMORY;
      memset( pDescrs, 0, sizeof (LPWSTR) * (*pdwNumSCs) );

      if (!(pSevs = ComAlloc<DWORD>( *pdwNumSCs )))
         throw E_OUTOFMEMORY;

      AeSubConditionDefiniton* pSubCondDef;
      for (i=0; i < *pdwNumSCs; i++) {
         pSubCondDef = m_mapSubCond.m_aVal[i];

         if (!(pNames[i] = pSubCondDef->Name().CopyCOM()))
            throw E_OUTOFMEMORY;
         if (!(pDefs[i] = pSubCondDef->Definition().CopyCOM()))
            throw E_OUTOFMEMORY;
         if (!(pDescrs[i] = pSubCondDef->Description().CopyCOM()))
            throw E_OUTOFMEMORY;
         pSevs[i] = pSubCondDef->Severity();
      }

   }
   catch(HRESULT hresEx) {

      for (i=0; i < *pdwNumSCs; i++) {
         if (pNames && pNames[i])
            WSTRFree( pNames[i], pIMalloc );
         if (pDefs && pDefs[i])
            WSTRFree( pDefs[i], pIMalloc );
         if (pDescrs && pDescrs[i])
            WSTRFree( pDescrs[i], pIMalloc );
      }
      if (pNames) {
         pIMalloc->Free( pNames );
         pNames = NULL;
      }
      if (pDefs) {
         pIMalloc->Free( pDefs );
         pDefs = NULL;
      }
      if (pDescrs) {
         pIMalloc->Free( pDescrs );
         pDescrs = NULL;
      }
      if (pSevs) {
         pIMalloc->Free( pSevs );
         pSevs = NULL;
      }
      *pdwNumSCs = 0;
      hres = hresEx;
   }
   m_csMem.Unlock();

   *pszSCNames          = pNames;
   *pszSCDefinitions    = pDefs;
   *pdwSCSeverities     = pSevs; 
   *pszSCDescriptions   = pDescrs;

   return hres;
}



//=========================================================================
// QuerySubConditionNames
// -----------------------
//    Gets the name of all sub conditions of this condition.
//    Memory is allocated from the Global COM Memory Manager.
//=========================================================================
HRESULT AeConditionDefiniton::QuerySubConditionNames(
   /* [out] */                   DWORD       *  pdwCount,
   /* [size_is][size_is][out] */ LPWSTR      ** ppszSubConditionNames )
{
   DWORD    i;
   HRESULT  hres = S_OK;
   LPWSTR*  pNames  = NULL;

   m_csMem.Lock();
   try {
      *pdwCount = m_mapSubCond.GetSize();

      if (!(pNames = ComAlloc<LPWSTR>( *pdwCount )))
         throw E_OUTOFMEMORY;
      memset( pNames, 0, sizeof (LPWSTR) * (*pdwCount) );

      for (i=0; i < *pdwCount; i++) {
         if (!(pNames[i] = m_mapSubCond.m_aVal[i]->Name().CopyCOM()))
            throw E_OUTOFMEMORY;
      }
   }
   catch(HRESULT hresEx) {
      if (pNames) {                             // Release all successfully copied sub condition names
         for (i=0; i < *pdwCount; i++) {
            if (pNames[i]) {
               WSTRFree( pNames[i], pIMalloc );
            }
         }
         pIMalloc->Free( pNames );
         pNames = NULL;
      }
      *pdwCount = 0;
      hres = hresEx;
   }
   m_csMem.Unlock();

   *ppszSubConditionNames = pNames;
   return hres;
}



//=========================================================================
// GetSubCondDefID
// ---------------
//    Returns the ID of the sub condition specified by name.
//=========================================================================
HRESULT AeConditionDefiniton::GetSubCondDefID(
   /* [in] */                    LPCWSTR szSubCondName,
   /* [out] */                   LPDWORD pdwSubCondDefID )
{
   HRESULT hres = E_INVALID_HANDLE;

   m_csMem.Lock();
   for (int i=0; i < m_mapSubCond.GetSize(); i++) {
      if (wcscmp( m_mapSubCond.m_aVal[i]->Name(), szSubCondName ) == 0) {
         *pdwSubCondDefID = m_mapSubCond.m_aKey[i];
         hres = S_OK;
         break;
      }
   }
   m_csMem.Unlock();
   return hres;
}



//=========================================================================
// GetSubCondDefID
// ---------------
//    Returns the ID of the sub condition specified by the object.
//=========================================================================
HRESULT AeConditionDefiniton::GetSubCondDefID(
   /* [in] */                    AeSubConditionDefiniton* pSubCond,
   /* [out] */                   LPDWORD pdwSubCondDefID )
{
   HRESULT hres = E_INVALID_HANDLE;

   m_csMem.Lock();
      // Don't use ReverseLookup() because 0 is a valid sub condition
      // id for single-state conditions.
   int nIndex = m_mapSubCond.FindVal( pSubCond );
   if (nIndex != -1)  {
      *pdwSubCondDefID = m_mapSubCond.GetKeyAt( nIndex );
      hres = S_OK;
   }
   m_csMem.Unlock();
   return hres;
}
//DOM-IGNORE-END