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

//-----------------------------------------------------------------------
// INLCUDE
//-----------------------------------------------------------------------
#include "stdafx.h"
#include "DaComBaseServer.h"
#include "UtilityFuncs.h"
#include "FixOutArray.h"
#include "Logger.h"

//-----------------------------------------------------------------------
// CODE
//-----------------------------------------------------------------------

///////////////////////////////////////////////////////////////////////////
//////////////////////////// IOPCItemProperties /////////////////////////// 
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCItemProperties::QueryAvailableProperties                   INTERFACE
// --------------------------------------------
//    Returns a list of all available item properties for
//    the specified ItemID.
//=========================================================================
STDMETHODIMP DaComBaseServer::QueryAvailableProperties(
   /* [in] */                             LPWSTR      szItemID,
   /* [out] */                            DWORD    *  pdwCount,
   /* [size_is][size_is][out] */          DWORD    ** ppPropertyIDs,
   /* [size_is][size_is][out] */          LPWSTR   ** ppDescriptions,
   /* [size_is][size_is][out] */          VARTYPE  ** ppvtDataTypes )
{
	USES_CONVERSION;
	LOGFMTI( "IOPCItemProperties::QueryAvailableProperties for item %s", W2A(szItemID) );

   *pdwCount         = 0;                       // Note : Proxy/Stub checks if the pointers are NULL
   *ppPropertyIDs    = NULL;
   *ppDescriptions   = NULL;
   *ppvtDataTypes    = NULL;

   HRESULT  hresRet = S_OK;                     // Returned result
   LPDWORD  pdwPropIDs = NULL;
   LPVOID   pCookie = NULL;

   CFixOutArray< DWORD >   aID;
   CFixOutArray< VARTYPE > aVT;
   CFixOutArray< LPWSTR >  aDescr;

   try {
      HRESULT        hres;                      // Temporary results
      DaItemProperty* pProp;
      DWORD          dwNumOfProperties = 0;
                                                // Get the Property IDs available for the specified item
      hres = daBaseServer_->OnQueryItemProperties(
                                          szItemID,
                                          &dwNumOfProperties,
                                          &pdwPropIDs,
                                          &pCookie );
      if (FAILED( hres )) throw hres;
      if (dwNumOfProperties == 0) throw E_FAIL; // There are no item properties available

                                                // Allocate and initialize the result arrays.
      aID.Init( dwNumOfProperties, ppPropertyIDs );    
      aVT.Init( dwNumOfProperties, ppvtDataTypes );
      aDescr.Init( dwNumOfProperties, ppDescriptions );

      for (DWORD i=0; i<dwNumOfProperties; i++) {

         hres = daBaseServer_->GetItemProperty( pdwPropIDs[i], &pProp );
         if (FAILED( hres )) {                  // Unspecified Property
            _ASSERTE( 0 );                      // Programming error. Property Definition not
            throw E_FAIL;                       // added to the ServerClassHandler.
         }

         aID[i] = pProp->dwPropID;
         aVT[i] = pProp->vtDataType;
         aDescr[i] = WSTRClone( pProp->pwszDescr, pIMalloc );
         if (aDescr[i] == NULL) {
            throw E_OUTOFMEMORY;
         }
      }
      hres = daBaseServer_->OnReleasePropertyCookie( pCookie );
      pCookie = NULL;                           // Mark as released
      if (FAILED( hres )) throw hres;

      *pdwCount = dwNumOfProperties;            // All succeeded
   }
   catch (HRESULT hresEx) {
      if (pCookie) {                            // Release cookie if not yet done
         daBaseServer_->OnReleasePropertyCookie( pCookie );
      }
      aID.Cleanup();
      aVT.Cleanup();
      aDescr.Cleanup();
      hresRet = hresEx;
   }
   if (pdwPropIDs) {
      delete [] pdwPropIDs;                     // Release temporary ID array
   }
   return hresRet;
}
                                                      


//=========================================================================
// IOPCItemProperties::GetItemProperties                          INTERFACE
// -------------------------------------
//    Returns the data value for the passed item property IDs.
//=========================================================================
STDMETHODIMP DaComBaseServer::GetItemProperties(
   /* [in] */                             LPWSTR      szItemID,
   /* [in] */                             DWORD       dwCount,
   /* [size_is][in] */                    DWORD    *  pdwPropertyIDs,
   /* [size_is][size_is][out] */          VARIANT  ** ppvData,
   /* [size_is][size_is][out] */          HRESULT  ** ppErrors )
{                                                
	USES_CONVERSION;
	LOGFMTI( "IOPCItemProperties::GetItemProperties for item %s", W2A(szItemID) );
                                                
   *ppvData = NULL;                             // Note : Proxy/Stub checks if the pointers are NULL
   *ppErrors = NULL;

   if (dwCount == 0)
      return E_INVALIDARG;

   HRESULT  hresRet = S_OK;                     // Returned result
   LPDWORD  pdwPropIDs = NULL;
   LPVOID   pCookie = NULL;

   CFixOutArray< VARIANT > aVar;
   CFixOutArray< HRESULT > aErr;

   try {
      HRESULT  hres;                            // Temporary results
      DWORD    dwNumOfProperties = 0;
                                                // Get the Property IDs available for the specified item
      hres = daBaseServer_->OnQueryItemProperties(
                                          szItemID,
                                          &dwNumOfProperties,
                                          &pdwPropIDs,
                                          &pCookie );
      if (FAILED( hres )) throw hres;

      aVar.Init( dwCount, ppvData );            // Allocate and initialize the result arrays
      aErr.Init( dwCount, ppErrors,
                  OPC_E_INVALID_PID );          // Initialize with default error code

      DWORD             dwItemCnt;              // Counts the properties of the Item
      DWORD             dwReqCnt;               // Counts the properties requested by the user

                                                // Handle all requested propreties
      for (dwReqCnt=0; dwReqCnt<dwCount; dwReqCnt++) {

         for (dwItemCnt=0; dwItemCnt<dwNumOfProperties; dwItemCnt++) {
                                                // Checks if the item has the requested property
            if (pdwPropIDs[dwItemCnt] == pdwPropertyIDs[ dwReqCnt ]) {
                                                // The item has the requested property
                                                // Get the value of the current property
               aErr[ dwReqCnt ] = daBaseServer_->OnGetItemProperty(
                                          szItemID,
                                          pdwPropertyIDs[ dwReqCnt ],
                                          &aVar[ dwReqCnt ],
                                          pCookie );
               break;                           // Handle next property
            }
         }
         if (FAILED( aErr[ dwReqCnt ] )) {
            hresRet = S_FALSE;                  // S_FALSE beacuse not all results in ppErrors are S_OK
         }
      }
      hres = daBaseServer_->OnReleasePropertyCookie( pCookie );
      pCookie = NULL;                           // Mark as released
      if (FAILED( hres )) throw hres;
   }
   catch (HRESULT hresEx) {
      if (pCookie) {                            // Release cookie if not yet done
         daBaseServer_->OnReleasePropertyCookie( pCookie );
      }
      aVar.Cleanup();
      aErr.Cleanup();
      hresRet = hresEx;
   }
   if (pdwPropIDs) {
      delete [] pdwPropIDs;                     // Release temporary ID array
   }
   return hresRet;
}
                                                      


//=========================================================================
// IOPCItemProperties::LookupItemIDs                              INTERFACE
// ---------------------------------
//    Returs a list of ItemIDs for each of the passed properties
//    (if available).
//=========================================================================
STDMETHODIMP DaComBaseServer::LookupItemIDs(
   /* [in] */                             LPWSTR      szItemID,
   /* [in] */                             DWORD       dwCount,
   /* [size_is][in] */                    DWORD    *  pdwPropertyIDs,
   /* [size_is][size_is][string][out] */  LPWSTR   ** ppszNewItemIDs,
   /* [size_is][size_is][out] */          HRESULT  ** ppErrors )
{
	USES_CONVERSION;
	LOGFMTI( "IOPCItemProperties::LookupItemIDs for item %s", W2A(szItemID) );

   *ppszNewItemIDs = NULL;                      // Note : Proxy/Stub checks if the pointers are NULL
   *ppErrors = NULL;

   if (dwCount == 0)
      return E_INVALIDARG;

   HRESULT  hresRet = S_OK;                     // Returned result
   LPDWORD  pdwPropIDs = NULL;
   LPVOID   pCookie = NULL;

   CFixOutArray< LPWSTR >  aNewID;
   CFixOutArray< HRESULT > aErr;

   try {
      HRESULT  hres;                            // Temporary results
      DWORD    dwNumOfProperties = 0;
                                                // Get the Property IDs available for the specified item
      hres = daBaseServer_->OnQueryItemProperties(
                                          szItemID,
                                          &dwNumOfProperties,
                                          &pdwPropIDs,
                                          &pCookie );
      if (FAILED( hres )) throw hres;

      aNewID.Init( dwCount, ppszNewItemIDs );   // Allocate and initialize the result arrays
      aErr.Init( dwCount, ppErrors,
                  OPC_E_INVALID_PID );          // Initialize with default error code

      DWORD             dwItemCnt;              // Counts the properties of the Item
      DWORD             dwReqCnt;               // Counts the properties requested by the user

      for (dwReqCnt=0; dwReqCnt<dwCount; dwReqCnt++) {

         for (dwItemCnt=0; dwItemCnt<dwNumOfProperties; dwItemCnt++) {
                                                // Checks if the item has the requested property
                                                // and is not an ID set 1 property
            if ((pdwPropIDs[dwItemCnt] == pdwPropertyIDs[ dwReqCnt ]) &&
                (pdwPropertyIDs[dwReqCnt] > OPC_PROPERTY_EU_INFO)) {
                                                // The item has the requested property and is not an ID Set1 property
               aErr[ dwReqCnt ] = daBaseServer_->OnLookupItemID(
                                          szItemID,
                                          pdwPropertyIDs[ dwReqCnt ],
                                          &aNewID[ dwReqCnt ],
                                          pCookie );
               break;                           // Handle next property
            }
         }
         if (FAILED( aErr[ dwReqCnt ] )) {
            hresRet = S_FALSE;                  // S_FALSE beacuse not all results in ppErrors are S_OK
         }
      }
      hres = daBaseServer_->OnReleasePropertyCookie( pCookie );
      pCookie = NULL;                           // Mark as released
      if (FAILED( hres )) throw hres;
   }
   catch (HRESULT hresEx) {
      if (pCookie) {                            // Release cookie if not yet done
         daBaseServer_->OnReleasePropertyCookie( pCookie );
      }
      aNewID.Cleanup();
      aErr.Cleanup();
      hresRet = hresEx;
   }
   if (pdwPropIDs) {
      delete [] pdwPropIDs;                     // Release temporary ID array
   }
   return hresRet;
}
//DOM-IGNORE-END
