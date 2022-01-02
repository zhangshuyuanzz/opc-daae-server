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
#include "enumclass.h"
#include "AeEventArea.h"
#include "AeAreaBrowser.h"

//-------------------------------------------------------------------------
// CODE
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
EventAreaBrowser::EventAreaBrowser()
{
   m_pCurrentArea = NULL;
}



//=========================================================================
// Initializer
// -----------
//    Must be called after construction.
//=========================================================================
HRESULT EventAreaBrowser::Create( EventArea* pRootArea )
{
   if (!pRootArea)
      return E_FAIL;

   m_pCurrentArea = pRootArea;
   return S_OK;
}



//=========================================================================
// Destructor
//=========================================================================
EventAreaBrowser::~EventAreaBrowser()
{
}



//-------------------------------------------------------------------------
// OPERATIONS
//-------------------------------------------------------------------------

//=========================================================================
// IOPCEventAreaBrowser::ChangeBrowsePosition                     INTERFACE
// ------------------------------------------
//    Changes the current position in the hierarchical event area space.
//=========================================================================
STDMETHODIMP EventAreaBrowser::ChangeBrowsePosition(
   /* [in] */                    OPCAEBROWSEDIRECTION
                                                   dwBrowseDirection,
   /* [string][in] */            LPCWSTR           szString )
{
   EventArea* pNewArea;
   HRESULT     hres = S_OK;

   switch (dwBrowseDirection) {
      case OPCAE_BROWSE_UP :
      //-------------------------------------------------------------------
         if (m_pCurrentArea == m_pCurrentArea->Root()) {
            // 22-jan-2002 MT 
            // Return code changed from E_INVALIDARG to E_FAIL
            // (see spec OPC AE 1.03).
            hres = E_FAIL;                // Already at root position
         }
         else {
            m_pCurrentArea = m_pCurrentArea->Parent();
         }
         return hres;

      case OPCAE_BROWSE_DOWN :
      //-------------------------------------------------------------------
         hres = m_pCurrentArea->PositionDown( szString, &pNewArea );
         if (SUCCEEDED( hres )) {
            m_pCurrentArea = pNewArea;
         }
         return hres;

      case OPCAE_BROWSE_TO :
      //-------------------------------------------------------------------
                                                // Move to the specified position
         hres = m_pCurrentArea->PositionTo( szString, &pNewArea );
         if (SUCCEEDED( hres )) {
            m_pCurrentArea = pNewArea;
         }
         return hres;

      default:
         return E_INVALIDARG;
   }
   return E_INVALIDARG;                         // Bad direction
}

        

//=========================================================================
// IOPCEventAreaBrowser::BrowseOPCAreas                           INTERFACE
// ------------------------------------
//    Returns an enumeration with the names of areas/sources at
//    the current position.
//=========================================================================
STDMETHODIMP EventAreaBrowser::BrowseOPCAreas(
   /* [in] */                    OPCAEBROWSETYPE   dwBrowseFilterType,
   /* [string][in] */            LPCWSTR           szFilterCriteria,
   /* [out] */                   LPENUMSTRING   *  ppIEnumString )
{
   HRESULT              hres = S_OK;
   COpcComEnumString*   pEnum = NULL;

   try {

      if ((dwBrowseFilterType != OPC_AREA) &&
          (dwBrowseFilterType != OPC_SOURCE)) {
         throw E_INVALIDARG;
      }

      pEnum = new COpcComEnumString;
      if (pEnum == NULL) throw E_OUTOFMEMORY;

      hres = m_pCurrentArea->BrowseAreas( this, pEnum, dwBrowseFilterType, szFilterCriteria );
      if (FAILED( hres )) throw hres;
                                                // Do not overwrite browse result because
                                                // it's S_FALSE if there is nothing to enumerate.
      HRESULT hresQI = pEnum->_InternalQueryInterface( IID_IEnumString, (LPVOID *)ppIEnumString );
      if (FAILED( hresQI )) throw hresQI;
   }
   catch (HRESULT hresEx) {
      if (pEnum) {
         delete pEnum;
      }
      hres = hresEx;
      *ppIEnumString = NULL;
   }
   return hres;
}

        

//=========================================================================
// IOPCEventAreaBrowser::GetQualifiedAreaName                     INTERFACE
// ------------------------------------------
//    Gets the fully qualified area name in the hierarchical event
//    area space.
//=========================================================================
STDMETHODIMP EventAreaBrowser::GetQualifiedAreaName(
   /* [in] */                    LPCWSTR           szAreaName,
   /* [string][out] */           LPWSTR         *  pszQualifiedAreaName )
{
   *pszQualifiedAreaName = NULL;

   return m_pCurrentArea->GetQualifiedAreaName( szAreaName, pszQualifiedAreaName );
}

        

//=========================================================================
// IOPCEventAreaBrowser::GetQualifiedSourceName                   INTERFACE
// --------------------------------------------
//    Gets the fully qualified source name in the hierarchical event
//    area space.
//=========================================================================
STDMETHODIMP EventAreaBrowser::GetQualifiedSourceName(
   /* [in] */                    LPCWSTR           szSourceName,
   /* [string][out] */           LPWSTR         *  pszQualifiedSourceName )
{
   *pszQualifiedSourceName = NULL;

   return m_pCurrentArea->GetQualifiedSourceName( szSourceName, pszQualifiedSourceName );
}
//DOM-IGNORE-END
#endif