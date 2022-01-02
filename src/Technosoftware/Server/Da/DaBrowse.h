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

#ifndef __BROWSESAS_H
#define __BROWSESAS_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000


/////////////////////////////////////////////////////////////////////////////
// Forward declarations
/////////////////////////////////////////////////////////////////////////////
class DaBaseServer;


/////////////////////////////////////////////////////////////////////////////
// Class declaration DaBrowseData
/////////////////////////////////////////////////////////////////////////////

typedef struct tagBROWSEDATA {

      // Actual path positions in the address space
   BSTR     m_bstrServerBrowsePosition;   // DA 2.0
   BSTR     m_bstrServerBrowsePosition3;  // DA 3.0

      // Holds the custom specific client related data.
   LPVOID   m_pCustomData;

} BROWSEDATA, *PBROWSEDATA;


class DaBrowseData : public tagBROWSEDATA
{
// Construction
public:
   DaBrowseData();
   HRESULT Create( DaBaseServer* pServerHandler );

// Destruction
   ~DaBrowseData();

// Implementation
protected:
   void Cleanup();
   BOOL                 m_fReleaseCustomData;
   DaBaseServer* m_pServerHandler;       // Used only to have access to
                                                // OnCreateCustomServerData() and
                                                // OnDestroyCustomServerData().
};


/////////////////////////////////////////////////////////////////////////////
// Class declaration DaBrowse
/////////////////////////////////////////////////////////////////////////////
class DaBrowse : public IOPCBrowseServerAddressSpace,
                   public IOPCBrowse
{
// Construction
public:
   DaBrowse::DaBrowse()
   {
      m_pBrowseData     = NULL;
      m_pServerHandler  = NULL;
   }

   HRESULT Create( PBROWSEDATA pBrowseData, DaBaseServer* pServerHandler )
   {
      m_pBrowseData     = pBrowseData;
      m_pServerHandler  = pServerHandler;
      return S_OK;
   }

// Destruction
   DaBrowse::~DaBrowse()
   {
   }

// Operations
public:

   //=========================================================================
   // OPC 2.0 Custom Interfaces
   //=========================================================================

   ///////////////////////////////////////////////////////////////////////////
   /////////////////////// IOPCBrowseServerAddressSpace //////////////////////
   ///////////////////////////////////////////////////////////////////////////

   STDMETHODIMP QueryOrganization(
                  /* [out] */                   OPCNAMESPACETYPE  *  pNameSpaceType
                  );

   STDMETHODIMP ChangeBrowsePosition(
                  /* [in] */                    OPCBROWSEDIRECTION   dwBrowseDirection,
                  /* [in, string] */            LPCWSTR              szString
                  );

   STDMETHODIMP BrowseOPCItemIDs(
                  /* [in] */                    OPCBROWSETYPE        dwBrowseFilterType,
                  /* [in, string] */            LPCWSTR              szFilterCriteria,
                  /* [in] */                    VARTYPE              vtDataTypeFilter,
                  /* [in] */                    DWORD                dwAccessRightsFilter,
                  /* [out] */                   LPENUMSTRING      *  ppIEnumString
                  );

   STDMETHODIMP GetItemID(
                  /* [in] */                    LPWSTR               szItemDataID,
                  /* [out, string] */           LPWSTR            *  szItemID
                  );

   STDMETHODIMP BrowseAccessPaths(
                  /* [in, string] */            LPCWSTR              szItemID,
                  /* [out] */                   LPENUMSTRING      *  ppIEnumString
                  );


   //=========================================================================
   // OPC 3.0 Custom Interfaces
   //=========================================================================
                                                                           
   ///////////////////////////////////////////////////////////////////////////
   //////////////////////////// IOPCBrowse ///////////////////////////////////
   ///////////////////////////////////////////////////////////////////////////
   
   STDMETHODIMP GetProperties(
                  /* [in] */                    DWORD                dwItemCount,
                  /* [size_is][string][in] */   LPWSTR            *  pszItemIDs,
                  /* [in] */                    BOOL                 bReturnPropertyValues,
                  /* [in] */                    DWORD                dwPropertyCount,
                  /* [size_is][in] */           DWORD             *  pdwPropertyIDs,
                  /* [size_is][size_is][out] */ OPCITEMPROPERTIES ** ppItemProperties
                  );
        
   STDMETHODIMP Browse(
                  /* [string][in] */            LPWSTR               szItemID,
                  /* [string][out][in] */       LPWSTR            *  pszContinuationPoint,
                  /* [in] */                    DWORD                dwMaxElementsReturned,
                  /* [in] */                    OPCBROWSEFILTER      dwBrowseFilter,
                  /* [string][in] */            LPWSTR               szElementNameFilter,
                  /* [string][in] */            LPWSTR               szVendorFilter,
                  /* [in] */                    BOOL                 bReturnAllProperties,
                  /* [in] */                    BOOL                 bReturnPropertyValues,
                  /* [in] */                    DWORD                dwPropertyCount,
                  /* [size_is][in] */           DWORD             *  pdwPropertyIDs,
                  /* [out] */                   BOOL              *  pbMoreElements,
                  /* [out] */                   DWORD             *  pdwCount,
                  /* [size_is][size_is][out] */ OPCBROWSEELEMENT  ** ppBrowseElements
                  );


   ///////////////////////////////////////////////////////////////////////////

// Impmementation
protected:
   PBROWSEDATA          m_pBrowseData;
   DaBaseServer* m_pServerHandler;

   HRESULT GetRevisedPropertyIDs(
                  /* [in] */                    const LPWSTR         szItemID,
                  /* [in] */                    const DWORD          dwPropertyCount,
                  /* [size_is][in] */           const DWORD       *  pdwPropertyIDs,
                  /* [out] */                   DWORD             *  pdwRevisedPropertyCount,
                  /* [size_is][size_is][out] */ DWORD             ** ppdwRevisedPropertyIDs,
                  /* [size_is][size_is][out] */ HRESULT           ** pphrErrorID,
                  /* [out] */                   LPVOID            *  ppCookie );



   HRESULT MoveElementIDsToBrowseElements(
                  // Source
                  /* [in] */                    const DWORD          dwElementType,
                  /* [in] */                    const DWORD          dwNumOfElementIDs,
                  /* [in(dwNumOfElementIDs)] */ BSTR              *  pElementIDs,
                  // Destination
                  /* [in] */                    const DWORD          dwNumOfElements,
                  /* [in,out] */                DWORD             *  pdwElementCount,
                  /* [in(dwNumOfElements)] */   OPCBROWSEELEMENT  *  pElements,
                  /* [in,out] */                LPWSTR            *  pszContinuationPoint,
                  // Property Related Parameters
                  /* [in] */                    const BOOL           bReturnAllProperties,
                  /* [in] */                    const BOOL           bReturnPropertyValues,
                  /* [in] */                    const DWORD          dwPropertyCount,
                  /* [in(dwPropertyCount)] */   DWORD             *  pdwPropertyIDs );

   // Small Utility Functions
   DWORD FilterElements( LPCWSTR szElementNameFilter, DWORD dwNumOfElements, BSTR* pElements );
   DWORD RemoveElementsInFrontOfContinuationPoint( LPCWSTR szContinuationPoint, DWORD dwNumOfElements, BSTR* pElements );
   void  InitArrayOfOPCITEMPROPERTY( DWORD dwNumOfProp, OPCITEMPROPERTY* pProperties );
   void  ReleaseArrayOfOPCITEMPROPERTY( DWORD dwNumOfProp, OPCITEMPROPERTY* pProperties );
   void  ReleaseOPCITEMPROPERTY( OPCITEMPROPERTY* pProperty );
};
//DOM-IGNORE-END

#endif // __BROWSESAS_H
