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
#include "WideString.h"
#include "MatchPattern.h"
#include "enumclass.h"
#include "DaBaseServer.h"
#include "DaBrowse.h"
#include "Logger.h"

//-------------------------------------------------------------------------
// CODE DaBrowseData
//-------------------------------------------------------------------------

//=========================================================================
// Constructor
//=========================================================================
DaBrowseData::DaBrowseData()
{
	m_bstrServerBrowsePosition = NULL;
	m_bstrServerBrowsePosition3 = NULL;
	m_pCustomData = NULL;
	m_pServerHandler = NULL;
	m_fReleaseCustomData = FALSE;
}



//=========================================================================
// Initializer
// -----------
//    Must be called after construction.
//=========================================================================
HRESULT DaBrowseData::Create(DaBaseServer* pServerHandler)
{
	_ASSERTE(pServerHandler);                  // Must not be NULL

	try {
		m_pServerHandler = pServerHandler;

		m_bstrServerBrowsePosition = SysAllocString(L"");
		_OPC_CHECK_PTR(m_bstrServerBrowsePosition);

		m_bstrServerBrowsePosition3 = SysAllocString(L"");
		_OPC_CHECK_PTR(m_bstrServerBrowsePosition3);

		HRESULT hr = m_pServerHandler->OnCreateCustomServerData(&m_pCustomData);
		_OPC_CHECK_HR(hr);

		m_fReleaseCustomData = TRUE;
		return S_OK;
	}
	catch (HRESULT hrEx) {
		Cleanup();
		return hrEx;
	}
	catch (...) {
		Cleanup();
		return E_FAIL;
	}
}



//=========================================================================
// Destructor
//=========================================================================
DaBrowseData::~DaBrowseData()
{
	Cleanup();
}



//=========================================================================
// Cleanup
// -------
//    Releases all resources
//=========================================================================
void DaBrowseData::Cleanup()
{
	SysFreeString(m_bstrServerBrowsePosition);
	m_bstrServerBrowsePosition = NULL;

	SysFreeString(m_bstrServerBrowsePosition3);
	m_bstrServerBrowsePosition3 = NULL;

	if (m_fReleaseCustomData) {
		m_pServerHandler->OnDestroyCustomServerData(&m_pCustomData);
		m_fReleaseCustomData = FALSE;
	}
}



//-------------------------------------------------------------------------
// CODE DaBrowse
//-------------------------------------------------------------------------

//=========================================================================
// OPC 2.0 Custom Interfaces
//=========================================================================

///////////////////////////////////////////////////////////////////////////
/////////////////////// IOPCBrowseServerAddressSpace //////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCBrowseServerAddressSpace::QueryOrganization                INTERFACE
//=========================================================================
STDMETHODIMP DaBrowse::QueryOrganization(
	/* [out] */                   OPCNAMESPACETYPE  *  pNameSpaceType)
{
	LOGFMTI("IOPCBrowseServerAddressSpace::QueryOrganization");

	if (!pNameSpaceType) {
		LOGFMTE("QueryOrganization() failed with invalid argument(s):   pNameSpaceType is NULL");
		return E_INVALIDARG;
	}
	*pNameSpaceType = m_pServerHandler->OnBrowseOrganization();
	return S_OK;
}



//=========================================================================
// IOPCBrowseServerAddressSpace::ChangeBrowsePosition             INTERFACE
//=========================================================================
STDMETHODIMP DaBrowse::ChangeBrowsePosition(
	/* [in] */                    OPCBROWSEDIRECTION   dwBrowseDirection,
	/* [in, string] */            LPCWSTR              szString)
{
	if (dwBrowseDirection == OPC_BROWSE_UP) {
		LOGFMTI("IOPCBrowseServerAddressSpace::ChangeBrowsePosition: UP");
	}
	else {                                    // DOWN or TO
		if (szString) {
			if (dwBrowseDirection == OPC_BROWSE_DOWN) {
				USES_CONVERSION;
				LOGFMTI("IOPCBrowseServerAddressSpace::ChangeBrowsePosition: Down into %s", W2A(szString));
			}
			else {
				USES_CONVERSION;
				LOGFMTI("IOPCBrowseServerAddressSpace::ChangeBrowsePosition: To %s", W2A(szString));
			}
		}
		else {
			if (dwBrowseDirection == OPC_BROWSE_DOWN) {
				LOGFMTI("IOPCBrowseServerAddressSpace::ChangeBrowsePosition: Down but branch is NULL");
				LOGFMTI("   Direction: OPC_BROWSE_DOWN   Branch is NULL");
			}
			else
				LOGFMTI("IOPCBrowseServerAddressSpace::ChangeBrowsePosition: To the 'root'");
		}
	}

	if ((dwBrowseDirection == OPC_BROWSE_DOWN) && (szString == NULL)) {
		return E_INVALIDARG;                      // Branch must not be NULL.
	}

	// Use current position as initial position.
	BSTR bstrNewPosition = m_pBrowseData->m_bstrServerBrowsePosition;

	HRESULT hr = m_pServerHandler->OnBrowseChangeAddressSpacePosition(
		dwBrowseDirection,
		szString,
		&bstrNewPosition,
		&m_pBrowseData->m_pCustomData);
	if (SUCCEEDED(hr)) {
		SysFreeString(m_pBrowseData->m_bstrServerBrowsePosition);
		m_pBrowseData->m_bstrServerBrowsePosition = bstrNewPosition;
	}
	return hr;
}



//=========================================================================
// IOPCBrowseServerAddressSpace::BrowseOPCItemIDs                 INTERFACE
//=========================================================================
STDMETHODIMP DaBrowse::BrowseOPCItemIDs(
	/* [in] */                    OPCBROWSETYPE        dwBrowseFilterType,
	/* [in, string] */            LPCWSTR              szFilterCriteria,
	/* [in] */                    VARTYPE              vtDataTypeFilter,
	/* [in] */                    DWORD                dwAccessRightsFilter,
	/* [out] */                   LPENUMSTRING      *  ppIEnumString)
{
	HRESULT           hres;
	BSTR              *pItemIDs;
	DWORD             nItemIDs;
	DWORD             i;

	USES_CONVERSION;
	LOGFMTI("IOPCBrowseServerAddressSpace::BrowseOPCItemIDs: Filter=%s", W2A(szFilterCriteria));

	if (ppIEnumString == NULL) {
		LOGFMTE("BrowseOPCItemIDs() failed with invalid argument(s): ppIEnumString is NULL");
		return E_INVALIDARG;
	}
	*ppIEnumString = NULL;

	// Check the passed arguments
	if (dwAccessRightsFilter & ~(OPC_READABLE | OPC_WRITEABLE)) {
		LOGFMTE("BrowseOPCItemIDs() failed with invalid argument(s): dwAccessRightsFilter with unknown type");
		return E_INVALIDARG;                      // Illegal Access Rights Filter
	}
	if ((dwBrowseFilterType != OPC_BRANCH) &&
		(dwBrowseFilterType != OPC_LEAF) &&
		(dwBrowseFilterType != OPC_FLAT)) {
		LOGFMTE("BrowseOPCItemIDs() failed with invalid argument(s): dwBrowseFilterType with unknown type");
		return E_INVALIDARG;
	}

	pItemIDs = NULL;
	nItemIDs = 0;

	hres = m_pServerHandler->OnBrowseItemIdentifiers(
		m_pBrowseData->m_bstrServerBrowsePosition,
		dwBrowseFilterType,
		(LPWSTR)szFilterCriteria,
		vtDataTypeFilter,
		dwAccessRightsFilter,
		&nItemIDs,
		&pItemIDs,
		&m_pBrowseData->m_pCustomData);

	if (FAILED(hres)) {
		return hres;                                 // Cannot build a snapshot
	}

	COpcComEnumString *pEnum = new COpcComEnumString;
	if (pEnum == NULL) {
		hres = E_OUTOFMEMORY;
	}
	else {


		try {
			hres = pEnum->Init(pItemIDs,
				&pItemIDs[nItemIDs],
				(IOPCServer *)this,  // This will do an AddRef to the parent
				AtlFlagCopy);
		}
		catch (HRESULT hresEx) {
			hres = hresEx;
		}

		if (SUCCEEDED(hres)) {
			hres = pEnum->_InternalQueryInterface(IID_IEnumString, (LPVOID *)ppIEnumString);
		}

		if (FAILED(hres)) {
			delete pEnum;
		}
		else {
			hres = nItemIDs ? S_OK : S_FALSE;
		}

	} // if Enum Object created successfully

													// Release the snapshot
	if (pItemIDs != NULL) {
		for (i = 0; i < nItemIDs; i++) {
			if (pItemIDs[i] != NULL) {
				SysFreeString(pItemIDs[i]);
			}
		}
		delete[] pItemIDs;
	}
	return hres;
}



//=========================================================================
// IOPCBrowseServerAddressSpace::GetItemID                        INTERFACE
//=========================================================================
STDMETHODIMP DaBrowse::GetItemID(
	/* [in] */                    LPWSTR               szItemDataID,
	/* [out, string] */           LPWSTR            *  szItemID)
{
	USES_CONVERSION;
	LOGFMTI("IOPCBrowseServerAddressSpace::GetItemID of item %s", W2A(szItemDataID));

	if (szItemID == NULL) {
		LOGFMTE("GetItemID() failed with invalid argument(s): szItemID is NULL");
		return E_INVALIDARG;
	}
	*szItemID = NULL;

	BSTR     bstr = NULL;
	HRESULT  hres = m_pServerHandler->OnBrowseGetFullItemIdentifier(
		m_pBrowseData->m_bstrServerBrowsePosition,
		szItemDataID, &bstr,
		&m_pBrowseData->m_pCustomData);

	if (SUCCEEDED(hres)) {
		// Custom Interface needs Global Memory
		*szItemID = WSTRClone(bstr, pIMalloc);
		SysFreeString(bstr);
		if (*szItemID == NULL) {
			hres = E_OUTOFMEMORY;
		}
	}
	return hres;
}



//=========================================================================
// IOPCBrowseServerAddressSpace::BrowseAccessPaths                INTERFACE
// -----------------------------------------------
//    Returns all the access paths for an item id
//=========================================================================
STDMETHODIMP DaBrowse::BrowseAccessPaths(
	/* [in, string] */            LPCWSTR              szItemID,
	/* [out] */                   LPENUMSTRING      *  ppIEnumString)
{
	HRESULT           hres;
	LPWSTR            *pAccessPaths;
	DWORD             nAccessPaths;
	DWORD             i;

	USES_CONVERSION;
	LOGFMTI("IOPCBrowseServerAddressSpace::BrowseAccessPaths for item %s", W2A(szItemID));

	if (ppIEnumString == NULL) {
		LOGFMTE("BrowseAccessPaths() failed with invalid argument(s):  ppIEnumString is NULL");
		return E_INVALIDARG;
	}

	*ppIEnumString = NULL;

	hres = m_pServerHandler->OnBrowseAccessPaths(
		(LPWSTR)(szItemID),
		&nAccessPaths,
		&pAccessPaths,
		&m_pBrowseData->m_pCustomData);

	if (FAILED(hres)) {
		return hres;                                 // Cannot build a snapshot
	}

	COpcComEnumString *pEnum = new COpcComEnumString;
	if (pEnum == NULL) {
		hres = E_OUTOFMEMORY;
	}
	else {

		try {
			hres = pEnum->Init(pAccessPaths,
				&pAccessPaths[nAccessPaths],
				(IOPCServer *)this,  // This will do an AddRef to the parent
				AtlFlagCopy);
		}
		catch (HRESULT hresEx) {
			hres = hresEx;
		}

		if (SUCCEEDED(hres)) {
			hres = pEnum->_InternalQueryInterface(IID_IEnumString, (LPVOID *)ppIEnumString);
		}

		if (FAILED(hres)) {
			delete pEnum;
		}

	} // if Enum Object created successfully

													// Release the snapshot
	if (pAccessPaths != NULL) {
		for (i = 0; i < nAccessPaths; i++) {
			if (pAccessPaths[i] != NULL) {
				SysFreeString(pAccessPaths[i]);
			}
		}
		delete[] pAccessPaths;
	}
	return hres;
}

///////////////////////////////////////////////////////////////////////////


//=========================================================================
// OPC 3.0 Custom Interfaces
//=========================================================================

///////////////////////////////////////////////////////////////////////////
//////////////////////////// IOPCBrowse ///////////////////////////////////
///////////////////////////////////////////////////////////////////////////

// Tokens to generate unique Continuation Points
#define BRANCHTOKENSTRING  L"BRANCH#{81aac52d-b4bb-42d1-94e4-3eca1fd3c45b}#"
#define ITEMTOKENSTRING    L"ITEM#{2f6b782d-50a8-434c-95bf-5d5854a00259}#"

//=========================================================================
// IOPCBrowse::GetProperties                                      INTERFACE
// -------------------------
//    Returns an array of OPCITEMPROPERTIES with one element
//    for each specified item.
//=========================================================================
STDMETHODIMP DaBrowse::GetProperties(
	/* [in] */                    DWORD                dwItemCount,
	/* [size_is][string][in] */   LPWSTR            *  pszItemIDs,
	/* [in] */                    BOOL                 bReturnPropertyValues,
	/* [in] */                    DWORD                dwPropertyCount,
	/* [size_is][in] */           DWORD             *  pdwPropertyIDs,
	/* [size_is][size_is][out] */ OPCITEMPROPERTIES ** ppItemProperties)
{
	LOGFMTI("IOPCBrowse::GetProperties");

	VARIANT value;
	VariantInit(&value);
	*ppItemProperties = NULL;                    // Note : Proxy/Stub checks if the pointers are NULL

	if (dwItemCount == 0) {
		LOGFMTE("GetProperties() failed with invalid argument(s): dwItemCount is 0");
		return E_INVALIDARG;
	}

	OPCITEMPROPERTIES* pProps = ComAlloc<OPCITEMPROPERTIES>(dwItemCount);
	if (!pProps) return E_OUTOFMEMORY;

	HRESULT hrRet = S_OK;
	for (DWORD i = 0; i < dwItemCount; i++) {    // Get Properties for all specified Items

		pProps[i].hrErrorID = OPC_E_UNKNOWNITEMID;
		pProps[i].dwNumProperties = 0;
		pProps[i].pItemProperties = NULL;
		pProps[i].dwReserved = 0;

		DWORD    dwRevisedPropertyCount = 0;
		DWORD*   pdwRevisedPropertyIDs = NULL;
		HRESULT* phrRevisedErrorIDs = NULL;

		LPVOID   pCookie = NULL;
		BOOL     fReleaseCookie = FALSE;
		try {

			HRESULT hr = GetRevisedPropertyIDs(
				pszItemIDs[i],
				dwPropertyCount,
				pdwPropertyIDs,
				&dwRevisedPropertyCount,
				&pdwRevisedPropertyIDs,
				&phrRevisedErrorIDs,
				&pCookie);

			_OPC_CHECK_HR(hr);
			fReleaseCookie = TRUE;

			pProps[i].pItemProperties = ComAlloc<OPCITEMPROPERTY>(dwRevisedPropertyCount);
			_OPC_CHECK_PTR(pProps[i].pItemProperties);

			InitArrayOfOPCITEMPROPERTY(dwRevisedPropertyCount, pProps[i].pItemProperties);
			pProps[i].dwNumProperties = dwRevisedPropertyCount;

			pProps[i].hrErrorID = S_OK;
			// Get Information for all requested Properties of an Item
			for (DWORD z = 0; z < dwRevisedPropertyCount; z++) {
				pProps[i].pItemProperties[z].hrErrorID = phrRevisedErrorIDs[z];
				if (SUCCEEDED(pProps[i].pItemProperties[z].hrErrorID)) {

					try {
						//
						// Attach vtDataType, dwPropertyID, szDescription
						//
						{
							DaItemProperty* pItemProp = NULL;
							if (FAILED(m_pServerHandler->GetItemProperty(pdwRevisedPropertyIDs[z], &pItemProp))) {
								throw OPC_E_INVALID_PID;
							}
							pProps[i].pItemProperties[z].vtDataType = pItemProp->vtDataType;
							pProps[i].pItemProperties[z].dwPropertyID = pItemProp->dwPropID;
							pProps[i].pItemProperties[z].szDescription = ComAllOPCtring(pItemProp->pwszDescr);
							_OPC_CHECK_PTR(pProps[i].pItemProperties[z].szDescription);
						}

						//
						// Attach szItemID
						//
						{
							if (pdwRevisedPropertyIDs[z] > OPC_PROPERTY_EU_INFO) {   // Not for ID Set1 properties
								HRESULT hrLookup = m_pServerHandler->OnLookupItemID(
									pszItemIDs[i],
									pdwRevisedPropertyIDs[z],
									&pProps[i].pItemProperties[z].szItemID,
									pCookie);

								if (FAILED(hrLookup)) {
									// Note: E_FAIL is a valid result code if there is no associated Item
									if (hrLookup == E_FAIL) {
										pProps[i].pItemProperties[z].szItemID = ComAllOPCtring(L"");
										_OPC_CHECK_PTR(pProps[i].pItemProperties[z].szItemID);
									}
									else {
										throw hrLookup;
									}

								}
							}
							else {
								pProps[i].pItemProperties[z].szItemID = ComAllOPCtring(L"");
								_OPC_CHECK_PTR(pProps[i].pItemProperties[z].szItemID);
							}
						}

						//
						// Attach vValue
						//
						{
							if (bReturnPropertyValues) {
								HRESULT hrGet = m_pServerHandler->OnGetItemProperty(
									pszItemIDs[i],
									pdwRevisedPropertyIDs[z],
									&value,
									pCookie);
								VariantCopy(&pProps[i].pItemProperties[z].vValue, &value);
								_OPC_CHECK_HR(hrGet);
							}
						}

					}
					catch (HRESULT hrEx) {
						pProps[i].pItemProperties[z].hrErrorID = hrEx;
					}
					catch (...) {
						pProps[i].pItemProperties[z].hrErrorID = OPC_E_INVALID_PID;
					}

				} // Only if Property ID is specified for this item

				if (FAILED(pProps[i].pItemProperties[z].hrErrorID)) {
					ReleaseOPCITEMPROPERTY(&pProps[i].pItemProperties[z]);
					pProps[i].hrErrorID = S_FALSE;
				}
				VariantClear(&value);

			} // Get Information for all requested Properties of an Item
		}
		catch (HRESULT hrEx) {
			pProps[i].hrErrorID = hrEx;
		}
		catch (...) {
			pProps[i].hrErrorID = E_FAIL;
		}

		if (FAILED(pProps[i].hrErrorID) || pProps[i].hrErrorID == S_FALSE) {
			hrRet = S_FALSE;
		}

		// Release temporary used resources
		if (pdwRevisedPropertyIDs) {
			delete[] pdwRevisedPropertyIDs;
		}
		if (phrRevisedErrorIDs) {
			delete[] phrRevisedErrorIDs;
		}

		if (fReleaseCookie) {
			m_pServerHandler->OnReleasePropertyCookie(pCookie);
			// Ignore return code, do not override the original error code
		}

	} // Get Properties for all specified Items

	*ppItemProperties = pProps;
	VariantClear(&value);
	return hrRet;
}



//=========================================================================
// IOPCBrowse::Browse                                             INTERFACE
// ------------------
//    Browses a single branch of the Server Address Spacend returns zero
//    ore more elements of structure OPCBROWSEELEMENT.
//=========================================================================
STDMETHODIMP DaBrowse::Browse(
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
	/* [size_is][size_is][out] */ OPCBROWSEELEMENT  ** ppBrowseElements)
{
	DWORD    dwNumOfItemIDs = 0;
	DWORD    dwNumOfBranchIDs = 0;
	BSTR*    pItemIDs = NULL;
	BSTR*    pBranchIDs = NULL;
	HRESULT  hr = S_OK;

	LOGFMTI("IOPCBrowse::Browse");
	// Note : Proxy/Stub checks if the pointers are NULL
	*pbMoreElements = FALSE;                     // We support Continuation Point
	*pdwCount = 0;
	*ppBrowseElements = NULL;

	OPCBROWSEELEMENT* pElements = NULL;
	try {
		// Check Parameters
		if ((dwBrowseFilter != OPC_BROWSE_FILTER_ALL) &&
			(dwBrowseFilter != OPC_BROWSE_FILTER_BRANCHES) &&
			(dwBrowseFilter != OPC_BROWSE_FILTER_ITEMS)) {
			LOGFMTE("Browse() failed with invalid argument(s): dwBrowseFilter with unknown type");
			throw E_INVALIDARG;
		}


		// Workaround for non-compliant clients which doesn't follow the OPC Specification.
		// Clients must pass a NUL string in the initial call (and not a NULL pointer)
		if (*pszContinuationPoint == NULL) {
			*pszContinuationPoint = ComAllOPCtring(L"");
			_OPC_CHECK_PTR(*pszContinuationPoint);
		}


		//
		// Change current position to the specified position
		//
		{
			// Use current position as initial position.
			BSTR bstrNewPosition = m_pBrowseData->m_bstrServerBrowsePosition3;

			if (**pszContinuationPoint == L'\0') {
				// There is no Continuation Point specified

// Change current position to the specified position
				hr = m_pServerHandler->OnBrowseChangeAddressSpacePosition(
					OPC_BROWSE_TO,
					szItemID,
					&bstrNewPosition,
					&m_pBrowseData->m_pCustomData);
				if (FAILED(hr)) {
					if (hr != E_OUTOFMEMORY) throw OPC_E_UNKNOWNITEMID;
					throw hr;
				}

				// Store the new position
				SysFreeString(m_pBrowseData->m_bstrServerBrowsePosition3);
				m_pBrowseData->m_bstrServerBrowsePosition3 = bstrNewPosition;
			}
		}


		//
		// Read the requested Elements from the current position
		//

		if (dwBrowseFilter == OPC_BROWSE_FILTER_ALL || dwBrowseFilter == OPC_BROWSE_FILTER_BRANCHES) {

			hr = m_pServerHandler->OnBrowseItemIdentifiers(
				m_pBrowseData->m_bstrServerBrowsePosition3,
				OPC_BRANCH,
				szVendorFilter,
				VT_EMPTY,   // Data Type Filter off
				0,          // Access Rights Filter off
				&dwNumOfBranchIDs,
				&pBranchIDs,
				&m_pBrowseData->m_pCustomData);

			_OPC_CHECK_HR(hr);                 // Cannot build a snapshot
		}

		if (dwBrowseFilter == OPC_BROWSE_FILTER_ALL || dwBrowseFilter == OPC_BROWSE_FILTER_ITEMS) {

			hr = m_pServerHandler->OnBrowseItemIdentifiers(
				m_pBrowseData->m_bstrServerBrowsePosition3,
				OPC_LEAF,
				szVendorFilter,
				VT_EMPTY,   // Data Type Filter off
				0,          // Access Rights Filter off
				&dwNumOfItemIDs,
				&pItemIDs,
				&m_pBrowseData->m_pCustomData);

			_OPC_CHECK_HR(hr);                 // Cannot build a snapshot
		}

		//
		// Continuation Point Handling
		//
		// Discards all Elements in front of the specified Continuation Point
		if (**pszContinuationPoint != L'\0') {
			// A Continuation Point is specified
// Check if the CP is a Branch-Element
			if (wcsncmp(*pszContinuationPoint, BRANCHTOKENSTRING, wcslen(BRANCHTOKENSTRING)) == 0) {

				LPCWSTR pBranch = &(*pszContinuationPoint)[wcslen(BRANCHTOKENSTRING)];
				dwNumOfBranchIDs = RemoveElementsInFrontOfContinuationPoint(pBranch, dwNumOfBranchIDs, pBranchIDs);

				if (dwNumOfBranchIDs == 0)          // At least the CP should exist as Element if valid
					hr = OPC_E_INVALIDCONTINUATIONPOINT;
			}
			// Check if the CP is a Leaf-Element
			else if (wcsncmp(*pszContinuationPoint, ITEMTOKENSTRING, wcslen(ITEMTOKENSTRING)) == 0) {

				LPCWSTR pItem = &(*pszContinuationPoint)[wcslen(ITEMTOKENSTRING)];
				dwNumOfItemIDs = RemoveElementsInFrontOfContinuationPoint(pItem, dwNumOfItemIDs, pItemIDs);

				if (dwNumOfItemIDs == 0)            // At least the CP should exist as Element if valid
					hr = OPC_E_INVALIDCONTINUATIONPOINT;
			}
			else {                                 // Invalid CP sytnax
				hr = OPC_E_INVALIDCONTINUATIONPOINT;
			}

			_OPC_CHECK_HR(hr);

			ComFreeString(*pszContinuationPoint);
			*pszContinuationPoint = ComAllOPCtring(L"");
			_OPC_CHECK_PTR(*pszContinuationPoint);
		}


		//
		// Standard Filtering
		//
		if (*szElementNameFilter) {
			dwNumOfBranchIDs = FilterElements(szElementNameFilter, dwNumOfBranchIDs, pBranchIDs);
			dwNumOfItemIDs = FilterElements(szElementNameFilter, dwNumOfItemIDs, pItemIDs);
		}


		//
		// Result Limitation
		//
		DWORD dwNumOfElements = dwNumOfBranchIDs + dwNumOfItemIDs;

		if (dwMaxElementsReturned) {
			if (dwNumOfElements > dwMaxElementsReturned) {
				dwNumOfElements = dwMaxElementsReturned;
			}
		}


		//
		// Initialize Result Buffer
		//
		OPCBROWSEELEMENT* pElements = ComAlloc<OPCBROWSEELEMENT>(dwNumOfElements);
		_OPC_CHECK_PTR(pElements)

			memset(pElements, 0, dwNumOfElements * sizeof(OPCBROWSEELEMENT));


		DWORD dwBrowseElementCount = 0;
		hr = MoveElementIDsToBrowseElements(
			OPC_BROWSE_HASCHILDREN,
			dwNumOfBranchIDs,
			pBranchIDs,
			dwNumOfElements,
			&dwBrowseElementCount,
			pElements,
			pszContinuationPoint,
			// We have no Properties for Branches
			FALSE, FALSE, 0, NULL);
		_OPC_CHECK_HR(hr);

		hr = MoveElementIDsToBrowseElements(
			OPC_BROWSE_ISITEM,
			dwNumOfItemIDs,
			pItemIDs,
			dwNumOfElements,
			&dwBrowseElementCount,
			pElements,
			pszContinuationPoint,
			// Properties Handling
			bReturnAllProperties,
			bReturnPropertyValues,
			dwPropertyCount,
			pdwPropertyIDs);
		_OPC_CHECK_HR(hr);

		*pdwCount = dwNumOfElements;
		*ppBrowseElements = pElements;

		//hr = S_OK;
	}
	catch (HRESULT hrEx) {
		hr = hrEx;
	}
	catch (...) {
		hr = E_FAIL;
	}

	if (FAILED(hr)) {
		ComFree(pElements);
		ComFreeString(*pszContinuationPoint);
		*pszContinuationPoint = NULL;
	}

	// Release temporary used resources
	if (pBranchIDs != NULL) {
		delete[] pBranchIDs;            // Note: Elements are already released
	}
	if (pItemIDs != NULL) {
		delete[] pItemIDs;
	}

	return hr;
}

///////////////////////////////////////////////////////////////////////////


//-------------------------------------------------------------------------
// IMPLEMENTATTION
//-------------------------------------------------------------------------

//=========================================================================
// GetRevisedPropertyIDs                                           INTERNAL
// ---------------------
//    Returns the requested Property IDs for the specified item. See
//    parameter description for the behavior.
//
//    The caller is responsible for freeing the memory allocated for the
//    returned array by calling the delete [] operator.
//
//    If the function call succeeds then it's required to release the
//    returned 'ccokie' with the method
//    DaBaseServer::OnReleasePropertyCookie().
//
// Parameters:
//    szItemID                The ItemID for which the caller wants the
//                            Prroperty IDs. Is only used if
//                            dwPropertyCount is 0.
//    dwPropertyCount         The number of Property IDs in
//                            pdwPropertyIDs. If 0 all available Property
//                            IDs for the specified item will be returned.
//                            Otherwise the Property IDs specified in
//                            pdwPropertyIDs will be returned.
//    pdwPropertyIDs          The predefined Property IDs. Are only used
//                            if dwPropertyCount is not 0.
//    pdwRevisedPropertyCount The number of Property IDs returned in
//                            ppdwRevisedPropertyIDs.The number of error
//                            codes returned in pphrErrorID.
//    ppdwRevisedPropertyIDs  The returned Property IDs.
//    pphrErrorID             The returned error codes (OPC_E_INVALID_PID
//                            if the Property ID is not specified for the
//                            item; otherwise S_OK)
//    ppCookie                used for following Item Property related
//                            DaBaseServer::On...() function calls.
//=========================================================================
HRESULT DaBrowse::GetRevisedPropertyIDs(
	/* [in] */                    const LPWSTR         szItemID,
	/* [in] */                    const DWORD          dwPropertyCount,
	/* [size_is][in] */           const DWORD       *  pdwPropertyIDs,
	/* [out] */                   DWORD             *  pdwRevisedPropertyCount,
	/* [size_is][size_is][out] */ DWORD             ** ppdwRevisedPropertyIDs,
	/* [size_is][size_is][out] */ HRESULT           ** pphrErrorID,
	/* [out] */                   LPVOID            *  ppCookie)
{
	*pdwRevisedPropertyCount = 0;
	*ppdwRevisedPropertyIDs = NULL;
	*pphrErrorID = NULL;

	HRESULT  hr = S_OK;
	DWORD    dwTmpPropertyCount = 0;
	DWORD*   pdwTmpPropertyIDs = NULL;

	try {
		//
		// Get all avaliable Property IDs for the specified item
		//
		*ppCookie = NULL;
		HRESULT hr = m_pServerHandler->OnQueryItemProperties(
			szItemID,
			&dwTmpPropertyCount,
			&pdwTmpPropertyIDs,
			ppCookie);
		_OPC_CHECK_HR(hr);

		//
		// Return the requested Property IDs inclusive result code (Error ID)
		//
		if (dwPropertyCount == 0) {               // Return all availabe Property IDs of this item

			*pphrErrorID = new HRESULT[dwTmpPropertyCount];
			_OPC_CHECK_PTR(*pphrErrorID);

			for (DWORD i = 0; i < dwTmpPropertyCount; i++) {
				(*pphrErrorID)[i] = S_OK;
			}

			*pdwRevisedPropertyCount = dwTmpPropertyCount;
			*ppdwRevisedPropertyIDs = pdwTmpPropertyIDs;
			pdwTmpPropertyIDs = NULL;              // Avoid release
		}
		else {                                    // Return only the requested Property IDs
												  // The revised IDs are identically with the predefined IDs
			*pphrErrorID = new HRESULT[dwPropertyCount];
			_OPC_CHECK_PTR(*pphrErrorID);

			*ppdwRevisedPropertyIDs = new DWORD[dwPropertyCount];
			_OPC_CHECK_PTR(*ppdwRevisedPropertyIDs);

			*pdwRevisedPropertyCount = dwPropertyCount;
			memcpy(*ppdwRevisedPropertyIDs, pdwPropertyIDs, dwPropertyCount * sizeof(DWORD));

			// Check if the specified Property IDs are defined for the item
			for (DWORD i = 0; i < dwPropertyCount; i++) {
				(*pphrErrorID)[i] = OPC_E_INVALID_PID;
				for (DWORD z = 0; z < dwTmpPropertyCount; z++) {
					if (pdwPropertyIDs[i] == pdwTmpPropertyIDs[z]) {
						(*pphrErrorID)[i] = S_OK;
						break;
					}
				}
			}
		}
	}
	catch (HRESULT hrEx) {
		hr = hrEx;
	}
	catch (...) {
		hr = E_FAIL;
	}

	if (FAILED(hr)) {

		*pdwRevisedPropertyCount = 0;

		if (*pphrErrorID) {
			delete[] * pphrErrorID;
			*pphrErrorID = NULL;
		}

		if (*ppdwRevisedPropertyIDs) {
			delete[] * ppdwRevisedPropertyIDs;
			*ppdwRevisedPropertyIDs = NULL;
		}

		if (*ppCookie) {
			m_pServerHandler->OnReleasePropertyCookie(*ppCookie);
			// Ignore return code, do not override the original error code
			*ppCookie = NULL;
		}
	}

	if (pdwTmpPropertyIDs) {
		delete[] pdwTmpPropertyIDs;
	}

	return hr;
}



//=========================================================================
// MoveElementIDsToBrowseElements                                  INTERNAL
// ------------------------------
//    Initializes elements of structure OPCBROWSEELEMENT with the
//    specified Element IDs as source.
//
//    - The number of elements to be copied can be specified.
//    - All specified Element IDs are released (in any case)
//    - If something goes wrong then all OPCBROWSEELEMENTS are released
//       (also the OPCBROWSEELEMENTS initialized by previous calls)
//    - Only elements are removed from the arrays and not the
//       arrays itself.
//
// Parameters:
//    dwElementType           The type of the Element IDs specified in 
//                            pElementIDs. Valid values are 
//                            OPC_BROWSE_HASCHILDREN and OPC_BROWSE_ISITEM.
//    dwNumOfElementIDs       The number of Element IDs in pElementIDs.
//    pElementIDs             The Element IDs, Branch or Item Names
//                            from the Server Address Space.
//                            All strings are released.
//    dwNumOfElements         The number of OPCBROWSEELEMENTS in 
//                            pElements (array size, counts initialized
//                            and free elements). Ensure that the buffer
//                            is initialized with 0 before first usage.
//    pdwElementCount         [in]  Index of the first OPCBROWSEELEMENT
//                                  to be used
//                            [out] Index of the next OPCBROWSEELEMENT
//                                  to be used
//    pElements               Array with the OPCBROWSEELEMENTS to be
//                            initialized. Are all released if something
//                            goes wrong.
//    pszContinuationPoint    If not all specified Element IDs can be
//                            'moved' then a 'Continuation Point' will
//                            be set.
//    Property Related Parameters
//                            Identicall with parameters of
//                            method GetProerties()
//
//=========================================================================
HRESULT DaBrowse::MoveElementIDsToBrowseElements(
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
	/* [in(dwPropertyCount)] */   DWORD             *  pdwPropertyIDs)
{
	_ASSERTE(*pdwElementCount <= dwNumOfElements);
	_ASSERTE(dwElementType == OPC_BROWSE_HASCHILDREN || dwElementType == OPC_BROWSE_ISITEM);

	DWORD    i;
	HRESULT  hr = S_OK;
	HRESULT  hrReturn = S_OK;

	try {

		for (i = 0; i < dwNumOfElementIDs; i++) {

			if (*pdwElementCount < dwNumOfElements) {

				OPCBROWSEELEMENT* pEl = &pElements[*pdwElementCount];

				pEl->szName = ComAllOPCtring(pElementIDs[i]);
				_OPC_CHECK_PTR(pEl->szName);

				BSTR bstrTmp = NULL;
				hr = m_pServerHandler->OnBrowseGetFullItemIdentifier(
					m_pBrowseData->m_bstrServerBrowsePosition3,
					pElementIDs[i], &bstrTmp,
					&m_pBrowseData->m_pCustomData);
				_OPC_CHECK_HR(hr);

				pEl->szItemID = ComAllOPCtring(bstrTmp);
				SysFreeString(bstrTmp);
				_OPC_CHECK_PTR(pEl->szItemID);

				pEl->dwFlagValue = dwElementType;
				pEl->dwReserved = 0;

				//
				// Handle Properties
				//
				if (bReturnAllProperties || dwPropertyCount > 0) {

					BSTR bstrFullyQualifiedItemID = NULL;

					hr = m_pServerHandler->OnBrowseGetFullItemIdentifier(
						m_pBrowseData->m_bstrServerBrowsePosition3,
						pElementIDs[i],
						&bstrFullyQualifiedItemID,
						&m_pBrowseData->m_pCustomData);
					_OPC_CHECK_HR(hr);

					OPCITEMPROPERTIES* pItemProperties = NULL;

					hr = GetProperties(1,
						&bstrFullyQualifiedItemID,
						bReturnPropertyValues,
						dwPropertyCount,
						pdwPropertyIDs,
						&pItemProperties);

					SysFreeString(bstrFullyQualifiedItemID);
					_OPC_CHECK_HR(hr);
					if (hr == S_FALSE)
					{
						hrReturn = S_FALSE;
					}

					pEl->ItemProperties = *pItemProperties;
					ComFree(pItemProperties);
				}
				else {
					pEl->ItemProperties.dwNumProperties = 0;
					pEl->ItemProperties.pItemProperties = ComAlloc<OPCITEMPROPERTY>(0);
					_OPC_CHECK_PTR(pEl->ItemProperties.pItemProperties);
				}

				(*pdwElementCount)++;
			}
			else if (**pszContinuationPoint == L'\0') {
				WideString wsTemp;                 // Continuation Point is not yet set

				LPCWSTR pToken;
				if (dwElementType == OPC_BROWSE_HASCHILDREN)
					pToken = BRANCHTOKENSTRING;
				else
					pToken = ITEMTOKENSTRING;

				_OPC_CHECK_HRFUNC(wsTemp.SetString(pToken));
				_OPC_CHECK_HRFUNC(wsTemp.AppendString(pElementIDs[i]))
					ComFreeString(*pszContinuationPoint);
				*pszContinuationPoint = wsTemp.CopyCOM();
				_OPC_CHECK_PTR(*pszContinuationPoint);
			}

			SysFreeString(pElementIDs[i]);
		}
	}
	catch (HRESULT hrEx) {
		hr = hrEx;
	}
	catch (...) {
		hr = E_FAIL;
	}


	if (FAILED(hr)) {

		// Release all Browse Elements
		for (i = 0; i < dwNumOfElements; i++) {
			OPCBROWSEELEMENT* pEl = &pElements[i];

			ComFreeString(pEl->szName);
			ComFreeString(pEl->szItemID);

			if (pEl->ItemProperties.pItemProperties) {
				ReleaseArrayOfOPCITEMPROPERTY(
					pEl->ItemProperties.dwNumProperties,
					pEl->ItemProperties.pItemProperties);
				ComFree(pEl->ItemProperties.pItemProperties);
			}
		}
	}

	// Release not yet releases Element IDs
	while (i < dwNumOfElementIDs) {
		SysFreeString(pElementIDs[i]);
		i++;
	}
	return hrReturn;
}



//=========================================================================
// FilterElements                                                  INTERNAL
// --------------
//    Removes all Elements which doesn't match the Standard Filter.
//    The Elements are shifted so there isn't any gap between the Elements.
//    This method returns the number of Elements which passed the filter.
//=========================================================================
DWORD DaBrowse::FilterElements(LPCWSTR     szElementNameFilter,
	DWORD       dwNumOfElements,
	BSTR     *  pElements)
{
	DWORD dwNumOfPassedElements = 0;
	for (DWORD i = 0; i < dwNumOfElements; i++) {
		if (::MatchPattern(pElements[i], szElementNameFilter)) {
			pElements[dwNumOfPassedElements] = pElements[i];
			dwNumOfPassedElements++;
		}
		else {
			SysFreeString(pElements[i]);
		}
	}
	return dwNumOfPassedElements;
}



//=========================================================================
// RemoveElementsInFrontOfContinuationPoint                        INTERNAL
// ----------------------------------------
//    Removes all Elements in front of the specified
//    'Continuation Point' - Element.
//    The Elements are shifted so there isn't any gap between the Elements.
//    This method returns the number of remained Elements.
//=========================================================================
DWORD DaBrowse::RemoveElementsInFrontOfContinuationPoint(
	LPCWSTR     szContinuationPoint,
	DWORD       dwNumOfElements,
	BSTR     *  pElements)
{
	DWORD i;
	for (i = 0; i < dwNumOfElements; i++) {
		if (wcscmp(pElements[i], szContinuationPoint) == 0) {
			break;
		}
		SysFreeString(pElements[i]);
	}

	DWORD dwNumOfPassedElements = dwNumOfElements - i;
	memcpy(pElements, &pElements[i], dwNumOfPassedElements * sizeof(BSTR*));
	return dwNumOfPassedElements;
}



//=========================================================================
// InitArrayOfOPCITEMPROPERTY                                      INTERNAL
//=========================================================================
void DaBrowse::InitArrayOfOPCITEMPROPERTY(DWORD dwNumOfProp, OPCITEMPROPERTY* pProperties)
{
	_ASSERTE(pProperties);                      // Must not be NULL

	for (DWORD i = 0; i < dwNumOfProp; i++) {
		OPCITEMPROPERTY* pProperty = &pProperties[i];

		pProperty->vtDataType = VT_EMPTY;
		pProperty->wReserved = 0;
		pProperty->dwPropertyID = 0;
		pProperty->szItemID = NULL;
		pProperty->szDescription = NULL;
		VariantInit(&pProperty->vValue);
		pProperty->hrErrorID = OPC_E_INVALID_PID;
		pProperty->dwReserved = 0;
	}
}



//=========================================================================
// ReleaseArrayOfOPCITEMPROPERTY                                   INTERNAL
//=========================================================================
void DaBrowse::ReleaseArrayOfOPCITEMPROPERTY(DWORD dwNumOfProp, OPCITEMPROPERTY* pProperties)
{
	_ASSERTE(pProperties);                      // Must not be NULL

	for (DWORD i = 0; i < dwNumOfProp; i++) {
		ReleaseOPCITEMPROPERTY(&pProperties[i]);
	}
}



//=========================================================================
// ReleaseOPCITEMPROPERTY                                          INTERNAL
//=========================================================================
void DaBrowse::ReleaseOPCITEMPROPERTY(OPCITEMPROPERTY* pProperty)
{
	_ASSERTE(pProperty);                       // Must not be NULL

	ComFreeString(pProperty->szItemID);
	ComFreeString(pProperty->szDescription);
	VariantClear(&pProperty->vValue);
}
//DOM-IGNORE-END