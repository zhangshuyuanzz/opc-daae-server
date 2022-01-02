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
#include "UtilityFuncs.h"
#include "AeComServer.h"
#include "AeComSubscription.h"
#include "AeAreaBrowser.h"
#include "AeBaseServer.h"
#include "AeSource.h"
#include "AeCategory.h"
#include "CoreMain.h"

//-------------------------------------------------------------------------
// CODE
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AeComBaseServer::AeComBaseServer()
{
	m_pServerHandler = NULL;
	memset( &m_ftStartTime, 0 , sizeof (m_ftStartTime) );
	memset( &m_ftLastUpdateTime, 0 , sizeof (m_ftLastUpdateTime) );
}



//=========================================================================
// Initializer
// -----------
//    Must be called after construction.
//=========================================================================
HRESULT AeComBaseServer::Create()
{
	HRESULT hres;

	hres = CoFileTimeNow( &m_ftStartTime );
	if (SUCCEEDED( hres )) {
		m_pServerHandler = ::OnGetAeBaseServer();
		if (m_pServerHandler == NULL) {
			hres = E_FAIL;
		}
	}
	if (SUCCEEDED( hres )) {
		hres = m_pServerHandler->OnConnectClient();
	}
	if (SUCCEEDED( hres )) {
		hres = m_pServerHandler->AddServerToList( this );
	}
	return hres;
}



//=========================================================================
// Destructor
//=========================================================================
AeComBaseServer::~AeComBaseServer()
{
	if (m_pServerHandler) {
		m_pServerHandler->OnDisconnectClient();
		m_pServerHandler->RemoveServerFromList( this );
	}
}



//-------------------------------------------------------------------------
// OPERATIONS
//-------------------------------------------------------------------------

///////////////////////////////////////////////////////////////////////////
/////////////////////////////// IOPCEventServer ///////////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCEventServer::GetStatus                                     INTERFACE
// --------------------------
//    Returns the status of the event server.
//=========================================================================
STDMETHODIMP AeComBaseServer::GetStatus(
									 /* [out] */                   OPCEVENTSERVERSTATUS
									 ** ppEventServerStatus )
{
	HRESULT                 hres;
	OPCEVENTSERVERSTATUS*   pESStat;

	*ppEventServerStatus = NULL;

	if (!m_pServerHandler)
		return E_FAIL;

	// Allocate memory for the status structure
	pESStat = ComAlloc<OPCEVENTSERVERSTATUS>();
	if (pESStat == NULL) {
		return E_OUTOFMEMORY;
	}

	LPWSTR pszVendor;

	// Set the time stamps
	hres = CoFileTimeNow( &pESStat->ftCurrentTime );
	if (SUCCEEDED( hres )) {

		pESStat->ftStartTime       = m_ftStartTime;
		pESStat->ftLastUpdateTime  = m_ftLastUpdateTime;

		// Get the current state of the server
		hres = m_pServerHandler->OnGetServerState( pESStat->dwServerState, pszVendor );
	}

	if (SUCCEEDED( hres )) {
		// Set the vendor name
		if (pszVendor) {
			USES_CONVERSION;                       // Custom interface needs global memory for strings
			pESStat->szVendorInfo = WSTRClone( pszVendor, pIMalloc );
			if (pESStat->szVendorInfo == NULL) {
				hres = E_OUTOFMEMORY;
			}
		}
		else {
			_ASSERTE( 0 );                         // The company name string must exist in the Version Info
			hres = E_FAIL;
		}
	}

	if (SUCCEEDED( hres )) {
		// Set the version numbers
		pESStat->wMajorVersion  = core_generic_main.m_VersionInfo.m_wMajor;
		pESStat->wMinorVersion  = core_generic_main.m_VersionInfo.m_wMinor;
		pESStat->wBuildNumber   = core_generic_main.m_VersionInfo.m_wBuild;

		*ppEventServerStatus = pESStat;          // Attach the event server status record to the result pointer.
	}
	else {
		if (pESStat->szVendorInfo) {
			pIMalloc->Free( pESStat->szVendorInfo );
		}
		pIMalloc->Free( pESStat );                // Something goes wrong
	}
	return hres;
}



//=========================================================================
// IOPCEventServer::CreateEventSubscription                       INTERFACE
// ----------------------------------------
//    Creates an event subscription object and adds it to the
//    event server.
//=========================================================================
STDMETHODIMP AeComBaseServer::CreateEventSubscription(
	/* [in] */                    BOOL           bActive,
	/* [in] */                    DWORD          dwBufferTime,
	/* [in] */                    DWORD          dwMaxSize,
	/* [in] */                    OPCHANDLE      hClientSubscription,
	/* [in] */                    REFIID         riid,
	/* [iid_is][out] */           LPUNKNOWN   *  ppUnk,
	/* [out] */                   DWORD       *  pdwRevisedBufferTime,
	/* [out] */                   DWORD       *  pdwRevisedMaxSize )
{
	HRESULT                             hres;
	CComObject<AeComSubscription>*  pCOMEventSubscr;

	*ppUnk = NULL;

	hres = CComObject<AeComSubscription>::CreateInstance( &pCOMEventSubscr );
	if (SUCCEEDED( hres )) {
		hres = pCOMEventSubscr->Create(  this,
			bActive, dwBufferTime, dwMaxSize,
			hClientSubscription,
			pdwRevisedBufferTime, pdwRevisedMaxSize );

		if (SUCCEEDED( hres )) {                  // Get the requested interface
			hres = pCOMEventSubscr->QueryInterface( riid, (LPVOID *)ppUnk );
		}
		if (FAILED( hres )) {
			if (hres == E_NOINTERFACE) {
				hres = E_INVALIDARG;
			}
			delete pCOMEventSubscr;
		}
	}
	return hres;
}



//=========================================================================
// IOPCEventServer::QueryAvailableFilters                         INTERFACE
// --------------------------------------
//    Gets the filter criterias supported by the server.
//    The toolkit supports all types of filter.
//=========================================================================
STDMETHODIMP AeComBaseServer::QueryAvailableFilters(
	/* [out] */                   DWORD       *  pdwFilterMask )
{
	*pdwFilterMask =  OPC_FILTER_BY_EVENT     |
		OPC_FILTER_BY_CATEGORY  |
		OPC_FILTER_BY_SEVERITY  |
		OPC_FILTER_BY_AREA      |
		OPC_FILTER_BY_SOURCE;
	return S_OK;
}



//=========================================================================
// IOPCEventServer::QueryEventCategories                          INTERFACE
// -------------------------------------
//    Gets the Event Categories supported by the server.
//=========================================================================
STDMETHODIMP AeComBaseServer::QueryEventCategories(
	/* [in] */                    DWORD          dwEventType,
	/* [out] */                   DWORD       *  pdwCount,
	/* [size_is][size_is][out] */ DWORD       ** ppdwEventCategories,
	/* [size_is][size_is][out] */ LPWSTR      ** ppszEventCategoryDescs )
{
	*pdwCount = 0;                               // Note : Proxy/Stub checks if the pointers are NULL
	*ppdwEventCategories = NULL;
	*ppszEventCategoryDescs = NULL;

	if (dwEventType == 0) {
		return E_INVALIDARG;
	}
	if (dwEventType & ~OPC_ALL_EVENTS) {
		return E_INVALIDARG;
	}

	if (!m_pServerHandler)
		return E_FAIL;

	return m_pServerHandler->QueryEventCategories(
		dwEventType,
		pdwCount,
		ppdwEventCategories,
		ppszEventCategoryDescs );
}



//=========================================================================
// IOPCEventServer::QueryConditionNames                           INTERFACE
// -------------------------------------
//    Gets the Condition Names of the specified Category.
//=========================================================================
STDMETHODIMP AeComBaseServer::QueryConditionNames(
	/* [in] */                    DWORD          dwEventCategory,
	/* [out] */                   DWORD       *  pdwCount,
	/* [size_is][size_is][out] */ LPWSTR      ** ppszConditionNames )
{
	*pdwCount = 0;                               // Note : Proxy/Stub checks if the pointers are NULL
	*ppszConditionNames = NULL;

	if (!m_pServerHandler)
		return E_FAIL;

	return m_pServerHandler->QueryConditionNames(
		dwEventCategory,
		pdwCount,
		ppszConditionNames );
}



//=========================================================================
// IOPCEventServer::QuerySubConditionNames                        INTERFACE
// ---------------------------------------
//    Gets the Sub Condition Names of the specified Condition.
//=========================================================================
STDMETHODIMP AeComBaseServer::QuerySubConditionNames(
	/* [in] */                    LPWSTR         szConditionName,
	/* [out] */                   DWORD       *  pdwCount,
	/* [size_is][size_is][out] */ LPWSTR      ** ppszSubConditionNames )
{
	*pdwCount = 0;                               // Note : Proxy/Stub checks if the pointers are NULL
	*ppszSubConditionNames = NULL;

	if (!m_pServerHandler)
		return E_FAIL;

	return m_pServerHandler->QuerySubConditionNames(
		szConditionName,
		pdwCount,
		ppszSubConditionNames );
}



//=========================================================================
// IOPCEventServer::QuerySourceConditions                         INTERFACE
// --------------------------------------
//    Gets the Condition Names associated with the specified Source.
//=========================================================================
STDMETHODIMP AeComBaseServer::QuerySourceConditions(
	/* [in] */                    LPWSTR         szSource,
	/* [out] */                   DWORD       *  pdwCount,
	/* [size_is][size_is][out] */ LPWSTR      ** ppszConditionNames )
{
	*pdwCount = 0;                               // Note : Proxy/Stub checks if the pointers are NULL
	*ppszConditionNames = NULL;

	if (!m_pServerHandler)
		return E_FAIL;

	return m_pServerHandler->QuerySourceConditions(
		szSource,
		pdwCount,
		ppszConditionNames );
}



//=========================================================================
// IOPCEventServer::QueryEventAttributes                          INTERFACE
// -------------------------------------
//    Gets the IDs, descriptions and types of all attributes associated
//    with the specified Event Category.
//=========================================================================
STDMETHODIMP AeComBaseServer::QueryEventAttributes(
	/* [in] */                    DWORD          dwEventCategory,
	/* [out] */                   DWORD       *  pdwCount,
	/* [size_is][size_is][out] */ DWORD       ** ppdwAttrIDs,
	/* [size_is][size_is][out] */ LPWSTR      ** ppszAttrDescs,
	/* [size_is][size_is][out] */ VARTYPE     ** ppvtAttrTypes )
{
	*pdwCount = 0;                               // Note : Proxy/Stub checks if the pointers are NULL
	*ppdwAttrIDs = NULL;
	*ppszAttrDescs = NULL;
	*ppvtAttrTypes = NULL;

	if (!m_pServerHandler)
		return E_FAIL;

	return m_pServerHandler->QueryEventAttributes(
		dwEventCategory,
		pdwCount,
		ppdwAttrIDs,
		ppszAttrDescs,
		ppvtAttrTypes );
}



//=========================================================================
// IOPCEventServer::TranslateToItemIDs                            INTERFACE
// -----------------------------------
//    Translates Event Attributes to corresponding Data Access Item IDs.
//=========================================================================
STDMETHODIMP AeComBaseServer::TranslateToItemIDs(
	/* [in] */                    LPWSTR         szSource,
	/* [in] */                    DWORD          dwEventCategory,
	/* [in] */                    LPWSTR         szConditionName,
	/* [in] */                    LPWSTR         szSubconditionName,
	/* [in] */                    DWORD          dwCount,
	/* [size_is][in] */           DWORD       *  pdwAssocAttrIDs,
	/* [size_is][size_is][out] */ LPWSTR      ** ppszAttrItemIDs,
	/* [size_is][size_is][out] */ LPWSTR      ** ppszNodeNames,
	/* [size_is][size_is][out] */ CLSID       ** ppCLSIDs )
{
	*ppszAttrItemIDs  = NULL;                    // Note : Proxy/Stub checks if the pointers are NULL
	*ppszNodeNames    = NULL;
	*ppCLSIDs         = NULL;

	if (!m_pServerHandler)
		return E_FAIL;

	if (dwCount == 0) {
		return E_INVALIDARG;
	}

	return m_pServerHandler->TranslateToItemIDs(
		szSource,          
		dwEventCategory,   
		szConditionName,   
		szSubconditionName,
		dwCount,           
		pdwAssocAttrIDs,   
		ppszAttrItemIDs,   
		ppszNodeNames,     
		ppCLSIDs );
}



//=========================================================================
// IOPCEventServer::GetConditionState                             INTERFACE
// ----------------------------------
//    Returns the current state for the condition instance corresponding
//    to the speicied source and condition definition name.
//=========================================================================
STDMETHODIMP AeComBaseServer::GetConditionState(
	/* [in] */                    LPWSTR         szSource,
	/* [in] */                    LPWSTR         szConditionName,
	/* [in] */                    DWORD          dwNumEventAttrs,
	/* [size_is][in] */           DWORD       *  pdwAttributeIDs,
	/* [out] */                   OPCCONDITIONSTATE
	** ppConditionState )
{
	*ppConditionState = NULL;                    // Note : Proxy/Stub checks if the pointers are NULL

	if (!m_pServerHandler)
		return E_FAIL;

	return m_pServerHandler->GetConditionState(
		szSource,
		szConditionName,
		dwNumEventAttrs,
		pdwAttributeIDs,
		ppConditionState );
}



//=========================================================================
// IOPCEventServer::EnableConditionByArea                         INTERFACE
// --------------------------------------
//    Enables all conditions for all sources within the specified area.
//=========================================================================
STDMETHODIMP AeComBaseServer::EnableConditionByArea(
	/* [in] */                    DWORD          dwNumAreas,
	/* [size_is][in] */           LPWSTR      *  pszAreas )
{
	if (!m_pServerHandler)
		return E_FAIL;

	return m_pServerHandler->EnableConditions(
		TRUE,                // Enable
		FALSE,               // By Area
		dwNumAreas,          // Number of Areas
		pszAreas );          // Area Names
}



//=========================================================================
// IOPCEventServer::EnableConditionBySource                       INTERFACE
// ----------------------------------------
//    Enables all conditions for the specified sources.
//=========================================================================
STDMETHODIMP AeComBaseServer::EnableConditionBySource(
	/* [in] */                    DWORD          dwNumSources,
	/* [size_is][in] */           LPWSTR      *  pszSources )
{
	if (!m_pServerHandler)
		return E_FAIL;

	return m_pServerHandler->EnableConditions(
		TRUE,                // Enable
		TRUE,                // By Source
		dwNumSources,        // Number of Sources
		pszSources );        // Source Names
}



//=========================================================================
// IOPCEventServer::DisableConditionByArea                        INTERFACE
// ---------------------------------------
//    Disables all conditions for all sources within the specified area.
//=========================================================================
STDMETHODIMP AeComBaseServer::DisableConditionByArea(
	/* [in] */                    DWORD          dwNumAreas,
	/* [size_is][in] */           LPWSTR      *  pszAreas )
{
	if (!m_pServerHandler)
		return E_FAIL;

	return m_pServerHandler->EnableConditions(
		FALSE,               // Disable
		FALSE,               // By Area
		dwNumAreas,          // Number of Areas
		pszAreas );          // Area Names
}



//=========================================================================
// IOPCEventServer::DisableConditionBySource                      INTERFACE
// -----------------------------------------
//    Disables all conditions for the specified sources.
//=========================================================================
STDMETHODIMP AeComBaseServer::DisableConditionBySource(
	/* [in] */                    DWORD          dwNumSources,
	/* [size_is][in] */           LPWSTR      *  pszSources )
{
	if (!m_pServerHandler)
		return E_FAIL;

	return m_pServerHandler->EnableConditions(
		FALSE,               // Disable
		TRUE,                // By Source
		dwNumSources,        // Number of Sources
		pszSources );        // Source Names
}



//=========================================================================
// IOPCEventServer::AckCondition                                  INTERFACE
// -----------------------------
//    Acknowledges one or more conidtions in the event sever.
//=========================================================================
STDMETHODIMP AeComBaseServer::AckCondition(
										/* [in] */                    DWORD          dwCount,
										/* [string][in] */            LPWSTR         szAcknowledgerID,
										/* [string][in] */            LPWSTR         szComment,
										/* [size_is][in] */           LPWSTR      *  pszSource,
										/* [size_is][in] */           LPWSTR      *  pszConditionName,
										/* [size_is][in] */           FILETIME    *  pftActiveTime,
										/* [size_is][in] */           DWORD       *  pdwCookie,
										/* [size_is][size_is][out] */ HRESULT     ** ppErrors )
{
	*ppErrors = NULL;                            // Note : Proxy/Stub checks if the pointers are NULL

	if (!m_pServerHandler)
		return E_FAIL;

	if (dwCount == 0) {
		return E_INVALIDARG;
	}

	if (*szAcknowledgerID == NULL) {             // A NULL string is not allowed
		return E_INVALIDARG;
	}

	return m_pServerHandler->AckCondition(
		dwCount,
		szAcknowledgerID,
		szComment,
		pszSource,
		pszConditionName,
		pftActiveTime,
		pdwCookie,
		ppErrors );
}



//=========================================================================
// IOPCEventServer::CreateAreaBrowser                             INTERFACE
// ----------------------------------
//    Creates an area browser object.
//=========================================================================
STDMETHODIMP AeComBaseServer::CreateAreaBrowser(
	/* [in] */                    REFIID         riid,
	/* [iid_is][out] */           LPUNKNOWN   *  ppUnk )
{
	HRESULT                             hres;
	CComObject<EventAreaBrowser>*   pCOMAreaBrowser;

	*ppUnk = NULL;                               // Note : Proxy/Stub checks if the pointers are NULL

	hres = CComObject<EventAreaBrowser>::CreateInstance( &pCOMAreaBrowser );
	if (SUCCEEDED( hres )) {
		hres = pCOMAreaBrowser->Create( &m_pServerHandler->m_RootArea );
		if (SUCCEEDED( hres )) {                  // Get the requested interface
			hres = pCOMAreaBrowser->QueryInterface( riid, (LPVOID *)ppUnk );
		}
		if (FAILED( hres )) {
			if (hres == E_NOINTERFACE) {
				hres = E_INVALIDARG;
			}
			delete pCOMAreaBrowser;
		}
	}
	return hres;
}



///////////////////////////////////////////////////////////////////////////

//=========================================================================
// ProcessEvents                                                     PUBLIC
// -------------
//    Handles new events for all subscriptions of this server.
//=========================================================================
HRESULT AeComBaseServer::ProcessEvents( DWORD dwNumOfEvents, AeEvent** ppEvents )
{
	// Handle all subscriptions
	m_csSubscrList.Lock();

	for (int i=0; i < m_listSubscr.GetSize(); i++) {
		m_listSubscr[i]->ProcessEvents( dwNumOfEvents, ppEvents );
	}

	m_csSubscrList.Unlock();
	return S_OK;
}



//-------------------------------------------------------------------------
// IMPLEMENTATION
//-------------------------------------------------------------------------

//=========================================================================
// AddSubscriptionToList
// ---------------------
//    Adds an event subscription instance to the subscription list.
//    This method is called when a client creates a new subscription.
//=========================================================================
HRESULT AeComBaseServer::AddSubscriptionToList( AeComSubscriptionManager* pSubscr )
{
	_ASSERTE( pSubscr != NULL );

	HRESULT hres;
	m_csSubscrList.Lock();
	hres = m_listSubscr.Add( pSubscr ) ? S_OK : E_OUTOFMEMORY;
	m_csSubscrList.Unlock();
	return hres;
}



//=========================================================================
// RemoveSubscriptionFromList
// --------------------------
//    Removes an event subscription instance from the subscription list.
//    This method is called when a client has released all interfaces
//    of the event subscription object.
//=========================================================================
HRESULT AeComBaseServer::RemoveSubscriptionFromList( AeComSubscriptionManager* pSubscr )
{
	_ASSERTE( pSubscr != NULL );

	HRESULT hres;
	m_csSubscrList.Lock();
	hres = m_listSubscr.Remove( pSubscr ) ? S_OK : E_FAIL;
	m_csSubscrList.Unlock();
	return hres;
}
//DOM-IGNORE-END


#endif