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
#include "CoreMain.h"
#include "DaComBaseServer.h"
#include "enumclass.h"
#include "DaGenericGroup.h"
#include "openarray.h"
#include "UtilityFuncs.h"
#include "FixOutArray.h"
#include "Logger.h"

///////////////////////////////////////////////////////////////////////////
//////////////////////////// IOPCServer ///////////////////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCServer.AddGroup
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::AddGroup(
	/*[in, string]       */ LPCWSTR     szName,
	/*[in]               */ BOOL        bActive,
	/*[in]               */ DWORD       dwRequestedUpdateRate,
	/*[in]               */ OPCHANDLE   hClientGroup,
	/*[unique, in]       */ LONG      * pTimeBias,
	/*[unique, in]       */ float     * pPercentDeadband,
	/*[in]               */ DWORD       dwLCID,
	/*[out]              */ OPCHANDLE * phServerGroup,
	/*[out]              */ DWORD     * pRevisedUpdateRate,
	/*[in]               */ REFIID      riid,
	/*[out, iid_is(riid)]*/ LPUNKNOWN * ppUnk   )
{
	DaGenericGroup         *theGroup;
	DaGenericGroup         *foundGroup;
	HRESULT               res;
	long                  ServerGroupHandle, sgh;
	LPWSTR                theName;

	bActive = bActive ? TRUE : FALSE;            // Guarantees that parameter bActive is only TRUE and FALSE.
	// Note : Some compilers specifies the value 'true' as nonzero
	//        (e.g. Delphi with value 0xFFFFFFFF).

	if( szName == NULL ) {
		LOGFMTI( "IOPCServer.AddGroup: Unnamed, C.Handle=%ld, Active=%d", hClientGroup, bActive );
	} else {
		OPCWSTOAS( szName )
			LOGFMTI( "IOPCServer.AddGroup '%s', C.Handle=%ld, Active=%d", OPCastr, hClientGroup, bActive );
   		}
	}

	if( phServerGroup != NULL ) {
		*phServerGroup = 0;
	}
	if( pRevisedUpdateRate != NULL ) {
		*pRevisedUpdateRate = 0;
	}
	if( ppUnk != NULL ) {
		*ppUnk = NULL;
	}

	if( ( phServerGroup == NULL ) || ( pRevisedUpdateRate == NULL ) || ( ppUnk == NULL ) ) {

		LOGFMTE("AddGroup() failed with invalid argument(s):" );
		if( phServerGroup == NULL ) {
			LOGFMTE("   phServerGroup is NULL" );
		}
		if( pRevisedUpdateRate == NULL ) {
			LOGFMTE("   pRevisedUpdateRate is NULL" );
		}
		if( ppUnk == NULL ) {
			LOGFMTE("   ppUnk is NULL" );
		}

		return E_INVALIDARG;
	}

	// manipulation of the group lists is critical
	EnterCriticalSection( &(daGenericServer_->m_GroupsCritSec) );

	if( (szName == NULL) || (wcscmp( szName , L"" ) == 0) ) {            
		// no group name specified

		// create unique name ...
		ServerGroupHandle = daGenericServer_->m_GroupList.Size();
		res = daGenericServer_->GetGroupUniqueName( ServerGroupHandle, &theName, FALSE );
		if( FAILED(res) ) {
			WSTRFree( theName, NULL );
			LeaveCriticalSection( &daGenericServer_->m_GroupsCritSec );
			return res;
		}
	}
	else {               // check there is not another one with the same name
		// between all private groups
		theName = WSTRClone( (WCHAR *)szName, NULL );
		if( theName == NULL ) {
			res = E_OUTOFMEMORY;
			WSTRFree( theName, NULL );
			LeaveCriticalSection( &daGenericServer_->m_GroupsCritSec );
			return res;
		}
		res = daGenericServer_->SearchGroup( FALSE, theName, &foundGroup, &sgh );
		if( SUCCEEDED(res) ) {               // found another group with that name!
			res = OPC_E_DUPLICATENAME;
			WSTRFree( theName, NULL );
			LeaveCriticalSection( &daGenericServer_->m_GroupsCritSec );
			return res;
		}
	}

	LeaveCriticalSection( &daGenericServer_->m_GroupsCritSec );

	// construct DaGenericGroup
	theGroup = new DaGenericGroup();
	if( theGroup == NULL ) {
		res = E_OUTOFMEMORY;
		WSTRFree( theName, NULL );
		return res;
	}

	// initialize instance
	res = theGroup->Create(
						   daGenericServer_,
						   theName,
						   bActive,
						   pTimeBias,
						   dwRequestedUpdateRate,
						   hClientGroup,
						   pPercentDeadband,
						   dwLCID,
						   FALSE,
						   0,
						   &ServerGroupHandle
						   );
	if( FAILED(res) ) {
		delete theGroup;
		WSTRFree( theName, NULL );
		return res;
	}

	// get the COM group and the requested interface
	res = daGenericServer_->GetCOMGroup( theGroup, riid, ppUnk );
	if (FAILED(res )) {
		delete theGroup;
		WSTRFree( theName, NULL );
		return res;
	}
	if( pRevisedUpdateRate != NULL ) {
		*pRevisedUpdateRate = theGroup->m_RevisedUpdateRate; 
	}

	*phServerGroup = ServerGroupHandle;

	WSTRFree( theName, NULL );

	// function succeeded
	if (dwRequestedUpdateRate != *pRevisedUpdateRate) {
		return OPC_S_UNSUPPORTEDRATE; 
	}
	return S_OK;

}



//=========================================================================
// IOPCServer.GetErrorString
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::GetErrorString(
	/*[in]         */ HRESULT  dwError,
	/*[in]         */ LCID     dwLocale,
	/*[out, string]*/ LPWSTR * ppString    )
{
	LOGFMTI( "IOPCServer::GetErrorString for error %lX", dwError );

	return OpcCommon::GetErrorString( dwError, dwLocale, ppString );
}



//=========================================================================
// IOPCServer.GetGroupByName
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::GetGroupByName(
	/*[in, string]       */ LPCWSTR szName,
	/*[in]               */ REFIID riid,
	/*[out, iid_is(riid)]*/ LPUNKNOWN * ppUnk    )
{
	HRESULT res;
	long                    ServerGroupHandle;
	DaGenericGroup           *foundGroup;

	USES_CONVERSION;
	LOGFMTI( "IOPCServer.GetGroupByName: '%s'", W2A(szName) );

	if( ppUnk == NULL ) { 
		LOGFMTE("IOPCServer.GetGroupByName failed with invalid argument(s):   ppUnk is NULL");
		return E_INVALIDARG;
	}
	*ppUnk = NULL;                // init result value

	EnterCriticalSection( &daGenericServer_->m_GroupsCritSec );

	// search among generic groups
	res = daGenericServer_->SearchGroup(
		FALSE,
		szName,
		&foundGroup,
		&ServerGroupHandle  );

	if( FAILED(res) ) {    // not found ...
		res = E_INVALIDARG;
		goto GetGroupByNameExit1;
	}

	// create COM Group
	res = daGenericServer_->GetCOMGroup( foundGroup, riid, ppUnk );
	if (FAILED( res )) {
		goto GetGroupByNameExit1;
	}

	LeaveCriticalSection( &daGenericServer_->m_GroupsCritSec );
	return S_OK;

GetGroupByNameExit1:
	LeaveCriticalSection( &daGenericServer_->m_GroupsCritSec );
	return res;
}



//=========================================================================
// IOPCServer::GetStatus
// ---------------------
//    Returns the status of the server.
//=========================================================================
STDMETHODIMP DaComBaseServer::GetStatus( 
									   /*[out]*/   OPCSERVERSTATUS ** ppServerStatus )
{
	OPCSERVERSTATUS*  pServerStatus;
	HRESULT           hres;

	LOGFMTI( "IOPCServer::GetStatus" );

	if (ppServerStatus == NULL) {
		LOGFMTE("GetStatus() failed with invalid argument(s):  ppServerStatus is NULL" );
		return E_INVALIDARG;
	}
	*ppServerStatus = NULL;

	// Allocate memory for the status structure
	pServerStatus = (OPCSERVERSTATUS *)pIMalloc->Alloc( sizeof (OPCSERVERSTATUS) );
	if (pServerStatus == NULL) {
		return E_OUTOFMEMORY;
	}

	// Set the time stamps      
	pServerStatus->ftStartTime       = daGenericServer_->m_StartTime;
	pServerStatus->ftLastUpdateTime  = daGenericServer_->m_LastUpdateTime;
	CoFileTimeNow( &pServerStatus->ftCurrentTime );
	pServerStatus->wReserved = 0;

	LPWSTR pszVendor;

	// Set the number of groups. The functionality is implemented in the
	// automation interface method get_Count():
	hres = get_Count( (long *)&pServerStatus->dwGroupCount );
	if (SUCCEEDED( hres )) {
		// Set the vendor specific status values.
		LOGFMTI( "daBaseServer_->OnGetServerState" );
		hres = daBaseServer_->OnGetServerState(  pServerStatus->dwBandWidth,
			pServerStatus->dwServerState,
			pszVendor);
	}
	if (SUCCEEDED( hres )) {
		// Set the vendor name
		LOGFMTI( "Set the vendor name" );
		if (pszVendor) {
			USES_CONVERSION;                       // Custom interface needs global memory for strings
			pServerStatus->szVendorInfo = WSTRClone( pszVendor, pIMalloc );
			if (pServerStatus->szVendorInfo == NULL) {
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
		LOGFMTI( "Get the version numbers" );
		pServerStatus->wMajorVersion  = core_generic_main.m_VersionInfo.m_wMajor;
		pServerStatus->wMinorVersion  = core_generic_main.m_VersionInfo.m_wMinor;
		pServerStatus->wBuildNumber   = core_generic_main.m_VersionInfo.m_wBuild;

		*ppServerStatus = pServerStatus;          // Attach the server status record to the result pointer.
	}
	else {
		pIMalloc->Free( pServerStatus );          // Something goes wrong
	} 

	LOGFMTI( "IOPCServer::GetStatus finished: 0x%X", hres );
	return hres;
}



//=========================================================================
// IOPCServer.RemoveGroup
// ----------------------
//    Removes a Private Group
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::RemoveGroup(
	/*[in]*/ OPCHANDLE hServerGroup,
	/*[in]*/ BOOL      bForce   )
{
	HRESULT                 hres;
	DaGenericGroup*          pGGroup;
	CComObject<DaGroup>*  pCOMGroup;
	long                    pgh;

	hres = daGenericServer_->GetGenericGroup ( hServerGroup, &pGGroup );
	if (FAILED( hres )) {
		LOGFMTE( "IOPCServer::RemoveGroup: Error: No such group." );
		return hres;
	}
	OPCWSTOAS( pGGroup->m_Name )
		LOGFMTI( "IOPCServer::RemoveGroup %s, Force=%d", OPCastr, bForce ? TRUE : FALSE );
	}

	if (pGGroup->GetPublicInfo( &pgh ) == TRUE ) {
		hres = OPC_E_PUBLIC;
		goto SRemoveGroupExit1;
	}

	hres = daGenericServer_->RemoveGenericGroup( hServerGroup );
	if (SUCCEEDED( hres )) {
		daGenericServer_->CriticalSectionCOMGroupList.BeginReading();          // Lock reading COM list

		daGenericServer_->m_COMGroupList.GetElem( hServerGroup, &pCOMGroup );
		if (pCOMGroup) {                          // There are interface references

			if (bForce) {                          // Also the COM object must be destroyed even
				// if there are interface references

				// Do not really destroy the COM object instance because
				// there are many clients which uses the interfaces anyway.
				//
				// (Also the CTT has problems in this case)
				// while (pCOMGroup->Release());
				//
				//
				daGenericServer_->m_COMGroupList.PutElem( hServerGroup, NULL );

			}
			else {
				hres = OPC_S_INUSE;                 // The COM object still exist
			}
		}
		daGenericServer_->CriticalSectionCOMGroupList.EndReading();        // Unlock reading COM list
	}

	SRemoveGroupExit1:
	daGenericServer_->ReleaseGenericGroup( hServerGroup );
	return hres;
}



//=========================================================================
// IOPCServer.CreateGroupEnumerator
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::CreateGroupEnumerator(
	/*[in]               */ OPCENUMSCOPE dwScope, 
	/*[in]               */ REFIID       riid, 
	/*[out, iid_is(riid)]*/ LPUNKNOWN   *ppUnk   )
{
	HRESULT          hres;
	int              GroupCount;
	LPUNKNOWN       *GroupList;

	LOGFMTI( "IOPCServer::CreateGroupEnumerator" );

	if (ppUnk == NULL) {
		LOGFMTE("CreateGroupEnumerator() failed with invalid argument(s): ppUnk is NULL");
		return E_INVALIDARG;
	}
	*ppUnk = NULL;    // init result value

	if ((riid != IID_IEnumString) && (riid != IID_IEnumUnknown)) {
		LOGFMTE("CreateGroupEnumerator() failed with invalid argument(s): riid must be IID_IEnumString or IID_IEnumUnknown");
		return E_NOINTERFACE;
	}

	// avoid deletion of groups while creating the list
	EnterCriticalSection( &daGenericServer_->m_GroupsCritSec );

	// Get a snapshot of the group list
	// Note: this does NOT do AddRefs to the groups
	hres = daGenericServer_->GetGrpList( riid, dwScope, &GroupList, &GroupCount );

	LeaveCriticalSection( &daGenericServer_->m_GroupsCritSec );

	if( FAILED(hres) ) {
		return hres;
	}

	// Create the Enumerator using the snapshot
	// Note that the enumerator will AddRef the server 
	if( riid == IID_IEnumUnknown) {

		COpcComEnumUnknown *pEnum = new COpcComEnumUnknown;
		if (pEnum == NULL) {
			hres = E_OUTOFMEMORY;
		}
		else {

			hres = pEnum->Init(  GroupList,
				&GroupList[ GroupCount ],
				(IOPCServer *)this,     // This will do an AddRef to the parent
				AtlFlagCopy );

			if (SUCCEEDED( hres )) {
				hres = pEnum ->_InternalQueryInterface( riid, (LPVOID *)ppUnk );
			}

			if (FAILED( hres )) {
				delete pEnum;
			}

		} // if Enum Object created successfully

	} else {

		COpcComEnumString *pEnum = new COpcComEnumString;
		if (pEnum == NULL) {
			hres = E_OUTOFMEMORY;
		}
		else {

			try {
				hres = pEnum->Init(  (LPOLESTR *)GroupList,
					(LPOLESTR *)&GroupList[ GroupCount ],
					(IOPCServer *)this,  // This will do an AddRef to the parent
					AtlFlagCopy );
			}
			catch (HRESULT hresEx) {
				hres = hresEx;
			}

			if (SUCCEEDED( hres )) {
				hres = pEnum ->_InternalQueryInterface( riid, (LPVOID *)ppUnk );
			}

			if (FAILED( hres )) {
				delete pEnum;
			}

		} // if Enum Object created successfully

	} // IEnumString requested

	daGenericServer_->FreeGrpList( riid, GroupList, GroupCount );

	if (SUCCEEDED( hres )) {                     // Return S_FALSE if there is nothing to enumerate
		if (GroupCount == 0) {
			hres = S_FALSE;
		}
	}
	return hres;
}



///////////////////////////////////////////////////////////////////////////
//////////////////////////// IOPCServerPublicGroups ///////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCServerPublicGroups.GetPublicGroupByName
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::GetPublicGroupByName(
	/*[in, string]       */ LPCWSTR     szName,
	/*[in]               */ REFIID      riid,
	/*[out, iid_is(riid)]*/ LPUNKNOWN * ppUnk  )
{
	DaGenericGroup           *theGroup;
	DaPublicGroupManager     *pPublicGroups;
	long                    pgHandle;
	HRESULT                 res;
	long                    pgres;
	DaPublicGroup            *pPG;
	long                    ServerGroupHandle;

	if( ppUnk == NULL ) {
		LOGFMTE("GetPublicGroupByName() failed with invalid argument(s):  ppUnk is NULL");
		return E_INVALIDARG;
	}
	*ppUnk = NULL;                // init result value

	USES_CONVERSION;
	LOGFMTI( "IOPCServerPublicGroups.GetPublicGroupByName '%s'", W2A(szName) );
	pPublicGroups = &daBaseServer_->publicGroups_;

	// handling of the group list is protected by critical section
	EnterCriticalSection( &daGenericServer_->m_GroupsCritSec );

	daGenericServer_->RefreshPublicGroups();

	// search among generic groups
	res = daGenericServer_->SearchGroup( TRUE, szName,
		&theGroup, &ServerGroupHandle  );

	if( FAILED(res) ) {    // not found ...
		res = OPC_E_NOTFOUND;
		goto SPGByNameExit1;
	}

	theGroup->GetPublicInfo( &pgHandle );

	// get and lock it
	pgres = pPublicGroups->GetGroup( pgHandle, &pPG );
	if( FAILED( pgres ) ) {    // public group was deleted 
		res = OPC_E_NOTFOUND;
		goto SPGByNameExit1;
	}

	// create COM Group
	res = daGenericServer_->GetCOMGroup( theGroup, riid, ppUnk );
	if (FAILED( res )) {
		goto SPGByNameExit2;
	}

	pPublicGroups->ReleaseGroup( pgHandle );
	LeaveCriticalSection( &daGenericServer_->m_GroupsCritSec );
	return S_OK;


SPGByNameExit2:
	pPublicGroups->ReleaseGroup( pgHandle );

SPGByNameExit1:
	LeaveCriticalSection( &daGenericServer_->m_GroupsCritSec );
	return res;
}



//=========================================================================
// IOPCServerPublicGroups.RemovePublicGroup
// ----------------------------------------
//    Removes a Public Group
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::RemovePublicGroup(
	/*[in]*/ OPCHANDLE hServerGroup,
	/*[in]*/ BOOL      bForce  )
{
	HRESULT                 res;
	DaGenericGroup           *group;
	long                    pgH;
	CComObject<DaGroup>   *pCOMGroup;

	res = daGenericServer_->GetGenericGroup ( hServerGroup, &group );
	if( FAILED(res)) {
		LOGFMTI("IOPCServerPublicGroups.RemovePublicGroup: Error: No such group." );
		return E_INVALIDARG;                      // required by the CTT in this case
	}
	OPCWSTOAS( group->m_Name )
		LOGFMTI( "IOPCServerPublicGroups.RemovePublicGroup %s, Active=%d", OPCastr, bForce ? TRUE : FALSE );
}

if( group->GetPublicInfo( &pgH ) == FALSE ) {
	daGenericServer_->ReleaseGenericGroup( hServerGroup );
	return E_FAIL;
}

daGenericServer_->ReleaseGenericGroup( hServerGroup );

daBaseServer_->publicGroups_.RemoveGroup( pgH );

// manipulating group list: critical section!
EnterCriticalSection( &daGenericServer_->m_GroupsCritSec );

// removed a public group: some DaGenericGroup may be involved
// so this will lazily kill it!
daGenericServer_->RefreshPublicGroups();

LeaveCriticalSection( &daGenericServer_->m_GroupsCritSec );

res = S_OK;
daGenericServer_->CriticalSectionCOMGroupList.BeginReading();          // Lock reading COM list
daGenericServer_->m_COMGroupList.GetElem( hServerGroup, &pCOMGroup );
if (pCOMGroup) {                          // There are interface references

	if (bForce) {                          // Also the COM object must be destroyed even
		// if there are interface references

		// Do not really destroy the COM object instance because
		// there are many clients which uses the interfaces anyway.
		//
		// (Also the CTT has problems in this case)
		// while (pCOMGroup->Release());
		//
		//
		daGenericServer_->m_COMGroupList.PutElem( hServerGroup, NULL );
	}
	else {
		res = OPC_S_INUSE;                  // The COM object still exist
	}
}
daGenericServer_->CriticalSectionCOMGroupList.EndReading();        // Unock reading COM list
return res;
}



//=========================================================================
// OPC 3.0 Custom Interfaces
//=========================================================================

///////////////////////////////////////////////////////////////////////////
//////////////////////////// IOPCItemIO ///////////////////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCItemIO::Read                                               INTERFACE
// ----------------
//    Reads one or more Values, Qualities and TimeStamps for
//    the specified Items.
//=========================================================================
STDMETHODIMP DaComBaseServer::Read(
								  /* [in] */                    DWORD             dwCount,
								  /* [size_is][in] */           LPCWSTR        *  pszItemIDs,
								  /* [size_is][in] */           DWORD          *  pdwMaxAge,
								  /* [size_is][size_is][out] */ VARIANT        ** ppvValues,
								  /* [size_is][size_is][out] */ WORD           ** ppwQualities,
								  /* [size_is][size_is][out] */ FILETIME       ** ppftTimeStamps,
								  /* [size_is][size_is][out] */ HRESULT        ** ppErrors )
{
	LOGFMTI( "IOPCItemIO::Read" );

	*ppvValues = NULL;                           // Note : Proxy/Stub checks if the pointers are NULL
	*ppwQualities = NULL;
	*ppftTimeStamps = NULL;
	*ppErrors = NULL;

	if (dwCount == 0) {
		LOGFMTE( "Read() failed with invalid argument(s): dwCount is 0" );
		return E_INVALIDARG;
	}

	HRESULT hrRet = S_OK;

	CFixOutArray< VARIANT >    aVal;             // Use global COM memory
	CFixOutArray< WORD >       aQual;
	CFixOutArray< FILETIME >   aTStamps;
	CFixOutArray< HRESULT >    aErr;

	DWORD          i;
	DaDeviceItem**  ppDItems = NULL;

	try {
		// Allocate and initialize the result arrays
		aVal.Init(     dwCount, ppvValues );
		aQual.Init(    dwCount, ppwQualities );
		aTStamps.Init( dwCount, ppftTimeStamps );
		aErr.Init(     dwCount, ppErrors, E_FAIL );

		ppDItems = new DaDeviceItem* [ dwCount ];
		_OPC_CHECK_PTR( ppDItems );
		memset( ppDItems, 0, sizeof (DaDeviceItem*) * dwCount );

		CAutoVectorPtr<OPCITEMDEF> avpItemDefs;
		if (!avpItemDefs.Allocate( dwCount )) throw E_OUTOFMEMORY;

		for (i = 0; i < dwCount; i++) {
			ppDItems[i] = NULL;

			OPCITEMDEF* pItemDef = &avpItemDefs[i];
			pItemDef->szAccessPath = L"";
			pItemDef->szItemID = (LPWSTR)pszItemIDs[i];
			pItemDef->bActive = TRUE;
			pItemDef->hClient = i;
			pItemDef->dwBlobSize = 0;
			pItemDef->pBlob = NULL;
			pItemDef->vtRequestedDataType = VT_EMPTY;
			pItemDef->wReserved = 0;
		}

		hrRet = daBaseServer_->OnValidateItems(
			OPC_VALIDATEREQ_DEVICEITEMS,
			FALSE,                  // No Blob Update
			dwCount,
			avpItemDefs,
			ppDItems,
			NULL,
			*ppErrors );

		_ASSERTE( SUCCEEDED( hrRet ) );           // Must return S_OK or S_FALSE

		avpItemDefs.Free();                       // No longer used

		DWORD dwNumOfItemsWithReadAccess = 0;
		for (i = 0; i < dwCount; i++) {           // Items must have Read Access
			if (ppDItems[i]) {
				if (!ppDItems[i]->HasReadAccess()) {
					ppDItems[i]->Detach();
					ppDItems[i] = NULL;
					aErr[i] = OPC_E_BADRIGHTS;
				}
				else {
					dwNumOfItemsWithReadAccess++;
				}
			}
		}
		// No reason to continue if there are no items with read access
		if (dwNumOfItemsWithReadAccess == 0) throw S_FALSE;

		// Contains the pointers to the Device Items for
		// which the value needs to be updated from the Device.
		CAutoVectorPtr<DaDeviceItem*> avpTmpBufferForItemsToReadFromDevice;
		if (!avpTmpBufferForItemsToReadFromDevice.Allocate( dwCount )) throw E_OUTOFMEMORY;
		memset( avpTmpBufferForItemsToReadFromDevice, NULL, dwCount * sizeof (DaDeviceItem*) );

		FILETIME ftNow;
		HRESULT hr = CoFileTimeNow( &ftNow );
		_OPC_CHECK_HR( hr );

		hrRet = daGenericServer_->InternalReadMaxAge(
			&ftNow,
			dwCount,
			ppDItems,
			avpTmpBufferForItemsToReadFromDevice,
			pdwMaxAge,
			*ppvValues,
			*ppwQualities,
			*ppftTimeStamps,
			*ppErrors );

		_ASSERTE( SUCCEEDED( hrRet ) );           // Must return S_OK or S_FALSE
	}
	catch (HRESULT hrEx) {
		if (hrEx != S_FALSE) {
			aVal.Cleanup();
			aQual.Cleanup();
			aTStamps.Cleanup();
			aErr.Cleanup();
		}
		hrRet = hrEx;
	}

	if (ppDItems) {
		for (i = 0; i < dwCount; i++) {
			if (ppDItems[i]) {
				ppDItems[i]->Detach();
			}
		}
		delete [] ppDItems;
	}

	return hrRet;
}



//=========================================================================
// IOPCItemIO::WriteVQT                                           INTERFACE
// --------------------
//    Writes one or more Values, Qualities and TimeStamps for
//    the specified Items.
//=========================================================================
STDMETHODIMP DaComBaseServer::WriteVQT(
									  /* [in] */                    DWORD             dwCount,
									  /* [size_is][in] */           LPCWSTR        *  pszItemIDs,
									  /* [size_is][in] */           OPCITEMVQT     *  pItemVQT,
									  /* [size_is][size_is][out] */ HRESULT        ** ppErrors )
{
	LOGFMTI( "IOPCItemIO::WriteVQT" );

	*ppErrors = NULL;                            // Note : Proxy/Stub checks if the pointers are NULL

	if (dwCount == 0) {
		LOGFMTE( "WriteVQT() failed with invalid argument(s): dwCount is 0" );
		return E_INVALIDARG;
	}

	HRESULT hrRet = S_OK;

	CFixOutArray< HRESULT > aErr;                // Use global COM memory

	DWORD          i;
	DaDeviceItem**  ppDItems = NULL;

	try {
		aErr.Init( dwCount, ppErrors, S_OK );     // Allocate and initialize the result arrays

		ppDItems = new DaDeviceItem* [ dwCount ];
		_OPC_CHECK_PTR( ppDItems );
		memset( ppDItems, NULL, sizeof (DaDeviceItem*) * dwCount );

		CAutoVectorPtr<OPCITEMDEF> avpItemDefs;
		if (!avpItemDefs.Allocate( dwCount )) throw E_OUTOFMEMORY;

		for (i = 0; i < dwCount; i++) {
			OPCITEMDEF* pItemDef = &avpItemDefs[i];
			pItemDef->szAccessPath = L"";
			pItemDef->szItemID = (LPWSTR)pszItemIDs[i];
			pItemDef->bActive = TRUE;
			pItemDef->hClient = i;
			pItemDef->dwBlobSize = 0;
			pItemDef->pBlob = NULL;
			pItemDef->vtRequestedDataType = V_VT( &pItemVQT[i].vDataValue );
			pItemDef->wReserved = 0;
		}

		hrRet = daBaseServer_->OnValidateItems(
			OPC_VALIDATEREQ_DEVICEITEMS,
			FALSE,                  // No Blob Update
			dwCount,
			avpItemDefs,
			ppDItems,
			NULL,
			*ppErrors );

		_ASSERTE( SUCCEEDED( hrRet ) );           // Must return S_OK or S_FALSE

		avpItemDefs.Free();                       // No longer used

		DWORD dwNumOfItemsWithWriteAccess = 0;
		for (i = 0; i < dwCount; i++) {           // Items must have Write Access
			if (ppDItems[i]) {
				if (!ppDItems[i]->HasWriteAccess()) {
					ppDItems[i]->Detach();
					ppDItems[i] = NULL;
					aErr[i] = OPC_E_BADRIGHTS;
				}
				else {
					dwNumOfItemsWithWriteAccess++;
				}
			}
		}
		// No reason to continue if there are no items with write access
		if (dwNumOfItemsWithWriteAccess == 0) throw S_FALSE;

		hrRet = daGenericServer_->InternalWriteVQT(
			dwCount,
			ppDItems, 
			pItemVQT,
			*ppErrors );

		_ASSERTE( SUCCEEDED( hrRet ) );           // Must return S_OK or S_FALSE
	}
	catch (HRESULT hrEx) {
		if (hrEx != S_FALSE) {
			aErr.Cleanup();
		}
		hrRet = hrEx;
	}

	if (ppDItems) {
		for (i = 0; i < dwCount; i++) {           // Detach all Device Items
			if (ppDItems[i]) {
				ppDItems[i]->Detach();
			}
		}
		delete [] ppDItems;
	}

	return hrRet;
}

///////////////////////////////////////////////////////////////////////////

//DOM-IGNORE-END
