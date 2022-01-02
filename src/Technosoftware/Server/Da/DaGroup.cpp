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

#include "stdafx.h"

#include "DaGroup.h"
#include "UtilityFuncs.h"
#include "DaComServer.h"
#include "DaEnumItemAttributes.h"
#include "DaPublicGroup.h"
#include "DaGenericGroup.h"
#include "FixOutArray.h"
#include "Logger.h"

//=========================================================================
// OPCGroup Constructor
//=========================================================================
DaGroup::DaGroup()
{
	m_pUnkMarshaler = NULL;
	m_pServer = NULL;
	m_pServerHandler = NULL;
	m_ServerGroupHandle = 0;
	m_pGroup = NULL;
}


//=========================================================================
// Create COM Group
// Initialize the Instance
//=========================================================================
HRESULT DaGroup::Create(
							  DaGenericServer *pServer,
							  long           ServerGroupHandle,
							  DaGenericGroup  *pGroup        )
{
	_ASSERTE(( ( pServer != NULL ) && ( ServerGroupHandle != 0 )) );

	pGroup->Attach();
	m_pGroup             = pGroup;

	m_dwCookieGITDataCb  = 0;
	// as long as attached to group
	// server is always available
	m_pServer            = pServer;
	m_pServerHandler     = pServer->m_pServerHandler;
	m_ServerGroupHandle  = ServerGroupHandle;
	return S_OK;
}


//=========================================================================
// OPCGroup Destructor
// delete the COM group from the Server COM list.
//=========================================================================
DaGroup::~DaGroup()
{
	// as long as COM Group attached to Generic Group
	// the Server is still around
	// delete this object from server list of attached COM Group objects
	// if there are interface references

	m_pServer->CriticalSectionCOMGroupList.BeginWriting();             // Lock writing COM list
	m_pServer->m_COMGroupList.PutElem( m_ServerGroupHandle, NULL );
	m_pServer->CriticalSectionCOMGroupList.EndWriting();           // Unlock writing COM list

	m_pGroup->Detach();
}


//=========================================================================
// DaGroup::GetGenericGroup
// Get link to the assotiated DaGenericGroup
// also a Nail is put to the group so that it cannot be deleted
//=========================================================================
HRESULT DaGroup::GetGenericGroup( DaGenericGroup **group )
{
	// get the group through the server for synchronisation
	return m_pServer->GetGenericGroup( m_ServerGroupHandle, group );
}


//=========================================================================
// Disconnect from the assotiated DaGenericGroup
// The Nail set with GetGenericGroup is removed and the group
// is deleted if ToKill is set and no other thread has 
// references to it!
//=========================================================================
HRESULT DaGroup::ReleaseGenericGroup( void )
{
	// release the group through the server for synchronisation
	return m_pServer->ReleaseGenericGroup( m_ServerGroupHandle );
}




//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCGroupStateMgt[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

//=========================================================================
// IOPCGroupStateMgt.GetState
//                                                        
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::GetState(
	/*[out]        */ DWORD     * pUpdateRate, 
	/*[out]        */ BOOL      * pActive,
	/*[out, string]*/ LPWSTR    * ppName,
	/*[out]        */ LONG      * pTimeBias,
	/*[out]        */ float     * pPercentDeadband,
	/*[out]        */ DWORD     * pLCID,
	/*[out]        */ OPCHANDLE * phClientGroup,
	/*[out]        */ OPCHANDLE * phServerGroup   )
{
	DaGenericGroup     *group;
	HRESULT           res;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
		LOGFMTE("IOPCGroupStateMgt.GetState: Internal error: No generic group." );
		return res;
	}
	USES_CONVERSION;
	LOGFMTI("IOPCGroupStateMgt.GetState of group %s", W2A(group->m_Name) );

	// this plausibility check should not be eliminated beacause COM
	// doesn't allow non optional out parameters to be NULL pointers
	if((pUpdateRate == NULL) 
		|| (pActive == NULL) 
		|| (ppName == NULL) 
		|| (pTimeBias == NULL) 
		|| (pPercentDeadband == NULL) 
		|| (pLCID == NULL) 
		|| (phClientGroup == NULL) 
		|| (phServerGroup == NULL) ) {

			LOGFMTE( "GetState() failed with invalid argument(s):" );
			if( pUpdateRate == NULL ) {
				LOGFMTE( "      pUpdateRate is NULL" );
			}
			if( pActive == NULL ) {
				LOGFMTE( "      pActive is NULL" );
			}
			if( ppName == NULL ) {
				LOGFMTE( "      ppName is NULL" );
			}
			if( pTimeBias == NULL ) {
				LOGFMTE( "      pTimeBias is NULL" );
			}
			if( pPercentDeadband == NULL ) {
				LOGFMTE( "      pPercentDeadband is NULL" );
			}
			if( pLCID == NULL ) {
				LOGFMTE( "      pLCID is NULL" );
			}
			if( phClientGroup == NULL ) {
				LOGFMTE( "      phClientGroup is NULL" );
			}
			if( phServerGroup == NULL ) {
				LOGFMTE( "      phServerGroup is NULL" );
			}

			ReleaseGenericGroup();
			return E_INVALIDARG;
	}

	*pUpdateRate      = group->m_RevisedUpdateRate;
	*pActive          = group->GetActiveState();

	EnterCriticalSection( &(m_pServer->m_GroupsCritSec) );
	*ppName           = WSTRClone( group->m_Name, pIMalloc );
	LeaveCriticalSection( &(m_pServer->m_GroupsCritSec) );

	*pTimeBias        = group->m_TimeBias;
	*pPercentDeadband = group->m_PercentDeadband;
	*pLCID            = (DWORD)group->m_dwLCID;
	*phServerGroup    = group->m_hServerGroupHandle;
	*phClientGroup    = group->m_hClientGroupHandle;

	ReleaseGenericGroup();
	return S_OK;
}



//=========================================================================
// IOPCGroupStateMgt.SetState
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::SetState( 
	/*[unique, in]*/ DWORD     * pRequestedUpdateRate, 
	/*[out]       */ DWORD     * pRevisedUpdateRate, 
	/*[unique, in]*/ BOOL      * pActive, 
	/*[unique, in]*/ LONG      * pTimeBias,
	/*[unique, in]*/ float     * pPercentDeadband,
	/*[unique, in]*/ DWORD     * pLCID,
	/*[unique, in]*/ OPCHANDLE * phClientGroup  )
{
	DaGenericGroup     *group;
	HRESULT           res;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
		LOGFMTE("IOPCGroupStateMgt.SetState: Internal error: No generic Group");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI("IOPCGroupStateMgt.SetState of group %s", W2A(group->m_Name) );

	// Percent Deadband
	if (pPercentDeadband) {
		if (*pPercentDeadband < 0.0 || *pPercentDeadband > 100.0) {
			LOGFMTE( "SetState() failed with invalid argument(s): Percent Deadband out of range (0...100)" );
			res = E_INVALIDARG;
			goto SetStateExit1;
		}
		LOGFMTI( "   Percent Deadband: %f%%", *pPercentDeadband );
		// use specified Percent Deadband
		group->m_PercentDeadband = *pPercentDeadband; 
	}

	if (pRequestedUpdateRate) {
		LOGFMTI( "   Requested Update Rate: %lu ms", *pRequestedUpdateRate );
		group->ReviseUpdateRate( *pRequestedUpdateRate );
	}
	if (pRevisedUpdateRate) {
		*pRevisedUpdateRate = group->m_RevisedUpdateRate;
	}

	if (pTimeBias) {
		LOGFMTI( "   Time Zone Bias: %ld minutes", *pTimeBias );
		group->m_TimeBias = *pTimeBias;
	}
	if (pLCID) {
		LOGFMTI( "   Locale ID: 0x%X", *pLCID );
		group->m_dwLCID = *pLCID;
	}
	if (phClientGroup) {
		LOGFMTI( "   Client Handle: %lu (0x%X)", *phClientGroup, *phClientGroup );
		group->m_hClientGroupHandle = *phClientGroup;
	}
	// Note : Some compilers specifies the value 'true' as nonzero
	if (pActive) {
		LOGFMTI( "   Active State: %s", *pActive ? "Active" : "Inactive" );
		group->SetActiveState( *pActive ? TRUE : FALSE );
	}

	ReleaseGenericGroup();

	// function succeeded
	if (pRequestedUpdateRate && (*pRequestedUpdateRate != *pRevisedUpdateRate)) {
		return OPC_S_UNSUPPORTEDRATE; 
	}
	return S_OK;

SetStateExit1:
	ReleaseGenericGroup();
	return res;
}



//=========================================================================
// IOPCGroupStateMgt.SetName
//    Change a name of a private group
//
//    this proc should return error if a group
//    with the given name already exists or a NULL is passed
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::SetName( 
	/*[in, string]*/ LPCWSTR szName  )
{
	DaGenericGroup  *group;
	HRESULT        res;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
		LOGFMTE("IOPCGroupStateMgt.SetName: Internal error: No generic Group");
		return res;
	}

	if (szName) {
		USES_CONVERSION;
		LOGFMTI("IOPCGroupStateMgt.SetName of group %s", W2A(group->m_Name) );
		res = m_pServer->ChangePrivateGroupName( group, (LPWSTR) szName );
	}
	else {
		LOGFMTE( "SetName() failed with invalid argument(s): szName is NULL" );
		res = E_INVALIDARG;
	}
	ReleaseGenericGroup();
	return res;
}




//=========================================================================
// IOPCGroupStateMgt.CloneGroup
//
//    Clones a private or a public group in a private group
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::CloneGroup(
	/*[in, string]       */ LPCWSTR     szName,
	/*[in]               */ REFIID      riid,
	/*[out, iid_is(riid)]*/ LPUNKNOWN * ppUnk  )
{
	DaGenericGroup           *group;
	HRESULT                 res;

	DaGenericGroup           *theGroup;
	DaGenericGroup           *foundGroup;
	long                    ServerGroupHandle, sgh;
	LPWSTR                  theName;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
		LOGFMTE("IOPCGroupStateMgt.CloneGroup: Internal error: No generic Group");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI("IOPCGroupStateMgt.CloneGroup %s", W2A(group->m_Name) );

	if( ppUnk == NULL ) {
		LOGFMTE("CloneGroup() failed with invalid argument(s): ppUnk is NULL");
		ReleaseGenericGroup();
		return E_INVALIDARG;
	}

	*ppUnk = NULL;

	// manipulation of the group lists is critical
	EnterCriticalSection( &(m_pServer->m_GroupsCritSec) );

	if( (szName == NULL) || (wcscmp( szName , L"" ) == 0) ) {            
		// no group name specified

		// create unique name ...
		ServerGroupHandle = m_pServer->m_GroupList.Size();
		res = m_pServer->GetGroupUniqueName( ServerGroupHandle, &theName, FALSE );
		if( FAILED(res) ) {
			goto CloneGroupExit00;
		}
	} else {                
		// check there is not another one with the same name
		// between all private groups
		theName = WSTRClone( (LPWSTR)szName, NULL );
		if( theName == NULL ) {
			res = E_OUTOFMEMORY;
			goto CloneGroupExit00;
		}
		res = m_pServer->SearchGroup( FALSE, theName, &foundGroup, &sgh );
		if( SUCCEEDED(res) ) {               // found another group with that name!
			res = OPC_E_DUPLICATENAME;
			goto CloneGroupExit0;
		}
	}

	// construct DaGenericGroup
	theGroup = new DaGenericGroup();
	if( theGroup == NULL ) {
		res = E_OUTOFMEMORY;
		goto CloneGroupExit0;
	}
	// create
	res = theGroup->Create( group,
		theName,
		&ServerGroupHandle
		);
	if( FAILED( res ) ) {
		delete theGroup;
		goto CloneGroupExit0;
	}

	res = m_pServer->GetCOMGroup( theGroup, riid, ppUnk );
	if (FAILED( res )) {
		goto CloneGroupExit1;
	}

	WSTRFree( theName, NULL );

	LeaveCriticalSection( &m_pServer->m_GroupsCritSec );
	ReleaseGenericGroup();
	return S_OK;

CloneGroupExit1:
	theGroup->Kill();

CloneGroupExit0:
	WSTRFree( theName, NULL );

CloneGroupExit00:
	LeaveCriticalSection( &m_pServer->m_GroupsCritSec );
	ReleaseGenericGroup();
	return res;

}


//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[ IOPCPublicGroupStateMgt[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

//=========================================================================
// IOPCPublicGroupStateMgt.GetState
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::GetState(
	/*[out]*/ BOOL *pPublic  )
{
	DaGenericGroup     *group;
	HRESULT           res;
	long              pgh;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
		LOGFMTE("IOPCPublicGroupStateMgt.GetState: Internal error: No generic Group");
		return E_FAIL;                            // required by the CTT in this case
	}
	USES_CONVERSION;
	LOGFMTI("IOPCPublicGroupStateMgt.GetState of public group %s", W2A(group->m_Name));

	if (pPublic == NULL) {
		LOGFMTE("GetState() failed with invalid argument(s): pPublic is NULL");
		res = E_INVALIDARG;  
	}
	else {
		*pPublic = group->GetPublicInfo( &pgh );
		res = S_OK;
	}

	ReleaseGenericGroup();
	return res;
}


//=========================================================================
// IOPCPublicGroupStateMgtDisp.MoveToPublic and
// IOPCPublicGroupStateMgt.MoveToPublic
//    This method is the implementation for both automation and custom
//    interfaces of MoveToPublic.
//    This method makes a public out of a private group.
//    Because other threads could access the private group while
//    moving it to public it is possible that for a while the
//    private group lives along with the new generated public group
//    until all threads have removed their "nails" from the private group
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::MoveToPublic( void  )
{
	DaGenericGroup     *group;
	HRESULT           res;
	DaPublicGroup      *pubGrp;
	OPCITEMDEF        *ItemDefs;
	OPCITEMDEF        *pID;
	ITEMDEFEXT        *ExtItemDefs = NULL;
	long              ni;
	long              ish, newish;
	DaGenericItem      *item;
	long              PGh;
	long              i;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
		LOGFMTE("IOPCPublicGroupStateMgt.MoveToPublic: Internal error: No generic Group");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI("IOPCPublicGroupStateMgt.MoveToPublic group %s", W2A(group->m_Name));

	if( group->GetPublicInfo( &PGh ) == TRUE ) {
		// must be a private group!
		res = OPC_E_PUBLIC;
		goto PGStateMoveToPublicExit0;
	}

	// no add and no remove allowed during moving
	EnterCriticalSection ( &(group->m_ItemsCritSec) );

	// make itemdefs out of DaGenericItem

	ni = group->m_oaItems.TotElem();
	// create array of itemdefs
	ItemDefs = new OPCITEMDEF[ ni ];
	if( ItemDefs == NULL ) {
		res = E_OUTOFMEMORY;
		goto PGStateMoveToPublicExit1;
	}
	i = 0;
	while ( i < ni ) {
		pID = & ( ItemDefs[i] );

		pID->szAccessPath = NULL;
		pID->szItemID = NULL;
		pID->pBlob = NULL;

		i++;
	}

	ExtItemDefs = new ITEMDEFEXT[ ni ];
	if( ExtItemDefs == NULL ) {
		delete ItemDefs;
		ItemDefs = NULL;
		res = E_OUTOFMEMORY;
		goto PGStateMoveToPublicExit1;
	}


	// get first item handle
	res = group->m_oaItems.First( &ish );
	ni = 0;
	while ( SUCCEEDED(res) ) {
		group->m_oaItems.GetElem( ish, &item ); 
		pID = & ( ItemDefs[ni] );

		// returns a copy of access paths
		item->get_AccessPath( &(pID->szAccessPath) );
		// returns a copy of the ItemID
		item->get_ItemIDCopy( &(pID->szItemID) );
		pID->bActive = item->get_Active();
		pID->hClient = item->get_ClientHandle();
		// returns a copy of blob
		item->get_Blob( &(pID->pBlob), &(pID->dwBlobSize) );
		pID->vtRequestedDataType = item->get_RequestedDataType();

		// get additional item data
		item->get_ExtItemDef( &ExtItemDefs[ni] );

		// increment item counter
		ni ++;
		// get next item handle
		res = group->m_oaItems.Next( ish, &newish );
		ish = newish;
	}

	pubGrp = new DaPublicGroup();
	if( pubGrp == NULL ) {
		res = E_OUTOFMEMORY;
		goto PGStateMoveToPublicExit2;
	}

	// lock the private group name
	// while creating the public group
	EnterCriticalSection ( &(m_pServer->m_GroupsCritSec) );

	res = pubGrp->Create( 
		group->m_Name,
		group->m_PercentDeadband,  // PercentDeadband
		group->m_dwLCID,           // LCID
		group->GetActiveState(),   // Active
		&group->m_TimeBias,        // TimeBias
		group->m_RevisedUpdateRate,// UpdateRate
		ni,                        // Nr. Items
		ItemDefs,                  // ItemDefsArray
		ExtItemDefs                // Additional Item Definitions
		);

	// unlock the group->name_
	LeaveCriticalSection ( &(m_pServer->m_GroupsCritSec) );


	if( FAILED( res ) ) {
		res = E_OUTOFMEMORY;
		goto PGStateMoveToPublicExit3;
	}

	res = m_pServerHandler->publicGroups_.AddGroup( pubGrp, &PGh );
	if( FAILED(res) ) {
		goto PGStateMoveToPublicExit3;
	}

	i = 0;
	while ( i < ni ) {
		pID = & ( ItemDefs[i] );

		if( pID->szAccessPath != NULL ) {
			delete pID->szAccessPath;
		}
		if( pID->szItemID != NULL ) {
			delete pID->szItemID;
		}
		if( pID->pBlob != NULL ) {
			delete pID->pBlob;
		}

		i ++;
	}
	delete [] ItemDefs;
	delete [] ExtItemDefs;

	LeaveCriticalSection ( &(group->m_ItemsCritSec) );

	// Now convert the DaGenericGroup from a Private to a Public Group

	EnterCriticalSection( &(m_pServer->m_GroupsCritSec) );

	group->MakePublic( PGh );
	// group->m_bPublicGroup = TRUE;
	// group->m_hPublicGroupHandle = PGh;

	LeaveCriticalSection( &(m_pServer->m_GroupsCritSec) );

	ReleaseGenericGroup();
	return S_OK;

PGStateMoveToPublicExit3:
	delete pubGrp;

PGStateMoveToPublicExit2:
	i = 0;
	while ( i < ni ) {
		pID = & ( ItemDefs[i] );

		if( pID->szAccessPath != NULL ) {
			delete pID->szAccessPath;
		}
		if( pID->szItemID != NULL ) {
			delete pID->szItemID;
		}
		if( pID->pBlob != NULL ) {
			delete pID->pBlob;
		}

		i ++;
	}
	delete [] ItemDefs;
	if (ExtItemDefs) {
		delete [] ExtItemDefs;
	}

PGStateMoveToPublicExit1:
	LeaveCriticalSection ( &(group->m_ItemsCritSec) );

PGStateMoveToPublicExit0:
	ReleaseGenericGroup();
	return res ;
}



//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCSyncIO [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

//=========================================================================
// IOPCSyncIO.Read
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::Read(
	/*[in]                       */ OPCDATASOURCE   dwSource,
	/*[in]                       */ DWORD           numItems, 
	/*[in, size_is(numItems)]  */ OPCHANDLE     * phServer, 
	/*[out, size_is(,numItems)]*/ OPCITEMSTATE ** ppItemValues,
	/*[out, size_is(,numItems)]*/ HRESULT      ** ppErrors   )
{     
	DaDeviceItem   **ppDItems;
	DaGenericItem   *GItem;
	DWORD           i;
	HRESULT        *pErr, res;
	DaGenericGroup  *group;
	HRESULT         theRes;
	DWORD           AccessRight;
	BOOL            fQualityOutOfService;
	BOOL           *pfItemActiveState = NULL;    // Array with active state of all generic items
	DWORD           dwNumOfDItems = 0;           // Number of successfully attached DeviceItems

	*ppErrors = NULL;                            // Note : Proxy/Stub checks if pointers are NULL
	*ppItemValues = NULL;                           

	res = GetGenericGroup( &group );             // check group state and get ptr
	if( FAILED(res) ) {
		LOGFMTE("IOPCSyncIO.Read: Internal error: No generic Group");
		return res;
	}
	OPCWSTOAS( group->m_Name );
	LOGFMTI( "IOPCSyncIO.Read %ld items of group %s", numItems, OPCastr );
}
// Check the passed arguments
if ((dwSource != OPC_DS_CACHE) && (dwSource != OPC_DS_DEVICE)) {
	LOGFMTE( "SetName() failed with invalid argument(s): dwSource with unknown type" );
	res = E_INVALIDARG;
	goto SIOReadExit0;
}
if(( numItems == 0)         
   || ( phServer == NULL) 
   || ( ppItemValues == NULL) 
   || ( ppErrors == NULL)  ) {            // illegal argument

	   LOGFMTE( "SetName() failed with invalid argument(s):" );
	   if( numItems == 0 ) {
		   LOGFMTE( "      numItems is 0" );
	   }
	   if( phServer == NULL ) {
		   LOGFMTE( "      phServer is NULL " );
	   }
	   if( ppItemValues == NULL ) {
		   LOGFMTE( "      ppItemValues is NULL" );
	   }
	   if( ppErrors == NULL ) {
		   LOGFMTE( "      ppErrors is NULL" );
	   }

	   res = E_INVALIDARG;
	   goto SIOReadExit0;
}

// create errors array
*ppErrors = pErr = (HRESULT*)pIMalloc->Alloc(numItems * sizeof( HRESULT ));
if( *ppErrors == NULL ) {
	res = E_OUTOFMEMORY;      
	goto SIOReadExit0 ;
}
i = 0;
while ( i < numItems ) {
	pErr[ i ] = S_OK;
	i ++;
}
// create itemvalues array
*ppItemValues = (OPCITEMSTATE*)pIMalloc->Alloc(numItems * sizeof( OPCITEMSTATE ));
if(*ppItemValues == NULL ) {
	res = E_OUTOFMEMORY;      
	goto SIOReadExit1;
}

// Generate DeviceItems List 
ppDItems = new DaDeviceItem* [numItems];    // create item array object
if( ppDItems == NULL) {
	res = E_OUTOFMEMORY;
	goto SIOReadExit2;
}
i = 0;
while ( i < numItems ) {
	ppDItems[i] = NULL;
	i ++;
}

// Generate array with active state of all generic items
pfItemActiveState = new BOOL [numItems];
if( pfItemActiveState == NULL ) {
	res = E_OUTOFMEMORY;
	delete [] ppDItems;
	goto SIOReadExit2;
}

for( i=0 ; i < numItems ; i++ ) {
	res = group->GetGenericItem( phServer[i], &GItem );
	if( SUCCEEDED(res) ) {
		if( GItem->AttachDeviceItem( &ppDItems[i] ) > 0 ) {
			// store the active state of the generic item
			pfItemActiveState[i] = GItem->get_Active();
			res = GItem->get_AccessRights( &AccessRight );
			if( SUCCEEDED(res) ) {
				if( (AccessRight & OPC_READABLE) != 0 ) {

					dwNumOfDItems++;

					// allowed to read this item
					(*ppItemValues)[i].hClient = GItem->get_ClientHandle( );
					// set the requested data type!
					V_VT(&(*ppItemValues)[i].vDataValue) = GItem->get_RequestedDataType( );
				} else {
					// cannot read this item
					pErr[i] = OPC_E_BADRIGHTS;     
				}
			} else {
				// invallid access rights 
				pErr[i] = res;     
			}
			if( FAILED( pErr[i] ) ) {
				// this item cannot be read
				ppDItems[i]->Detach();
				ppDItems[i] = NULL;
			}

		}
		else {
			// invalid handle
			// The associated DeviceItem has the killed flag set.
			pErr[i] = OPC_E_UNKNOWNITEMID;
			ppDItems[i] = NULL;  // do not call the read function of the user.
		}
		group->ReleaseGenericItem( phServer[i] );
	} else {
		// There is no generic item.
		pErr[i] = OPC_E_INVALIDHANDLE;                   
	}
}

if (dwNumOfDItems) {
	theRes = group->InternalRead(             // perform the read
		dwSource,
		numItems,
		ppDItems,
		*ppItemValues,
		pErr
		);
}
// Note : theRes is always S_OK or S_FALSE

// if read from cache and the group is inactive then
// all items must return the quality OPC_QUALITY_OUT_OF_SERVICE
if( ( dwSource == OPC_DS_CACHE ) && ( group->GetActiveState() == FALSE ) ) {
	fQualityOutOfService = TRUE;
} else {
	fQualityOutOfService = FALSE;
}
// release the read items!
for( i=0 ; i < numItems ; i++ ) {
	if( FAILED(pErr[i]) ) {
		theRes = S_FALSE;
		V_VT(&(*ppItemValues)[i].vDataValue) = VT_EMPTY;
	} else {
		// Set Quality to 'Bad: Out of Service' if
		// Group or Item is inactive.

		if ( fQualityOutOfService ||                                      // Group is inactive 
			( (dwSource == OPC_DS_CACHE) &&                                // Item is inactive
			(pfItemActiveState[i] == FALSE) ) ) {

				(*ppItemValues)[i].wQuality = 
					OPC_QUALITY_BAD |                                  // new quality is 'bad'
					OPC_QUALITY_OUT_OF_SERVICE |                       // new substatus is 'out of service'
					((*ppItemValues)[i].wQuality & OPC_LIMIT_MASK);    // left the limit unchanged
		}
	}
	if( ppDItems[i] != NULL ) {
		ppDItems[i]->Detach();
	}
}
delete [] ppDItems;
delete [] pfItemActiveState;                                               // release list with active state of generic items
pfItemActiveState = NULL;

ReleaseGenericGroup();
return theRes;

//-------------
SIOReadExit2:
pIMalloc->Free( *ppItemValues );
*ppItemValues = NULL;

SIOReadExit1:
pIMalloc->Free( pErr );
*ppErrors     = NULL;

SIOReadExit0:
ReleaseGenericGroup();

if( pfItemActiveState ) {        // release array with active state of generic items
	delete [] pfItemActiveState;
}
return res;
}



//=========================================================================
// IOPCSyncIO::Write                                              INTERFACE
// -----------------
//    Writes values to one or more items in a group.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::Write(
	/*[in]                       */ DWORD        numItems, 
	/*[in, size_is(numItems)]  */ OPCHANDLE  * phServer, 
	/*[in, size_is(numItems)]  */ VARIANT    * pItemValues, 
	/*[out, size_is(,numItems)]*/ HRESULT   ** ppErrors   )
{
	LOGFMTI( "IOPCSyncIO::Write" );
#ifdef WITH_SHALLOW_COPY
	HRESULT hr = S_OK;
	try {
		CAutoVectorPtr<OPCITEMVQT> avpVQTs;
		if (!avpVQTs.Allocate( numItems )) throw E_OUTOFMEMORY;

		memset( avpVQTs, 0, sizeof (OPCITEMVQT) * numItems );
		for (DWORD i = 0; i < numItems; i++) {
			avpVQTs[i].vDataValue = pItemValues[i];   // Shallow copy
		}

		hr = WriteSync( numItems, phServer, avpVQTs, ppErrors );
	}
	catch (HRESULT hrEx) {  hr = hrEx; }
	catch (...) {           hr = E_FAIL; }
	return hr;
#else
	DWORD    i;
	HRESULT  hr = S_OK;
	CAutoVectorPtr<OPCITEMVQT> avpVQTs;
	try {

		if (!avpVQTs.Allocate( numItems )) throw E_OUTOFMEMORY;

		memset( avpVQTs, 0, sizeof (OPCITEMVQT) * numItems );
		for (i = 0; i < numItems; i++) {
			hr = VariantCopy( &avpVQTs[i].vDataValue, &pItemValues[i] );   // Deep copy
			_OPC_CHECK_HR( hr );
		}

		hr = WriteSync( numItems, phServer, avpVQTs, ppErrors );

	}
	catch (HRESULT hrEx) {  hr = hrEx; }
	catch (...) {           hr = E_FAIL; }

	if (avpVQTs) {
		for (DWORD i = 0; i < numItems; i++) {
			VariantClear( &avpVQTs[i].vDataValue );
		}
	}

	return hr;
#endif
}



//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCAsyncIO [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

//=========================================================================
// IOPCAsyncIO::Read                                              INTERFACE
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::Read(
	/*[in]                       */ DWORD           dwConnection,
	/*[in]                       */ OPCDATASOURCE   dwSource,
	/*[in]                       */ DWORD           numItems,
	/*[in, size_is(numItems)]  */ OPCHANDLE     * phServer,
	/*[out]                      */ DWORD         * pTransactionID,
	/*[out, size_is(,numItems)]*/ HRESULT      ** ppErrors   )
{
	DaDeviceItem   **ppDItems;
	DaGenericItem   *GItem;
	DWORD           i;
	HRESULT        *pErr, res;
	DaGenericGroup  *group;
	DaAsynchronousThread  *pAdv ;
	OPCITEMSTATE   *pItemStates ;
	long           th;
	BOOL           WithTime;
	BOOL           bInvalidHandles;
	BOOL           *pfItemActiveState = NULL;    // List with active state of all generic items

	*pTransactionID = 0;                         // Initialize output parameters
	*ppErrors = NULL;                            // Note : Proxy/Stub checks if pointers are NULL

	res = GetGenericGroup( &group );             // check group state and get ptr
	if( FAILED(res) ) {
		LOGFMTE("IOPCAsyncIO::Read: Internal error: No generic Group");
		return res;
	}
	OPCWSTOAS( group->m_Name );
	LOGFMTI( "IOPCAsyncIO::Read %ld items of group %s", numItems, OPCastr );

}
// Check arguments
if (numItems == 0) {                       // There must be items
	LOGFMTE( "Read() failed with invalid argument(s): numItems is 0" );
	res = E_INVALIDARG;
	goto ASIOReadExit0;
}
// Check the passed arguments
if ((dwSource != OPC_DS_CACHE) && (dwSource != OPC_DS_DEVICE)) {
	LOGFMTE( "Read() failed with invalid argument(s): dwSource with unknown type" );
	res = E_INVALIDARG;
	goto ASIOReadExit0;
}
// Check the connection type
// Note:
// group->m_DataCallback, group->m_DataTimeCallback access
// is not protected by critical section
// because it will be controlled during callback sending
if( dwConnection == STREAM_DATA_CONNECTION ) {
	if( group->m_DataCallback == NULL ) {
		// no advise sink where to send results!
		LOGFMTE( "Read() failed with error: The client has not registered a callback through IDataObject::DAdvise" );
		res = CONNECT_E_NOCONNECTION;
		goto ASIOReadExit0;
	}
	WithTime = FALSE;
} else if( dwConnection == STREAM_DATATIME_CONNECTION ) {
	if( group->m_DataTimeCallback == NULL ) {
		// no advise sink where to send results!
		LOGFMTE( "Read() failed with error: The client has not registered a callback through IDataObject::DAdvise" );
		res = CONNECT_E_NOCONNECTION;
		goto ASIOReadExit0;
	}
	WithTime = TRUE;
} else {
	// invalid connection number
	LOGFMTE( "Read() failed with error: invalid argument : dwConnection is not a connection number returned from IDataObject::DAdvise" );
	res = E_INVALIDARG;
	goto ASIOReadExit0;
}

*ppErrors = pErr = (HRESULT*)pIMalloc->Alloc(numItems * sizeof( HRESULT ));
if( pErr == NULL ) {
	res = E_OUTOFMEMORY;
	goto ASIOReadExit0;
}
// fill with S_OKs
i = 0;
while ( i < numItems ) {
	pErr[ i ] = S_OK;
	i ++;
}

// flag to indicate if there are invalid handles
// used also as condition for freeing memory
bInvalidHandles = FALSE;

// 
pItemStates = new OPCITEMSTATE [numItems];
if (pItemStates == NULL) {
	res = E_OUTOFMEMORY;
	goto ASIOReadExit1;
}
memset( pItemStates, 0, numItems * sizeof( OPCITEMSTATE ) ); // fill with 0s

// create item array object
ppDItems = new DaDeviceItem* [numItems];    
if(ppDItems == NULL) {
	res = E_OUTOFMEMORY;
	goto ASIOReadExit2;
}
// initialize to NULLs
for (i=0; i<numItems; i++) {
	ppDItems[i] = NULL;
}


// Generate list with active state of all generic items
pfItemActiveState = new BOOL [numItems];
if (pfItemActiveState == NULL) {
	res = E_OUTOFMEMORY;
	goto ASIOReadExit3;
}

// check if all handles are valid
i = 0;
while ( i < numItems ) {
	res = group->GetGenericItem( phServer[i], &GItem );
	if( SUCCEEDED(res) ) {
		if( GItem->AttachDeviceItem( &ppDItems[i] ) > 0 ) {
			// store the active state of the generic item
			pfItemActiveState[i] = GItem->get_Active();
			// set client handle into result array
			pItemStates[i].hClient = GItem->get_ClientHandle() ;
			// set requested data type
			V_VT(&pItemStates[i].vDataValue) = GItem->get_RequestedDataType() ;
		} else {
			bInvalidHandles = TRUE;
			pErr[i] = OPC_E_INVALIDHANDLE;
		}
		group->ReleaseGenericItem( phServer[i] );
	} else {
		bInvalidHandles = TRUE;
		pErr[i] = OPC_E_INVALIDHANDLE;
	}
	i ++;
}

if (bInvalidHandles == TRUE) {
	res = S_FALSE;
	goto ASIOReadExit33;
}


// create advice object
pAdv =  new DaAsynchronousThread( group );          
if (pAdv == NULL) {
	goto ASIOReadExit33;
}

// add read advise thread to the list of the group
EnterCriticalSection( &( group->m_AsyncThreadsCritSec ) );
th = group->m_oaAsyncThread.New();
res = group->m_oaAsyncThread.PutElem( th, pAdv );
LeaveCriticalSection( &( group->m_AsyncThreadsCritSec ) );

if (FAILED( res )) {
	goto ASIOReadExit4;
}

// call create thread for READ and REFRESH !
res = pAdv->CreateCustomRead(
							 WithTime,               
							 dwSource,               // from cache or device
							 numItems,             // number of items to handle
							 ppDItems,               // array of DeviceItems to be read
							 pItemStates,            // OPCITEMSTATE result array 
							 pfItemActiveState,      // array with active state of generic items
							 NULL,
							 pTransactionID );       // out parameter

// Note :   Do no longer use pAdr after Create function because object deletes itself.
//          Also all provided arrays will be deleted.
ReleaseGenericGroup();
return res;

//-------------
ASIOReadExit4:
delete pAdv;

ASIOReadExit33:

ASIOReadExit3:
for (i=0; i<numItems; i++) {
	if (ppDItems[i] != NULL) {
		ppDItems[i]->Detach() ;
	}
}
delete [] ppDItems;

ASIOReadExit2:
delete [] pItemStates;

ASIOReadExit1:
if (bInvalidHandles == FALSE) {
	pIMalloc->Free( pErr );
	*ppErrors     = NULL;
}

ASIOReadExit0:
ReleaseGenericGroup();

if (pfItemActiveState) {                     // release list with active state of generic items
	delete [] pfItemActiveState;
}
return res;
}



//=========================================================================
// IOPCAsyncIO::Write                                             INTERFACE
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::Write(
	/*[in]                       */ DWORD        dwConnection,
	/*[in]                       */ DWORD        numItems, 
	/*[in, size_is(numItems)]  */ OPCHANDLE *  phServer,
	/*[in, size_is(numItems)]  */ VARIANT   *  pItemValues, 
	/*[out]                      */ DWORD     *  pTransactionID,
	/*[out, size_is(,numItems)]*/ HRESULT   ** ppErrors  )
{
	DaDeviceItem          **ppDItems;
	DaGenericItem         *GItem;
	DWORD                i;
	HRESULT              *pErr, res;
	DaGenericGroup        *group;
	DaAsynchronousThread        *pAdv ;
	OPCITEMHEADERWRITE   *pItemHdr ;
	long                 th;
	BOOL                 bInvalidHandles;

	*pTransactionID = 0;                         // Initialize output parameters
	*ppErrors = NULL;                            // Note : Proxy/Stub checks if pointers are NULL

	res = GetGenericGroup( &group );             // Check group state and get the pointer
	if (FAILED( res )) {
		LOGFMTE( "IOPCAsyncIO::Write: Internal error: No generic Group" );
		return res;
	}

	OPCWSTOAS( group->m_Name );
	LOGFMTI( "IOPCAsyncIO::Write %ld items of group %s", numItems, OPCastr );
}

// Check arguments
if (numItems == 0) {                       // There must be items
	LOGFMTE( "Write() failed with invalid argument(s): numItems is 0" );
	res = E_INVALIDARG;
	goto ASIOWriteExit0;
}
// Check the connection type
if (dwConnection != STREAM_WRITE_CONNECTION) {
	LOGFMTE( "Write() failed with invalid argument(s): dwConnection is not the connection number returned from IDataObject::DAdvise" );
	res = E_INVALIDARG;
	goto ASIOWriteExit0;
}
// There mus be a registered callback function
// Note: group->m_WriteCallback access is
//       not protected by critical section because 
//       it will be controlled during callback sending.
if (group->m_WriteCallback == NULL) {
	LOGFMTE( "Write() failed with error: The client has not registered a callback through IDataObject::DAdvise" );
	res = CONNECT_E_NOCONNECTION;
	goto ASIOWriteExit0;
}

// flag to indicate if there are invalid handles
// used also as condition for freeing memory
bInvalidHandles = FALSE;

*ppErrors = pErr = (HRESULT*)pIMalloc->Alloc(numItems * sizeof( HRESULT ));
if( *ppErrors == NULL ) {
	res = E_OUTOFMEMORY;
	goto ASIOWriteExit0;
}
i = 0;
while ( i < numItems ) {
	pErr[ i ] = S_OK;
	i ++;
}


pItemHdr = new OPCITEMHEADERWRITE [numItems];
if( pItemHdr == NULL ) {
	res = E_OUTOFMEMORY;
	goto ASIOWriteExit1;
}
memset( pItemHdr, 0, numItems * sizeof( OPCITEMHEADERWRITE ) ); // fill with NULLs

// create item array object
ppDItems = new DaDeviceItem* [numItems];
if(ppDItems == NULL) {
	res = E_OUTOFMEMORY;
	goto ASIOWriteExit2;
}
i = 0;
while ( i < numItems ) {
	ppDItems[i] = NULL;
	i++;
}

// check if all handles are valid
i = 0;
while ( i < numItems ) {
	res = group->GetGenericItem( phServer[i], &GItem );
	if( SUCCEEDED(res) ) {
		if( GItem->AttachDeviceItem( &ppDItems[i] ) > 0 ) { 
			// set client handle into result array
			pItemHdr[i].hClient = GItem->get_ClientHandle() ; 
		} else {
			bInvalidHandles = TRUE;
			pErr[i] = OPC_E_INVALIDHANDLE;
		}
		group->ReleaseGenericItem( phServer[i] );
	} else {
		bInvalidHandles = TRUE;
		pErr[i] = OPC_E_INVALIDHANDLE;
	}
	i ++;
}

if( bInvalidHandles == TRUE ) {
	res = S_FALSE;
	goto ASIOWriteExit4;
}

pAdv =  new DaAsynchronousThread( group );          // create advice object
if( pAdv == NULL ) {
	goto ASIOWriteExit4;
}

// add write advise thread to the list of the group
EnterCriticalSection( &( group->m_AsyncThreadsCritSec ) );
th = group->m_oaAsyncThread.New();
res = group->m_oaAsyncThread.PutElem( th, pAdv );
LeaveCriticalSection( &( group->m_AsyncThreadsCritSec ) );
if( FAILED(res) ) {
	goto ASIOWriteExit5;
}

// call the Create function for a CUSTOM WRITE Callback
res = pAdv->CreateCustomWrite(   numItems,             // number of items to handle
							  ppDItems,               // array of DeviceItems to be read
							  pItemHdr,               // pass partly filled result array
							  pItemValues,            // items values to set
							  NULL,
							  pTransactionID );       // out parameter

// Note :   Do no longer use pAdr after Create function because object deletes itself.
//          Also all provided arrays will be deleted.
ReleaseGenericGroup();
return res;

//-------------
ASIOWriteExit5:
delete pAdv;

ASIOWriteExit4:
for (i=0; i<numItems; i++) {
	if (ppDItems[i] != NULL) {
		ppDItems[i]->Detach() ;
	}
}
delete [] ppDItems;

ASIOWriteExit2:
delete [] pItemHdr;

ASIOWriteExit1:
if (bInvalidHandles == FALSE) {
	pIMalloc->Free( *ppErrors );
	*ppErrors = NULL;
}

ASIOWriteExit0:
ReleaseGenericGroup();
return res;
}



//=========================================================================
// IOPCAsyncIO::Refresh                                           INTERFACE
// 
// Refresh is handled as an Async Read for all active items.
// The UpdateRate generated refresh is not affected by this Refresh
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::Refresh(
	/*[in] */ DWORD           dwConnection,
	/*[in] */ OPCDATASOURCE   dwSource,
	/*[out]*/ DWORD         * pTransactionID  )
{
	DaDeviceItem   **ppDItems;
	DaGenericItem    *GItem;
	DWORD           i; 
	DWORD           nItemsToSend;
	DWORD           nGroupItems;
	long            ih, newih;
	HRESULT         res;
	DaGenericGroup  *group;
	DaAsynchronousThread  *pAdv ;
	OPCITEMSTATE   *pItemStates ;
	long            th;
	BOOL            WithTime;
	BOOL           *pfItemActiveState = NULL;    // Array with active state of all generic items

	*pTransactionID = 0;                         // Initialize output parameter
	// Note : Proxy/Stub checks if pointers are NULL

	res = GetGenericGroup( &group );             // check group state and get ptr
	if( FAILED(res) ) {
		LOGFMTE("IOPCAsyncIO::Refresh: Internal error: No generic Group");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCAsyncIO::Refresh group %s", W2A(group->m_Name));

	// Check the passed arguments
	if ((dwSource != OPC_DS_CACHE) && (dwSource != OPC_DS_DEVICE)) {
		LOGFMTE( "Refresh() failed with invalid argument(s): dwSource with unknown type" );
		res = E_INVALIDARG;
		goto ASIORefreshExit0;
	}

	if( dwConnection == STREAM_DATA_CONNECTION ) {
		if( group->m_DataCallback == NULL ) {
			// no advise sink where to send results!
			LOGFMTE( "Refresh() failed with error: The client has not registered a callback through IDataObject::DAdvise" );
			res = CONNECT_E_NOCONNECTION;
			goto ASIORefreshExit0;
		}
		WithTime = FALSE;
	} else if( dwConnection == STREAM_DATATIME_CONNECTION ) {
		if( group->m_DataTimeCallback == NULL ) {
			// no advise sink where to send results!
			LOGFMTE( "Refresh() failed with error: The client has not registered a callback through IDataObject::DAdvise" );
			res = CONNECT_E_NOCONNECTION;
			goto ASIORefreshExit0;
		}
		WithTime = TRUE;
	} else {
		// invalid connection number
		LOGFMTE( "Refresh() failed with invalid argument(s): dwConnection is not a connection number returned from IDataObject::DAdvise" );
		res = E_INVALIDARG;
		goto ASIORefreshExit0;
	}

	if( ( dwSource == OPC_DS_CACHE ) && ( group->GetActiveState() == FALSE ) ) {
		res = E_FAIL;
		goto ASIORefreshExit0;
	}

	// we do not allow changes to the item list
	// while building arrays of device items, etc.
	EnterCriticalSection( &( group->m_ItemsCritSec) );

	nGroupItems = (DWORD) group->m_oaItems.TotElem();

	pItemStates = new OPCITEMSTATE [ nGroupItems ];
	if(  pItemStates == NULL ) {
		LeaveCriticalSection( &( group->m_ItemsCritSec ) );
		res = E_OUTOFMEMORY;
		goto ASIORefreshExit0;
	}
	i = 0;
	while ( i < nGroupItems ) {
		//VariantInit( &(pItemStates[i].vDataValue) );
		i ++;
	}

	ppDItems = new DaDeviceItem* [ nGroupItems ];    // create item array object
	if(ppDItems == NULL) {
		LeaveCriticalSection( &( group->m_ItemsCritSec ) );
		res = E_OUTOFMEMORY;
		goto ASIORefreshExit1;
	}
	i = 0;
	while ( i < nGroupItems ) {
		ppDItems[i] = NULL;
		i ++;
	}

	// get through the items of a group and
	// and build the arrays taking into account only those
	// which are active
	// !!!??? in theory only those R or RW should be taken into account!?
	nItemsToSend = 0;
	res = group->m_oaItems.First( &ih );
	while ( SUCCEEDED( res ) ) {
		group->m_oaItems.GetElem( ih, &GItem );

		if( GItem->get_Active() == TRUE ) {
			// include the active item
			if( GItem->AttachDeviceItem( &ppDItems[ nItemsToSend ] ) > 0 )  { 
				// DeviceItem attached
				pItemStates[ nItemsToSend ].hClient = GItem->get_ClientHandle() ;
				// set requested data type
				V_VT(&pItemStates[ nItemsToSend ].vDataValue) = GItem->get_RequestedDataType() ;
				nItemsToSend ++;
			}
		}
		res = group->m_oaItems.Next( ih, &newih );
		ih = newih;
	}

	LeaveCriticalSection( &( group->m_ItemsCritSec ) );

	if( nItemsToSend == 0 ) {
		res = E_FAIL;
		goto ASIORefreshExit22;
	}

	// Generate array with active state of all generic items
	pfItemActiveState = new BOOL [nItemsToSend];
	if( pfItemActiveState == NULL ) {
		res = E_OUTOFMEMORY;
		goto ASIORefreshExit22;
	}
	for (i=0; i<nItemsToSend; i++) {             // In functinon refresh() all items found are active
		pfItemActiveState[i] = TRUE;
	}

	// if all there are no active items than fail!

	pAdv =  new DaAsynchronousThread( group );          // create advice object
	if( pAdv == NULL ) {
		goto ASIORefreshExit3;
	}

	// add async read advise thread to the thread list of the group
	EnterCriticalSection( &( group->m_AsyncThreadsCritSec ) );
	th = group->m_oaAsyncThread.New();
	res = group->m_oaAsyncThread.PutElem( th, pAdv );
	LeaveCriticalSection( &( group->m_AsyncThreadsCritSec ) );
	if( FAILED(res) ) {
		// could not add to Thread list
		goto ASIORefreshExit4;
	}

	res = pAdv->CreateCustomRead( WithTime,
		dwSource,               // from cache or device
		nItemsToSend,           // number of items to handle
		ppDItems,               // array of DeviceItems to be read
		pItemStates,            // OPCITEMSTATE result array
		pfItemActiveState,      // array with active state of generic items
		NULL,
		pTransactionID );       // out parameter

	// Note :   Do no longer use pAdr after Create function because object deletes itself.
	//          Also all provided arrays will be deleted.
	ReleaseGenericGroup();
	return res;

	//-------------
ASIORefreshExit4:
	delete pAdv;

ASIORefreshExit3:
	delete [] pfItemActiveState;                 // release list with active state of generic items

ASIORefreshExit22:
	for (i=0; i<nItemsToSend; i++) {
		if (ppDItems[i] != NULL) {
			ppDItems[i]->Detach() ;
		}
	}
	delete [] ppDItems;

ASIORefreshExit1:
	delete [] pItemStates;

ASIORefreshExit0:
	ReleaseGenericGroup();
	return res;
}



//=========================================================================
// IOPCAsyncIO.Cancel
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::Cancel(
	/*[in]*/ DWORD dwTransactionID  )
{
	HRESULT         res;
	DaGenericGroup  *group;
	DaAsynchronousThread  *pAdv ;
	long           i, newi;

	res = GetGenericGroup( &group );       // check group state and get ptr
	if( FAILED(res) ) {
		LOGFMTE("IOPCAsyncIO.Cancel: Internal error: No generic Group");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCAsyncIO.Cancel IO request on group %s", W2A(group->m_Name));

	// check for valid transaction ID
	if( dwTransactionID == 0 ) {
		LOGFMTE( "Cancel() failed with invalid argument(s): dwTransactionID is 0" );
		ReleaseGenericGroup();
		return E_FAIL;
	}

	// look in the different advice objects
	// and check their Transaction ID!
	// Call their Cancel method so that the thread walking over them
	// will kill itself as soon as possible and in the
	// same time kill the advice object
	EnterCriticalSection( &( group->m_AsyncThreadsCritSec ) );

	res = group->m_oaAsyncThread.First( &i );
	while ( SUCCEEDED(res) ) {
		res = group->m_oaAsyncThread.GetElem( i, &pAdv );

		if (pAdv->IsTransactionID( dwTransactionID )) {
			// found the transaction
			pAdv->Cancel();
		}

		res = group->m_oaAsyncThread.Next( i, &newi);
		i = newi;
	}
	LeaveCriticalSection( &( group->m_AsyncThreadsCritSec ) );

	ReleaseGenericGroup();
	// don't know if callback will happen or not!
	return S_FALSE;

}




//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCItemMgt [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[


//=========================================================================
// IOPCItemMgt.AddItems
//
// Add the specified items to the group.
// The items must be succesfully validated bzy the application specific 
// server part ( COpcAppHandler::ValidateItems in App_Server.cpp )
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::AddItems( 
	/*[in]                       */ DWORD           numItems,
	/*[in, size_is(numItems)]  */ OPCITEMDEF     *pItemArray,
	/*[out, size_is(,numItems)]*/ OPCITEMRESULT **ppAddResults,
	/*[out, size_is(,numItems)]*/ HRESULT       **ppErrors  )
{
	return AddItemsInternal( numItems,
		pItemArray,
		ppAddResults,
		ppErrors,
		FALSE );   // It's not a Physical Item Value.
	// Only CALL-R Items can use
	// Items with physical values.
}



//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
// DaGroup Object
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


//=========================================================================
// IOPCItemMgt.AddItemsInternal
//
// Called by AddItems() and AddCallrPyvalItems().
//
// Add the specified items to the group.
// The items must be succesfully validated bzy the application specific 
// server part ( COpcAppHandler::ValidateItems in App_Server.cpp )
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::AddItemsInternal( 
	/*[in]                       */ DWORD           numItems,
	/*[in, size_is(numItems)]  */ OPCITEMDEF     *pItemArray,
	/*[out, size_is(,numItems)]*/ OPCITEMRESULT **ppAddResults,
	/*[out, size_is(,numItems)]*/ HRESULT       **ppErrors,
	BOOL            fPhyValItem )
{
	DaGenericItem   *NewGI;
	DaDeviceItem   **ppDItems;
	HRESULT         res;
	DWORD           i;
	long            hServerHandle;
	DaGenericGroup  *group;
	OPCITEMRESULT  *pRes;
	HRESULT        *pErr;
	OPCITEMDEF     *pID;
	HRESULT        hres;
	long           pgh;

	res = GetGenericGroup( &group );       // check group state and get ptr
	if( FAILED(res) ) {
		LOGFMTE("IOPCItemMgt.AddItems: Internal error: No generic Group");
		return res;
	}
	OPCWSTOAS( group->m_Name );
	LOGFMTI( "IOPCItemMgt.AddItems: Add %ld items to group %s", numItems, OPCastr );
}

if( ppAddResults != NULL ) {
	*ppAddResults = NULL; 
}
if( ppErrors != NULL ) {
	*ppErrors = NULL;
}

if( ( ppAddResults == NULL ) || ( ppErrors == NULL ) || ( numItems == 0 ) 
   || ( pItemArray == NULL ) ) {

	   LOGFMTE( "AddItemsInternal() failed with invalid argument(s):" );
	   if( ppAddResults == NULL ) {
		   LOGFMTE( "      ppAddResults is NULL" );
	   }
	   if( ppErrors == NULL ) {
		   LOGFMTE( "      ppErrors is NULL " );
	   }
	   if( numItems == 0 ) {
		   LOGFMTE( "      numItems is %ld ", numItems );
	   }
	   if( pItemArray == NULL ) {
		   LOGFMTE( "      pItemArray is NULL " );
	   }

	   res = E_INVALIDARG;
	   goto IMAddItemsBadExit0;
}

if( group->GetPublicInfo( &pgh ) == TRUE ) {
	// cannot add items to a public group
	res = OPC_E_PUBLIC;
	goto IMAddItemsBadExit0;
}


// ppAddResults and ppErrors must be created with the global
// COM memory management object
*ppErrors = pErr = (HRESULT*)pIMalloc->Alloc(numItems * sizeof( HRESULT ));
if(pErr == NULL ) {
	res = E_OUTOFMEMORY;
	goto IMAddItemsBadExit0;
}
i = 0;
while ( i < numItems ) {
	pErr[i] = S_OK;
	i ++;
}

// create DeviceItem temp work array were the generated item will be stored
ppDItems = new DaDeviceItem*[ numItems ];
if( ppDItems == NULL ) {
	res = E_OUTOFMEMORY;
	goto IMAddItemsBadExit1;
}
i = 0;
while ( i < numItems ) {
	ppDItems[i] = NULL;
	i ++;
}

// create return array of item results
*ppAddResults = pRes = (OPCITEMRESULT *) pIMalloc->Alloc(numItems * sizeof( OPCITEMRESULT ) );
if( pRes == NULL ) {      
	res = E_OUTOFMEMORY;
	goto IMAddItemsBadExit2;
}
i = 0;
while ( i < numItems ) {
	pRes[i].hServer = 0;
	pRes[i].vtCanonicalDataType = VT_EMPTY;
	pRes[i].wReserved = 0;
	pRes[i].dwAccessRights = 0;
	pRes[i].dwBlobSize = 0;
	pRes[i].pBlob = NULL;
	i ++;
}

// check the item objects
res = m_pServerHandler->OnValidateItems(
										OPC_VALIDATEREQ_DEVICEITEMS,
										TRUE,                         // Blob Update
										numItems,
										pItemArray,
										ppDItems,
										NULL,
										pErr );
if( FAILED(res) ) {
	goto IMAddItemsBadExit3;
}


hres = S_OK;
// insert the successfully validated items in the groups list
// and itemresult datastructures
EnterCriticalSection( &group->m_ItemsCritSec ); // no deadlock!
for( i=0 ; i < numItems ; i++ ) {

	if( SUCCEEDED( pErr[i] ) ) {
		if (ppDItems[i]->Killed()) {
			ppDItems[i]->Detach();              // Attach from OnValidateItems() no longer required
			pErr[i] = OPC_E_UNKNOWNITEMID;
		}
	}
	if( SUCCEEDED( pErr[i] ) ) {
		// Valid Item

		// OPCITEMDEF struct
		pID = &pItemArray[i] ;

		// Create a new GenericItem
		NewGI = new DaGenericItem();
		if( NewGI == NULL ) {
			pErr[i] = E_OUTOFMEMORY ;
			goto IMAddItemsBadExitX;
		}

		res = NewGI->Create( pID->bActive ? TRUE : FALSE,  // active mode flag
			pID->hClient,                 // client handle
			pID->vtRequestedDataType,
			group,
			ppDItems[i],                  // link to DeviceItem
			&hServerHandle,
			fPhyValItem );
		if( FAILED(res) ) {
			delete NewGI;
			pErr[i] = res ;
			goto IMAddItemsBadExitX;
		}

		// fill results into OPCITEMRESULT
		res = ppDItems[i]->get_OPCITEMRESULT( TRUE, &pRes[i] );
		if (FAILED( res )) {
			NewGI->Kill();
			pErr[i] = res ;
			goto IMAddItemsBadExitX;
		}
		pRes[i].hServer = hServerHandle;                   // Use the Server Handle from
		// the Generic Item.

		// used to skip
IMAddItemsBadExitX:
		ppDItems[i]->Detach();              // Attach from OnValidateItems() no longer required
		; // keep this!
	}
	if( FAILED( pErr[i] ) ) {
		hres = S_FALSE;
	}
}
LeaveCriticalSection( &group->m_ItemsCritSec );

delete [] ppDItems;
ReleaseGenericGroup();
return hres;

//IMAddItemsBadExit4:
//   LeaveCriticalSection( &(group->m_ItemsCritSec) );

IMAddItemsBadExit3:
pIMalloc->Free( pRes );
*ppAddResults = NULL;

IMAddItemsBadExit2:
delete [] ppDItems;

IMAddItemsBadExit1:
pIMalloc->Free( pErr );
*ppErrors = NULL;

IMAddItemsBadExit0:
ReleaseGenericGroup();
return res;
}




//=========================================================================
// IOPCItemMgt.ValidateItems
//
// This function checks if some items are supported by the server
//
// The Validation results contain information about
//    Canonical Data Type
//    Blob (optional)
//    DaAccessRights
// 
//    Server Handle is set to 0 ( invalid handle ) because there can be
//    more items with the same ID in a group 
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::ValidateItems( 
	/*[in]                       */ DWORD             numItems,
	/*[in, size_is(numItems)]  */ OPCITEMDEF      * pItemArray,
	/*[in]                       */ BOOL              bBlobUpdate,
	/*[out, size_is(,numItems)]*/ OPCITEMRESULT  ** ppValidationResults,
	/*[out, size_is(,numItems)]*/ HRESULT        ** ppErrors  )
{
	HRESULT        hres;
	DWORD          i;
	DaGenericGroup  *pGGroup;
	OPCITEMRESULT  *pRes;
	HRESULT        *pErr;
	long           pgh;

	hres = GetGenericGroup( &pGGroup );          // Check group state and get ptr
	if (FAILED( hres )) {
		LOGFMTE("IOPCItemMgt.ValidateItems: Internal error: No generic Group");
		return hres;
	}
	OPCWSTOAS( pGGroup->m_Name );
	LOGFMTI( "IOPCItemMgt.ValidateItems: Validate %ld items in group %s", numItems, OPCastr );
}

*ppValidationResults = NULL;                 // Note : Proxy/Stub checks if the pointers are NULL
*ppErrors = NULL;

if (numItems == 0) {
	LOGFMTE( "ValidateItems() failed with invalid argument(s): numItems is 0" );
	hres = E_INVALIDARG;
	goto IMValidateItemsBadExit0;
}

if (pGGroup->GetPublicInfo( &pgh ) == TRUE) {
	hres = OPC_E_PUBLIC;                      // Cannot validate add items to a public group
	goto IMValidateItemsBadExit0;
}
// ppValidationResults and ppErrors must be created
pErr = ComAlloc<HRESULT>( numItems );      // with the global COM memory management object
if (!pErr) {
	hres = E_OUTOFMEMORY;
	goto IMValidateItemsBadExit0;
}
pRes = ComAlloc<OPCITEMRESULT>( numItems );// Create return array of item results
if (!pRes) {      
	hres = E_OUTOFMEMORY;
	goto IMValidateItemsBadExit1;
}
// Initialize the arrays
for (i=0; i < numItems; i++) {
	pErr[i]                       = S_OK;

	pRes[i].hServer               = 0;
	pRes[i].vtCanonicalDataType   = VT_EMPTY;
	pRes[i].wReserved             = 0;
	pRes[i].dwAccessRights        = 0;
	pRes[i].dwBlobSize            = 0;
	pRes[i].pBlob                 = NULL;
}
// Check the item objects
hres = m_pServerHandler->OnValidateItems(
	OPC_VALIDATEREQ_ITEMRESULTS,
	bBlobUpdate,
	numItems,
	pItemArray,
	NULL,
	pRes,
	pErr );

if (FAILED( hres )) {
	goto IMValidateItemsBadExit2;
}

*ppErrors = pErr;                            // All succeeded
*ppValidationResults = pRes;

ReleaseGenericGroup();
return hres;

IMValidateItemsBadExit2:
pIMalloc->Free( pRes );

IMValidateItemsBadExit1:
pIMalloc->Free( pErr );

IMValidateItemsBadExit0:
ReleaseGenericGroup();
return hres;
}



//=========================================================================
// IOPCItemMgt.RemoveItems
//
// Items are removed even though some operation is performed over them!
// So the client must care to avoid access to them while removing
// Both CGeneric- and COM-Items are removed! 
// 
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::RemoveItems( 
	/*[in]                       */ DWORD        numItems,
	/*[in, size_is(numItems)]  */ OPCHANDLE  * phServer,
	/*[out, size_is(,numItems)]*/ HRESULT   ** ppErrors  )
{
	OPCHANDLE      sh;
	HRESULT       *pErr, res;
	BOOL           SomeBadHandle;
	DWORD          i;
	DaGenericGroup  *group;
	long           pgh;

	res = GetGenericGroup( &group );       // check group state and get ptr
	if( FAILED(res) ) {
		LOGFMTE("IOPCItemMgt.RemoveItems: Internal error: No generic Group");
		return res;
	}
	OPCWSTOAS( group->m_Name );
	LOGFMTI( "IOPCItemMgt.RemoveItems: Remove %ld items from group %s", numItems, OPCastr );
}

if( ppErrors != NULL ) {
	*ppErrors = NULL;
}

if( ( ppErrors == NULL) || ( numItems == 0 ) || ( phServer == NULL ) ) {

	LOGFMTE( "RemoveItems() failed with invalid argument(s):" );
	if( ppErrors == NULL ) {
		LOGFMTE( "      ppErrors is NULL " );
	}
	if( numItems == 0 ) {
		LOGFMTE( "      numItems is %ld ", numItems );
	}
	if( phServer == NULL ) {
		LOGFMTE( "      phServer is NULL " );
	}

	res = E_INVALIDARG;
	goto IMRemoveItemsExit0;
}

if( group->GetPublicInfo( &pgh ) == TRUE ) {       // is public group !
	res = OPC_E_PUBLIC;
	goto IMRemoveItemsExit0;
}
// alloc result array
*ppErrors = pErr = (HRESULT*)pIMalloc->Alloc(numItems * sizeof( HRESULT ));
if( *ppErrors == NULL ) {
	res = E_OUTOFMEMORY;
	goto IMRemoveItemsExit0;
}
i = 0;
while ( i < numItems ) {
	pErr[i] = S_OK;
	i ++;
}

EnterCriticalSection( &(group->m_ItemsCritSec) ); 

SomeBadHandle = FALSE;
for( i=0 ; i < numItems ; i++ ) {                // Item loop
	sh = phServer[ i ];                // get the server handle from input
	res = group->RemoveGenericItemNoLock( sh );     // kill GenericItem and 
	// detach from DeviceItem
	if( FAILED(res) ) {
		SomeBadHandle = TRUE;
	}
	pErr[i] = res ;
}

LeaveCriticalSection( &(group->m_ItemsCritSec) );

if( SomeBadHandle == TRUE ) {         // at least one item failed
	res = S_FALSE;
} else {
	res = S_OK;
}

IMRemoveItemsExit0:
ReleaseGenericGroup();
return res;
}




//=========================================================================
// IOPCItemMgt.SetActiveState
//
// 
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::SetActiveState(
	/*[in]                       */ DWORD        numItems,
	/*[in, size_is(numItems)]  */ OPCHANDLE  * phServer,
	/*[in]                       */ BOOL         bActive, 
	/*[out, size_is(,numItems)]*/ HRESULT   ** ppErrors )
{
	OPCHANDLE               sh;
	HRESULT                *pErr, res;
	DaGenericItem            *pItem;
	BOOL                    SomeBadHandle;
	DWORD                   i;
	DaGenericGroup           *group;

	res = GetGenericGroup( &group );       // check group state and get ptr
	if( FAILED(res) ) {
		LOGFMTE("IOPCItemMgt.SetActiveState: Internal error: No generic Group");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCItemMgt.SetActiveState of group %s", W2A(group->m_Name));

	if( ppErrors != NULL ) {
		*ppErrors = NULL;
	}

	if( ( ppErrors == NULL ) || ( phServer == NULL ) || ( numItems == 0 ) ) {

		LOGFMTE( "SetActiveState() failed with invalid argument(s):" );
		if( ppErrors == NULL ) {
			LOGFMTE( "      ppErrors is NULL " );
		}
		if( numItems == 0 ) {
			LOGFMTE( "      numItems is %ld ", numItems );
		}
		if( phServer == NULL ) {
			LOGFMTE( "      phServer is NULL " );
		}

		res = E_INVALIDARG;
		goto IMSetActiveStateExit1;
	}

	// alloc result array
	*ppErrors = pErr = (HRESULT*)pIMalloc->Alloc(numItems * sizeof( HRESULT ));
	if(*ppErrors == NULL ) {
		goto IMSetActiveStateExit1;
	}
	i = 0;
	while ( i < numItems ) {
		pErr[i] = S_OK;
		i ++;
	}

	SomeBadHandle = FALSE;
	for( i=0 ; i < numItems ;i++ ) {
		sh = phServer[ i ];                 // get the server handle from input

		// get and lock the associated item
		res = group->GetGenericItem( sh, &pItem );
		if( SUCCEEDED(res) ) {
			res = pItem->set_Active( bActive ? TRUE : FALSE );
			group->ReleaseGenericItem( sh );
		}
		if( FAILED(res) ) {                 // item failed
			SomeBadHandle = TRUE;
		}
		pErr[i] = res ;
	}

	if( SomeBadHandle == TRUE ) {         // at least one item failed
		res = S_FALSE;
	} else {
		res = S_OK;
	}

IMSetActiveStateExit1:
	ReleaseGenericGroup();
	return res;
}




//=========================================================================
// IOPCItemMgt.SetClientHandles
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::SetClientHandles(
	/*[in]                       */ DWORD        numItems,
	/*[in, size_is(numItems)]  */ OPCHANDLE  * phServer,
	/*[in, size_is(numItems)]  */ OPCHANDLE  * phClient,
	/*[out, size_is(,numItems)]*/ HRESULT   ** ppErrors  )
{
	OPCHANDLE               sh;
	OPCHANDLE               ch;
	HRESULT                *pErr, res;
	DaGenericItem            *pItem;
	BOOL                    SomeBadHandle;
	DWORD                   i;
	DaGenericGroup           *group;

	res = GetGenericGroup( &group );       // check group state and get ptr
	if( FAILED(res) ) {
		LOGFMTE("IOPCItemMgt.SetClientHandles: Internal error: No generic Group");
		return res ;
	}
	OPCWSTOAS( group->m_Name );
	LOGFMTI( "IOPCItemMgt.SetClientHandles for %ld items in group %s", numItems, OPCastr );
}

if( ppErrors != NULL ) {
	*ppErrors = NULL;
}

if( ( ppErrors == NULL) || ( phServer == NULL ) 
   || ( phClient == NULL ) || ( numItems == 0 ) ) {

	   LOGFMTE( "SetClientHandles() failed with invalid argument(s):" );
	   if( ppErrors == NULL ) {
		   LOGFMTE( "      ppErrors is NULL " );
	   }
	   if( numItems == 0 ) {
		   LOGFMTE( "      numItems is %ld ", numItems );
	   }
	   if( phServer == NULL ) {
		   LOGFMTE( "      phServer is NULL " );
	   }
	   if( phClient == NULL ) {
		   LOGFMTE( "      phClient is NULL " );
	   }

	   res = E_INVALIDARG;
	   goto IMSetClientHandlesBadExit1;
}

*ppErrors = pErr = (HRESULT*)pIMalloc->Alloc(numItems * sizeof( HRESULT ));
if(*ppErrors == NULL ) {
	res = E_OUTOFMEMORY;
	goto IMSetClientHandlesBadExit1;
}
i = 0;
while ( i < numItems ) {
	pErr[i] = S_OK;
	i ++;
}

SomeBadHandle = FALSE;
i = 0;
while ( i < numItems ) {
	// get the server handle from input
	sh = phServer[ i ];
	ch = phClient[ i ];

	// get the associated item
	res = group->GetGenericItem( sh, &pItem );
	if( SUCCEEDED(res)) {
		pItem->set_ClientHandle( ch );
		group->ReleaseGenericItem( sh );
	} 
	if( FAILED(res) ) {
		SomeBadHandle = TRUE;
	}

	pErr[i] = res ;
	i++;
}

if( SomeBadHandle == TRUE ) {
	res = S_FALSE;
} else {
	res = S_OK;
}

IMSetClientHandlesBadExit1:
ReleaseGenericGroup();
return res;
}




//=========================================================================
// IOPCItemMgt.SetDatatypes
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::SetDatatypes(
	/*[in]                       */ DWORD        numItems,
	/*[in, size_is(numItems)]  */ OPCHANDLE  * phServer,
	/*[in, size_is(numItems)]  */ VARTYPE    * pRequestedDatatypes,
	/*[out, size_is(,numItems)]*/ HRESULT   ** ppErrors  )
{
	OPCHANDLE               sh;
	VARTYPE                 rdt;
	HRESULT                *pErr, res;
	DaGenericItem            *pItem;
	BOOL                    SomeBadHandle;
	DWORD                   i;
	DaGenericGroup           *group;

	res = GetGenericGroup( &group );       // check group state and get ptr
	if( FAILED(res) ) {
		LOGFMTE("IOPCItemMgt.SetDataTypes: Internal error: No generic Group");
		return res ;
	}
	OPCWSTOAS( group->m_Name );
	LOGFMTI( "IOPCItemMgt.SetDataTypes for %ld items in group %s", numItems, OPCastr );
}

if( ppErrors != NULL ) {
	*ppErrors = NULL;
}

if( ( ppErrors == NULL) 
   ||  ( phServer == NULL ) 
   ||  ( pRequestedDatatypes == NULL ) 
   ||  ( numItems == 0 ) ) {

	   LOGFMTE( "SetDatatypes() failed with invalid argument(s):" );
	   if( ppErrors == NULL ) {
		   LOGFMTE( "      ppErrors is NULL " );
	   }
	   if( numItems == 0 ) {
		   LOGFMTE( "      numItems is %ld ", numItems );
	   }
	   if( phServer == NULL ) {
		   LOGFMTE( "      phServer is NULL " );
	   }
	   if( pRequestedDatatypes == NULL ) {
		   LOGFMTE( "      pRequestedDatatypes is NULL " );
	   }

	   res = E_INVALIDARG;
	   goto IMSetRequestedDatatypesExit1;
}
// alloc result array
pErr = *ppErrors = (HRESULT*)pIMalloc->Alloc(numItems * sizeof( HRESULT ));
if(*ppErrors == NULL ) {
	res = E_OUTOFMEMORY;
	goto IMSetRequestedDatatypesExit1;
}

SomeBadHandle = FALSE;
for( i=0 ; i < numItems ; i++) {
	sh = phServer[ i ];                 // get the server handle from input
	rdt = pRequestedDatatypes[ i ];
	res = group->GetGenericItem( sh, &pItem );   // check item and get ptr
	if( SUCCEEDED(res)) {
		res = pItem->set_RequestedDataType( rdt );   // now set the type
		group->ReleaseGenericItem( sh );
	} 
	if( FAILED(res) ) {
		SomeBadHandle = TRUE;
	}
	pErr[i] = res;
}

if( SomeBadHandle == TRUE ) {         // at least one item failed
	res = S_FALSE;
} else {
	res = S_OK;
}

IMSetRequestedDatatypesExit1:
ReleaseGenericGroup();
return res;
}




//=========================================================================
// IOPCItemMgt.CreateEnumerator
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::CreateEnumerator(
	/*[in]               */ REFIID      riid,
	/*[out, iid_is(riid)]*/ LPUNKNOWN * ppUnk  )
{
	DaEnumItemAttributes* temp;
	long                    ItemCount;
	OPCITEMATTRIBUTES*      ItemList;
	HRESULT                 res;
	DaGenericGroup*          group;

	LOGFMTI( "IOPCItemMgt::CreateEnumerator" );

	if (ppUnk != NULL) {
		*ppUnk = NULL;
	}

	if ((ppUnk == NULL) || (riid != IID_IEnumOPCItemAttributes)) {
		LOGFMTE( "CreateEnumerator() failed with invalid argument(s):" );
		if (ppUnk == NULL) {
			LOGFMTE( "      ppUnk is NULL " );
		}
		if (riid != IID_IEnumOPCItemAttributes) {
			LOGFMTE( "      riid is not IID_IEnumOPCItemAttributes" );
		}
		return E_INVALIDARG;
	}

	res = GetGenericGroup( &group );             // Check group state and get object ptr
	if( FAILED(res) ) {
		return res;
	}

	EnterCriticalSection( &group->m_ItemsCritSec );

	// Get a snapshot of the group list 
	// Note: this does NOT do AddRefs to the groups
	res = GetItemAttrList( group, &ItemList, &ItemCount);

	LeaveCriticalSection( &group->m_ItemsCritSec );
	ReleaseGenericGroup();

	if (FAILED( res )) {
		return res;
	}
	// Create the Enumerator using the snapshot
	// Note that the enumerator will AddRef the server 
	// and also all of the groups.
	temp = new DaEnumItemAttributes(  (IOPCGroupStateMgt*)this,
		ItemCount,
		ItemList,
		NULL
		);

	FreeItemAttrList( ItemList, ItemCount );

	if (temp == NULL) { 
		return E_OUTOFMEMORY;
	}

	res = temp->QueryInterface( riid, (LPVOID*)ppUnk );
	if (FAILED( res )) {
		delete temp;
	}
	else {                                       // SUCCEEDED
		if (ItemCount == 0) {
			res = S_FALSE;                         // There is nothing to enumerate.
		}
	}
	return res;
}





//=========================================================================
// DaGroup::GetItemAttrList
//
// returns the items of this group as an array of OPCITEMATTRIBUTES
//=========================================================================
HRESULT DaGroup::GetItemAttrList( 
									   DaGenericGroup      *group, 
									   OPCITEMATTRIBUTES **ItemList, 
									   long               *ItemCount )
{
	HRESULT        res;
	long           idx, newidx, i;
	DaGenericItem   *GI;

	*ItemCount = group->m_oaItems.TotElem();
	*ItemList = new OPCITEMATTRIBUTES [ *ItemCount ];
	if (*ItemList == NULL) {
		return E_OUTOFMEMORY;
	}

	res = group->m_oaItems.First( &idx );
	i = 0;
	while ( SUCCEEDED(res) ) {
		group->m_oaItems.GetElem( idx, &GI );
		if( GI->Killed() == FALSE ) {
			GI->get_Attr( &((*ItemList)[i]), idx );
			i++;
		}
		res = group->m_oaItems.Next( idx, &newidx );
		idx = newidx;
	}
	*ItemCount = i;                     // don't count killed items
	return S_OK;
}

//=========================================================================
// DaGroup::FreeItemAttrList
//
// frees array of OPCITEMATTRIBUTES allocated in GetItemAttrList
//=========================================================================
HRESULT DaGroup::FreeItemAttrList( OPCITEMATTRIBUTES *ItemList, long ItemCount )
{
	long i;
	OPCITEMATTRIBUTES  *IA;


	if( ItemList == NULL ) {
		return S_OK;
	}

	i = 0;
	while ( i < ItemCount ) {
		IA = &(ItemList[i]);

		if( IA->szAccessPath != NULL ) {
			delete IA->szAccessPath;
		}
		if( IA->szItemID != NULL ) {
			delete IA->szItemID;
		};
		if( IA->pBlob != NULL ) {
			delete IA->pBlob;
		}
		VariantClear( &IA->vEUInfo );

		i ++;
	}

	delete [] ItemList;

	return S_OK;
}



//=========================================================================
// DaGroup::GetCOMItemList
//
// Inputs:
//    group: must be nailed outside and 
//           this method should be called inside it's ItemsCritSec 
// Outputs:
//    returns the items of this group as an array of OPCITEMATTRIBUTES
//=========================================================================
HRESULT DaGroup::GetCOMItemList( DaGenericGroup *group, LPUNKNOWN **ItemList, long *ItemCount )
{

	long size;
	long sh,nsh;
	DaItem    *theCOMItem;
	BOOL         bCOMItemCreated;
	LPUNKNOWN   pUnk;
	long        count;
	HRESULT     res;
	DaGenericItem   *item;


	// size of array to create
	size = group->m_oaItems.TotElem();                   
	if( size == 0 ) {  
		// no items
		res = S_OK;
		goto GetCOMItemListExit0;
	}

	*ItemList = new LPUNKNOWN[size];
	if( *ItemList == NULL ) {
		res = E_OUTOFMEMORY;
		goto GetCOMItemListExit0;
	}

	count = 0;

	res = group->m_oaItems.First( &sh );
	while ( SUCCEEDED( res ) ) {
		// get the item
		res = group->m_oaItems.GetElem( sh , &item );
		if( ( SUCCEEDED(res)) && ( item->Killed() == FALSE )  ) {

			// create a new COM item for this group
			res = group->GetCOMItem( sh, item, &theCOMItem, &bCOMItemCreated );
			if( FAILED(res) ) {      
				goto GetCOMItemListExit1;
			}
			// get a new interface for this item
			res = theCOMItem->QueryInterface( IID_IUnknown, ( LPVOID* )&pUnk);
			if( FAILED(res) ) {    
				// could not get an interface
				if( bCOMItemCreated == TRUE ) {
					// as long as there won't be a release
					// we must delete the just created item
					delete theCOMItem;
				}
				goto GetCOMItemListExit1;
			}

			(*ItemList)[count] = pUnk ;
			++count;                         
		}
		res = group->m_oaItems.Next( sh, &nsh );
		sh = nsh;
	}  

	*ItemCount = count;

	return S_OK;

GetCOMItemListExit1:
	// release the COM items
	count--;
	while ( count >= 0 ) {
		(*ItemList)[ count ]->Release();
		count--;
	}
	delete ItemList;

GetCOMItemListExit0:
	*ItemList = NULL;
	*ItemCount = 0;
	return res;
}


//=========================================================================
// DaGroup::FreeCOMItemList
//
// frees array allocated in GetCOMItemList
//=========================================================================
HRESULT DaGroup::FreeCOMItemList( LPUNKNOWN *ItemList, long ItemCount )
{
	long i;

	i = 0;
	while ( i < ItemCount ) {
		ItemList[i]->Release();
		i ++;
	}
	delete ItemList;
	return S_OK;
}


//=========================================================================
// OPC 3.0 Custom Interfaces
//=========================================================================

///////////////////////////////////////////////////////////////////////////
//////////////////////////// IOPCGroupStateMgt2 ///////////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCGroupStateMgt2::SetKeepAlive                               INTERFACE
// --------------------------------
//    Sets the keep-alive time for the subscription of this group.
//=========================================================================
STDMETHODIMP DaGroup::SetKeepAlive(
	/* [in] */                    DWORD             dwKeepAliveTime,
	/* [out] */                   DWORD          *  pdwRevisedKeepAliveTime )
{
	LOGFMTI( "IOPCGroupStateMgt2::SetKeepAlive" );

	HRESULT        hr;
	DaGenericGroup* pGGroup;

	hr = GetGenericGroup( &pGGroup );         // Check group state and get the pointer
	if (FAILED( hr )) {
		LOGFMTE( "IOPCGroupStateMgt2::SetKeepAlive: Internal error: No generic Group" );
		return hr;
	}

	hr = pGGroup->SetKeepAlive( dwKeepAliveTime, pdwRevisedKeepAliveTime );

	ReleaseGenericGroup();
	return hr;
}



//=========================================================================
// IOPCGroupStateMgt2::GetKeepAlive                               INTERFACE
// --------------------------------
//    Returns the currently active keep-alive time for the subscription
//    of this group.
//=========================================================================
STDMETHODIMP DaGroup::GetKeepAlive(
	/* [out] */                   DWORD          *  pdwKeepAliveTime )
{
	LOGFMTI( "IOPCGroupStateMgt2::GetKeepAlive" );

	HRESULT        hr;
	DaGenericGroup* pGGroup;

	hr = GetGenericGroup( &pGGroup );         // Check group state and get the pointer
	if (FAILED( hr )) {
		LOGFMTE( "IOPCGroupStateMgt2::GetKeepAlive: Internal error: No generic Group" );
		return hr;
	}

	*pdwKeepAliveTime = pGGroup->KeepAliveTime();

	ReleaseGenericGroup();
	return hr;
}



///////////////////////////////////////////////////////////////////////////
//////////////////////////// IOPCSyncIO2 //////////////////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCSyncIO2::ReadMaxAge                                        INTERFACE
// -----------------------
//    Reads one or more Values, Qualities and TimeStamps for
//    the specified Items.
//=========================================================================
STDMETHODIMP DaGroup::ReadMaxAge(
									   /* [in] */                    DWORD             dwCount,
									   /* [size_is][in] */           OPCHANDLE      *  phServer,
									   /* [size_is][in] */           DWORD          *  pdwMaxAge,
									   /* [size_is][size_is][out] */ VARIANT        ** ppvValues,
									   /* [size_is][size_is][out] */ WORD           ** ppwQualities,
									   /* [size_is][size_is][out] */ FILETIME       ** ppftTimeStamps,
									   /* [size_is][size_is][out] */ HRESULT        ** ppErrors )
{
	LOGFMTI( "IOPCSyncIO2::ReadMaxAge" );

	*ppvValues = NULL;                           // Note : Proxy/Stub checks if the pointers are NULL
	*ppwQualities = NULL;
	*ppftTimeStamps = NULL;
	*ppErrors = NULL;

	if (dwCount == 0) {
		LOGFMTE( "ReadMaxAge() failed with invalid argument(s): dwCount is 0" );
		return E_INVALIDARG;
	}

	DaGenericGroup* pGGroup;
	HRESULT hrRet = GetGenericGroup( &pGGroup ); // Check group state and get the pointer
	if (FAILED( hrRet )) {
		LOGFMTE( "IOPCSyncIO2::ReadMaxAge: Internal error: No generic Group" );
		return hrRet;
	}

	CFixOutArray< VARIANT >    aVal;             // Use global COM memory
	CFixOutArray< WORD >       aQual;
	CFixOutArray< FILETIME >   aTStamps;

	DWORD          i;
	DWORD          dwCountValid = 0;             // Number of valid items
	DaDeviceItem**  ppDItems = NULL;              // Array of Device Item Pointers

	try {
		USES_CONVERSION;
		LOGFMTI( "IOPCSyncIO2::ReadMaxAge %ld items of group %ls", dwCount, pGGroup->m_Name );

		aVal.Init(     dwCount, ppvValues );      // Allocate and initialize the result arrays
		aQual.Init(    dwCount, ppwQualities );
		aTStamps.Init( dwCount, ppftTimeStamps );

		// Get the device items associated with the specified server item handles
		hrRet = pGGroup->GetDItems(
			dwCount,                // Number of requested items
			phServer,               // Server item handles of the requested items
			OPC_READABLE,           // Returened items must have read access
			*ppvValues,             // Returns the Requested Data Types
			&dwCountValid,          // Returns the number valid Device Item pointers 
			ppErrors,               // Indicates the valid Device Item pointers 
			&ppDItems               // The Device Item pointers
			);

		_OPC_CHECK_HR( hrRet );
		if (dwCountValid == 0) throw S_FALSE;     // There are no valid items

		// Contains the pointers to the Device Items for
		// which the value needs to be updated from the Device.
		CAutoVectorPtr<DaDeviceItem*> avpTmpBufferForItemsToReadFromDevice;
		if (!avpTmpBufferForItemsToReadFromDevice.Allocate( dwCount )) throw E_OUTOFMEMORY;
		memset( avpTmpBufferForItemsToReadFromDevice, NULL, dwCount * sizeof (DaDeviceItem*) );

		FILETIME ftNow;
		HRESULT hrRet = CoFileTimeNow( &ftNow );
		_OPC_CHECK_HR( hrRet );

		hrRet = m_pServer->InternalReadMaxAge(
			&ftNow,
			dwCount,
			ppDItems,
			avpTmpBufferForItemsToReadFromDevice,
			pdwMaxAge,
			*ppvValues,
			*ppwQualities,
			*ppftTimeStamps,
			*ppErrors
			);

		_ASSERTE( SUCCEEDED( hrRet ) );           // Must return S_OK or S_FALSE
	}
	catch (HRESULT hrEx) {
		if (hrEx != S_FALSE) {
			aVal.Cleanup();
			aQual.Cleanup();
			aTStamps.Cleanup();

			if (*ppErrors) {
				ComFree( *ppErrors );
				*ppErrors = NULL;
			}
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

	ReleaseGenericGroup();
	return hrRet;
}



//=========================================================================
// IOPCSyncIO2::WriteVQT                                          INTERFACE
// ---------------------
//    Writes one or more Values, Qualities and TimeStamps for
//    the specified Items.
//=========================================================================
STDMETHODIMP DaGroup::WriteVQT(
									 /* [in] */                    DWORD             dwCount,
									 /* [size_is][in] */           OPCHANDLE      *  phServer,
									 /* [size_is][in] */           OPCITEMVQT     *  pItemVQT,
									 /* [size_is][size_is][out] */ HRESULT        ** ppErrors )
{
	LOGFMTI( "IOPCSyncIO2::WriteVQT" );

	return WriteSync( dwCount, phServer, pItemVQT, ppErrors );
}



///////////////////////////////////////////////////////////////////////////
//////////////////////////// IOPCItemDeadBandMgt //////////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCItemDeadbandMgt::SetItemDeadband                           INTERFACE
// ------------------------------------
//    Sets the PercentDeadband for the specified items.
//=========================================================================
STDMETHODIMP DaGroup::SetItemDeadband(
	/* [in] */                    DWORD             dwCount,
	/* [size_is][in] */           OPCHANDLE      *  phServer,
	/* [size_is][in] */           FLOAT          *  pPercentDeadband,
	/* [size_is][size_is][out] */ HRESULT        ** ppErrors)
{
	LOGFMTI( "IOPCItemDeadbandMgt::SetItemDeadband" );

	return ItemDeadband( dwCount, phServer, ppErrors, pPercentDeadband );
}



//=========================================================================
// IOPCItemDeadbandMgt::GetItemDeadband                           INTERFACE
// ------------------------------------
//    Gets the PercentDeadband for the specified items.
//=========================================================================
STDMETHODIMP DaGroup::GetItemDeadband( 
	/* [in] */                    DWORD             dwCount,
	/* [size_is][in] */           OPCHANDLE      *  phServer,
	/* [size_is][size_is][out] */ FLOAT          ** ppPercentDeadband,
	/* [size_is][size_is][out] */ HRESULT        ** ppErrors)
{
	LOGFMTI( "IOPCItemDeadbandMgt::GetItemDeadband" );

	CFixOutArray< FLOAT >   fxaPercentDeadband;  // Use global COM memory
	HRESULT                 hr = S_OK;
	try {
		fxaPercentDeadband.Init( dwCount, ppPercentDeadband, 0.0 );
		hr = ItemDeadband( dwCount, phServer, ppErrors, NULL, *ppPercentDeadband );
		_OPC_CHECK_HR( hr );
	}
	catch (HRESULT hrEx) {
		fxaPercentDeadband.Cleanup();
		hr = hrEx;
	}
	return hr;
}



//=========================================================================
// IOPCItemDeadbandMgt::ClearItemDeadband                         INTERFACE
// --------------------------------------
//    Clears the PercentDeadband for the specified items.
//=========================================================================
STDMETHODIMP DaGroup::ClearItemDeadband(
	/* [in] */                    DWORD             dwCount,
	/* [size_is][in] */           OPCHANDLE      *  phServer,
	/* [size_is][size_is][out] */ HRESULT        ** ppErrors)
{
	LOGFMTI( "IOPCItemDeadbandMgt::ClearItemDeadband" );

	return ItemDeadband( dwCount, phServer, ppErrors );
}

///////////////////////////////////////////////////////////////////////////



//=========================================================================
// WriteSync                                                       INTERNAL
// ---------
//    Implementation for IOPCSyncIO::Write() and IOPCSyncIO2::WriteVQT().
//=========================================================================
HRESULT DaGroup::WriteSync(
								 /* [in] */                    DWORD             dwCount,
								 /* [size_is][in] */           OPCHANDLE      *  phServer,
								 /* [size_is][in] */           OPCITEMVQT     *  pItemVQT,
								 /* [size_is][size_is][out] */ HRESULT        ** ppErrors )
{
	*ppErrors = NULL;                            // Note : Proxy/Stub checks if the pointers are NULL

	if (dwCount == 0) {
		LOGFMTE( "WriteSync() failed with invalid argument(s): dwCount is 0" );
		return E_INVALIDARG;
	}

	DaGenericGroup* pGGroup;
	HRESULT hrRet = GetGenericGroup( &pGGroup ); // Check group state and get the pointer
	if (FAILED( hrRet )) {
		LOGFMTE( "WriteSync() failed with error: No generic Group" );
		return hrRet;
	}

	DWORD          i;
	DWORD          dwCountValid= 0;              // Number of valid items
	DaDeviceItem**  ppDItems    = NULL;           // Array of Device Item Pointers

	try {
		USES_CONVERSION;
		LOGFMTI( "WriteSync() %ld items of group %s", dwCount, W2A( pGGroup->m_Name ) );

		// Get the Device Items associated with the specified server item handles
		hrRet = pGGroup->GetDItems(
			dwCount,                // Number of requested items
			phServer,               // Server item handles of the requested items
			OPC_WRITEABLE,          // Returened items must have write access
			NULL,                   // Data Types not required
			&dwCountValid,          // Returns the number valid Device Item pointers 
			ppErrors,               // Indicates the valid Device Item pointers 
			&ppDItems               // The Device Item pointers
			);    

		_OPC_CHECK_HR( hrRet );
		if (dwCountValid == 0)  throw S_FALSE;

		HRESULT hrWrite = pGGroup->InternalWriteVQT(
			dwCount,
			ppDItems,
			pItemVQT,
			*ppErrors
			);     

		_ASSERTE( SUCCEEDED( hrWrite ) );         // Must return S_OK or S_FALSE

		if (hrWrite == S_FALSE && hrRet == S_OK) {
			hrRet = S_FALSE;
		}
	}
	catch (HRESULT hrEx) {
		if (hrEx != S_FALSE) {
			if (*ppErrors) {
				ComFree( *ppErrors );
				*ppErrors = NULL;
			}
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

	ReleaseGenericGroup();
	return hrRet;
}



//=========================================================================
// ItemDeadband                                                     INTERNAL
// ------------
//    Implementation for IOPCItemDeadbandMgt::SetItemDeadband(), 
//    IOPCItemDeadbandMgt::GetItemDeadband() and
//    IOPCItemDeadbandMgt::ClearItemDeadband().
//
// The parameters pPercentDeadbandIn and pPercentDeadbandOut specified
// which functionality is performed:
//
// SetItemDeadband
//    pPercentDeadbandIn != NULL, pPercentDeadbandOut = NULL
//
// GetItemDeadband
//    pPercentDeadbandIn = NULL,  pPercentDeadbandOut != NULL
//
// ClearItemDeadband
//    pPercentDeadbandIn = NULL,  pPercentDeadbandOut = NULL         
//
//=========================================================================
HRESULT DaGroup::ItemDeadband(
									/* [in] */                    DWORD             dwCount,
									/* [size_is][in] */           OPCHANDLE      *  phServer,
									/* [size_is][size_is][out] */ HRESULT        ** ppErrors,
									/* [size_is][in] */           FLOAT          *  pPercentDeadbandIn /* = NULL */,
									/* [size_is][in] */           FLOAT          *  pPercentDeadbandOut /* = NULL */ )
{
	*ppErrors = NULL;                            // Note : Proxy/Stub checks if the pointers are NULL

	if (dwCount == 0) {
		LOGFMTE( "ItemDeadband() failed with invalid argument(s): dwCount is 0" );
		return E_INVALIDARG;
	}

	DaGenericGroup* pGGroup;
	HRESULT hrRet = GetGenericGroup( &pGGroup ); // Check group state and get the pointer
	if (FAILED( hrRet )) {
		LOGFMTE( "ItemDeadband() failed with error: No generic Group" );
		return hrRet;
	}

	DWORD          i;
	DWORD          dwCountValid= 0;              // Number of valid items
	DaDeviceItem**  ppDItems    = NULL;           // Array of Device Item Pointers

	try {
		USES_CONVERSION;
		LOGFMTI( "ItemDeadband() %ld items of group %s", dwCount, W2A( pGGroup->m_Name ) );

		// Get the Device Items associated with the specified server item handles
		hrRet = pGGroup->GetDItems(
			dwCount,                // Number of requested items
			phServer,               // Server item handles of the requested items
			0,                      // No access rights filtering
			NULL,                   // Data Types not required
			&dwCountValid,          // Returns the number valid Device Item pointers 
			ppErrors,               // Indicates the valid Device Item pointers 
			&ppDItems               // The Device Item pointers
			);

		_OPC_CHECK_HR( hrRet );
		if (dwCountValid == 0)  throw S_FALSE;

		// Handle Deadband for all Items
		for (i=0; i<dwCount; i++) {
			if (SUCCEEDED( (*ppErrors)[i] )) {

				if (pPercentDeadbandIn) {
					(*ppErrors)[i] = ppDItems[i]->SetItemDeadband( pPercentDeadbandIn[i] );
				}
				else if (pPercentDeadbandOut) {
					(*ppErrors)[i] = ppDItems[i]->GetItemDeadband( &pPercentDeadbandOut[i] );
				}
				else {
					(*ppErrors)[i] = ppDItems[i]->ClearItemDeadband();
				}            
				if (FAILED( (*ppErrors)[i] )) {
					hrRet = S_FALSE;
				}
			}
		}
	}
	catch (HRESULT hrEx) {
		if (hrEx != S_FALSE) {
			if (*ppErrors) {
				ComFree( *ppErrors );
				*ppErrors = NULL;
			}
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

	ReleaseGenericGroup();
	return hrRet;
}
//DOM-IGNORE-END
