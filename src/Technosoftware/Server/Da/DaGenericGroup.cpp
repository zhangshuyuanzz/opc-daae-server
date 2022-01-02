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
#include "FixOutArray.h"
#include "DaGenericGroup.h"
#include "DaGenericServer.h"
#include "DaPublicGroup.h"
#include "Logger.h"

//=====================================================================================
// Constructor
//=====================================================================================
DaGenericGroup::DaGenericGroup( void )
{
	m_Created = FALSE;

	m_RefCount = 0;                                     
	m_ToKill = FALSE;
	m_Active = FALSE;
	m_fCallbackEnable = TRUE;     // Default must be TRUE.

	m_pServer = NULL;
	m_pServerHandler = NULL;

	m_Name = NULL;

	m_DataCallback          = NULL;
	m_DataTimeCallback      = NULL;
	m_WriteCallback         = NULL;
	m_DataTimeCallbackDisp  = NULL;
	m_WriteCallbackDisp     = NULL;

	m_StreamData      = 0;        // Async handling formats
	m_StreamDataTime  = 0;
	m_StreamWrite     = 0;

	m_dwKeepAliveTime = 0;
	m_dwKeepAliveCount= 0;

	// for access to members of this group (mostly m_RefCount and m_ToKill)
	InitializeCriticalSection( &m_CritSec );

	// for synchronisation of access to items of this group
	InitializeCriticalSection( &m_ItemsCritSec );

	// initialize critical section for accessing 
	// asynchronouus threads list 
	InitializeCriticalSection( &m_AsyncThreadsCritSec );

	// initialize critical section for accessing
	// callback members 
	InitializeCriticalSection( &m_CallbackCritSec );

	//
	InitializeCriticalSection( &m_UpdateRateCritSec );
}

//=====================================================================================
// Initializer
//=====================================================================================
HRESULT DaGenericGroup::Create( DaGenericServer *pServer,
							  LPCWSTR         Name,
							  BOOL            bActive,
							  long            *pTimeBias,
							  long            dwRequestedUpdateRate,
							  long            hClientGroupHandle,
							  float          *pPercentDeadband,
							  long            dwLCID,
							  BOOL            bPublicGroup,
							  long            hPublicGroupHandle,
							  long           *phServerGroupHandle
							  )
{

	HRESULT                 res;

	DaPublicGroupManager     *pgHandler;
	DaPublicGroup            *pPG;
	long                    i, ni;
	HRESULT                 *pErr;
	DaGenericItem            *NewGI;
	DaDeviceItem             **ppDItems;
	OPCITEMDEF              *pID;
	long                    hServerHandle;


	_ASSERTE( pServer != NULL );
	_ASSERTE( phServerGroupHandle != NULL );

	if (m_Created == TRUE) {
		// Already created this group
		return E_FAIL;
	}

	// Register the Clipboard Formats for asynchronous callback handling
	m_StreamData     = RegisterClipboardFormat( _T("OPCSTMFORMATDATA") );
	if (!m_StreamData) {
		return HRESULT_FROM_WIN32( GetLastError() );
	}
	m_StreamDataTime = RegisterClipboardFormat( _T("OPCSTMFORMATDATATIME") );
	if (!m_StreamDataTime) {
		return HRESULT_FROM_WIN32( GetLastError() );
	}
	m_StreamWrite    = RegisterClipboardFormat( _T("OPCSTMFORMATWRITECOMPLETE") );
	if (!m_StreamWrite) {
		return HRESULT_FROM_WIN32( GetLastError() );
	}


	m_pServer = pServer;
	m_pServerHandler = pServer->m_pServerHandler;

	m_ActualBaseUpdateRate = m_pServer->GetActualBaseUpdateRate();
	m_RefCount = 0;
	m_ToKill = FALSE;

	// initialized tick information for client update
	m_Ticks = 1;
	m_TickCount = 1;

	m_Name = WSTRClone( Name, NULL);
	if ( m_Name == NULL ) {
		res = E_OUTOFMEMORY;
		goto CreateExit0;

	}
	m_bPublicGroup = bPublicGroup;
	m_dwLCID = dwLCID;

	// Percent Deadband
	if (pPercentDeadband == NULL)
		m_PercentDeadband = 0.0;                  // use default Percent Deadband
	else {
		// range check
		if (*pPercentDeadband < 0.0 || *pPercentDeadband > 100.0) {
         LOGFMTE( "Create() failed with invalid argument(s): Percent Deadband out of range (0...100)" );
			res = E_INVALIDARG;
			goto CreateExit2;
		}
		m_PercentDeadband = *pPercentDeadband;    // use specified Percent Deadband
	}

	// Time Bias
	if (pTimeBias == NULL) {                     // use default Time Bias

		TIME_ZONE_INFORMATION tzi;
		if (GetTimeZoneInformation( &tzi ) == TIME_ZONE_ID_INVALID) {
			res = HRESULT_FROM_WIN32( GetLastError() );
			goto CreateExit2;
		}
		m_TimeBias = tzi.Bias;
	}
	else                                         // use specified Time Bias
		m_TimeBias = *pTimeBias;

	ReviseUpdateRate( dwRequestedUpdateRate );

	m_hClientGroupHandle = hClientGroupHandle;

	// handle 0 corresponds to invalid item ...
	// so fill with NULL and it won't be used again!
	m_oaItems.PutElem( 0, NULL );
	m_oaCOMItems.PutElem( 0, NULL );


	// if it is a public group: add the items listed in the public group!
	// note that if there are errors they are ignored!
	if ( bPublicGroup == TRUE ) {

		// get the public group handler 
		pgHandler = &(m_pServerHandler->publicGroups_);
		// get and nail the group
		res = pgHandler->GetGroup( hPublicGroupHandle, &pPG );
		if ( FAILED( res ) ) {
			goto PGCreateExit0;
		}
		m_hPublicGroupHandle = hPublicGroupHandle;

		// get number of public groups in the item  
		ni = pPG->m_TotItems;

		// create errors array (indeed they are ignored)
		pErr = new HRESULT [ ni ];
		if (pErr == NULL ) {
			res = E_OUTOFMEMORY;
			goto PGCreateExit1;
		}

		// create DeviceItem temp work array were the generated items will be stored
		ppDItems = new DaDeviceItem*[ ni ];
		if( ppDItems == NULL ) {
			res = E_OUTOFMEMORY;
			goto PGCreateExit2;
		}

		// initialize new allocated arrays
		for (i=0; i < ni; i++) {
			pErr[i] = S_OK;
			ppDItems[i] = NULL;
		}

		// check the item objects
		res = m_pServerHandler->OnValidateItems(OPC_VALIDATEREQ_DEVICEITEMS,
			TRUE,               // Blob Update
			ni,
			pPG->m_ItemDefs,
			ppDItems,
			NULL,
			pErr );
		if ( FAILED(res) ) {
			goto PGCreateExit3;
		}

		// insert the successfully validated items in the groups item list
		// and itemresult datastructures
		i = 0;
		while ( i < ni ) {

			if (SUCCEEDED( pErr[i] )) {
				// add only valid item to the group
				// because public groups may contain bad items

				if (ppDItems[i]->Killed()) {
					pErr[i] = OPC_E_UNKNOWNITEMID;
				}
				else {

					// OPCITEMDEF struct
					pID = &(pPG->m_ItemDefs[i]);

					BOOL fPhyvalItem = FALSE;
					if (pPG->m_ExtItemDefs) {
						fPhyvalItem = pPG->m_ExtItemDefs[i].m_fPhyvalItem;
					}
					// Create a new GenericItem
					NewGI = new DaGenericItem();
					if( NewGI == NULL ) {
						res = E_OUTOFMEMORY;
						goto PGCreateExit4;
					}
					res = NewGI->Create( 
						pID->bActive,      // active mode flag
						pID->hClient,      // client handle
						pID->vtRequestedDataType,
						this,
						ppDItems[i],
						&hServerHandle,    // link to DeviceItem
						fPhyvalItem );

					if( FAILED( res ) ) {
						delete NewGI;
						goto PGCreateExit4;
					}
				} // kill flag is not set

				ppDItems[i]->Detach();              // Attach from OnValidateItems() no longer required
			}

			i ++;
		}

		// release referenced public group
		pgHandler->ReleaseGroup( hPublicGroupHandle );

		delete [] ppDItems;
		delete [] pErr;

	} // bPublicGroup == TRUE


	// to ensure activation recognition
	m_Active = FALSE ;                    

	// attach to the server object
	m_pServer->Attach();

	// insert in server group list
	EnterCriticalSection( &(m_pServer->m_GroupsCritSec) );

	m_hServerGroupHandle = m_pServer->m_GroupList.New();

	// insert in the server GroupList
	res = m_pServer->m_GroupList.PutElem( m_hServerGroupHandle , this);
	if ( FAILED(res) ) {
		goto CreateExit1;
	}
	LeaveCriticalSection( &(m_pServer->m_GroupsCritSec) );

	// ::::::::::::::::::::::
	// could create the group
	m_Created = TRUE;

	// if group active then activate the update thread
	SetActiveState( bActive );           

	*phServerGroupHandle = m_hServerGroupHandle;

	return S_OK;

CreateExit1:
	LeaveCriticalSection( &(m_pServer->m_GroupsCritSec) );
	m_pServer->Detach();
	goto CreateExit0;

PGCreateExit4:
	// destroy the already created generic items
	{
		HRESULT  hrtmp;
		long     ih;

		hrtmp = m_oaItems.First( &ih );
		while (SUCCEEDED( hrtmp )) {
			m_oaItems.GetElem( ih, &NewGI );
			if (NewGI)  delete NewGI;
			hrtmp = m_oaItems.Next( ih, &ih );
		}
	}
	for (; i < ni; i++) {
		if (ppDItems[i]) {
			ppDItems[i]->Detach();
		}
	}

PGCreateExit3:
	delete [] ppDItems;

PGCreateExit2:
	delete [] pErr;

PGCreateExit1:
	pgHandler->ReleaseGroup( hPublicGroupHandle );

PGCreateExit0:
	;

CreateExit2:
	WSTRFree( m_Name );
	m_Name = NULL;

CreateExit0:
	m_pServer = NULL;
	m_pServerHandler = NULL;
	return res;
}


//=====================================================================================
// Create (Clone)
// --------------
// Clone To Private Group Initializer
//
//    Active is set to FALSE !   
//    Actually also the asynch stuff is not copied!
//=====================================================================================
HRESULT DaGenericGroup::Create(
							  DaGenericGroup   *pCloned,
							  LPCWSTR         NewName,
							  long            *phServerGroupHandle
							  )
{
	long           itemHandle;
	long           clonedIH, clonedNIH;
	DaGenericItem   *pClonedItem, *pItem;
	HRESULT        res;

	_ASSERTE( pCloned != NULL );
	_ASSERTE( phServerGroupHandle != NULL );

	if (m_Created == TRUE) {
		// already created this group
		return E_FAIL;
	}

	// Register the Clipboard Formats for asynchronous callback handling
	m_StreamData     = RegisterClipboardFormat( _T("OPCSTMFORMATDATA") );
	if (!m_StreamData) {
		return HRESULT_FROM_WIN32( GetLastError() );
	}
	m_StreamDataTime = RegisterClipboardFormat( _T("OPCSTMFORMATDATATIME") );
	if (!m_StreamDataTime) {
		return HRESULT_FROM_WIN32( GetLastError() );
	}
	m_StreamWrite    = RegisterClipboardFormat( _T("OPCSTMFORMATWRITECOMPLETE") );
	if (!m_StreamWrite) {
		return HRESULT_FROM_WIN32( GetLastError() );
	}

	m_pServer = pCloned->m_pServer;
	m_pServerHandler = m_pServer->m_pServerHandler;

	m_Name = WSTRClone( NewName, NULL);
	if ( m_Name == NULL ) {
		res = E_OUTOFMEMORY;
		goto CreateCloneExit0;
	}

	m_ActualBaseUpdateRate = m_pServer->GetActualBaseUpdateRate();
	m_RefCount = 0;                                      
	m_ToKill = FALSE;

	// initialized tick information for client update
	m_Ticks = 1;
	m_TickCount = 1;

	// Clone to private!
	m_bPublicGroup = FALSE;

	m_PercentDeadband = pCloned->m_PercentDeadband;

	m_dwLCID = pCloned->m_dwLCID;
	m_TimeBias = pCloned->m_TimeBias;

	ReviseUpdateRate( pCloned->m_RequestedUpdateRate );   

	m_hClientGroupHandle = pCloned->m_hClientGroupHandle;

	// handle 0 corresponds to invalid item ...
	// so fill with NULL and it won't be used again!
	m_oaItems.PutElem( 0, NULL );
	m_oaCOMItems.PutElem( 0, NULL );


	// enter critical section of cloned group items
	EnterCriticalSection( &( pCloned->m_ItemsCritSec) );

	// copy all the items of the group to be cloned
	res = pCloned->m_oaItems.First( &clonedIH );
	while ( SUCCEEDED( res ) ) {
		// get the item of the cloned group
		pCloned->m_oaItems.GetElem( clonedIH, &pClonedItem );
		if ( pClonedItem->Killed() == FALSE ) {
			pItem = new DaGenericItem( );
			if ( pItem == NULL ) {
				res = E_OUTOFMEMORY;
				goto CreateCloneExit1;
			}

			res = pItem->Create( 
				this, 
				pClonedItem, 
				&itemHandle
				);
			if ( FAILED(res) ) {
				delete pItem;
				goto CreateCloneExit1;
			}

		}

		res = pCloned->m_oaItems.Next( clonedIH, &clonedNIH );
		clonedIH = clonedNIH;
	}
	LeaveCriticalSection( &( pCloned->m_ItemsCritSec) );


	// to ensure activation recognition
	m_Active = FALSE ;                    


	// attach to the server object
	m_pServer->Attach();


	// insert in the server GroupList
	EnterCriticalSection( &(m_pServer->m_GroupsCritSec) );

	m_hServerGroupHandle = m_pServer->m_GroupList.New();

	// insert in the GroupList
	res = m_pServer->m_GroupList.PutElem( m_hServerGroupHandle, this );
	if ( FAILED(res) ) {
		goto CreateCloneExit2;
	}
	LeaveCriticalSection( &(m_pServer->m_GroupsCritSec) );

	// ::::::::::::::::::::::
	// could create the group
	m_Created = TRUE;

	// if active then create the update thread
	SetActiveState( FALSE );

	*phServerGroupHandle = m_hServerGroupHandle;

	return S_OK;

CreateCloneExit2:
	LeaveCriticalSection( &(m_pServer->m_GroupsCritSec) );
	m_pServer->Detach();

CreateCloneExit1:
	// destroy the already created generic items
	{
		HRESULT  hrtmp;
		long     ih;

		hrtmp = m_oaItems.First( &ih );
		while (SUCCEEDED( hrtmp )) {
			m_oaItems.GetElem( ih, &pItem );
			if (pItem)  delete pItem;
			hrtmp = m_oaItems.Next( ih, &ih );
		}
	}

	LeaveCriticalSection( &( pCloned->m_ItemsCritSec) );

CreateCloneExit0:
	m_pServer = NULL;
	m_pServerHandler = NULL;
	return res;
}

//=====================================================================================
// Destructor
//    
//=====================================================================================
DaGenericGroup::~DaGenericGroup()
{

	if (m_Created == TRUE) {
		// here there are only shut-down 
		// proceedings reguarding a successfully created instance

		EnterCriticalSection( &m_CallbackCritSec );

      if (m_DataCallback) {
         core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( m_DataCallback );
      }
      if (m_DataTimeCallback) {
         core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( m_DataTimeCallback );
      }
      if (m_WriteCallback) {
         core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( m_WriteCallback );
      }

      if (m_DataTimeCallbackDisp) {
         core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( m_DataTimeCallbackDisp );
      }
      if (m_WriteCallbackDisp) {
         core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( m_WriteCallbackDisp );
      }

		LeaveCriticalSection( &m_CallbackCritSec );

		// delete from server list
		EnterCriticalSection( &m_pServer->m_GroupsCritSec );
		m_pServer->m_GroupList.PutElem( m_hServerGroupHandle, NULL );
		LeaveCriticalSection( &m_pServer->m_GroupsCritSec );

		m_pServer->Detach();
	}

	// delete group data
	if (m_Name) {
		WSTRFree( m_Name );
	}

	// kill local data
	DeleteCriticalSection( &m_CritSec );
	DeleteCriticalSection( &m_ItemsCritSec );
	DeleteCriticalSection( &m_AsyncThreadsCritSec );
	DeleteCriticalSection( &m_CallbackCritSec );
	DeleteCriticalSection( &m_UpdateRateCritSec );
}




//=====================================================================================
//  Attach to class by incrementing the RefCount
//  Returns the current RefCount
//  or -1 if the ToKill flag is set
//=====================================================================================
int  DaGenericGroup::Attach( void ) 
{
	long i;

	EnterCriticalSection( &m_CritSec );

	if( m_ToKill ) {
		LeaveCriticalSection( &m_CritSec );
		return -1 ;
	}
	m_RefCount ++;
	i = m_RefCount;

	LeaveCriticalSection( &m_CritSec );
	return i;
}


//=====================================================================================
//  Detach from class by decrementing the RefCount.
//  Kill the class instance if request pending.
//  Returns the current RefCount or -1 if it was killed.
//=====================================================================================
int  DaGenericGroup::Detach( void )
{
	long i;

	EnterCriticalSection( &m_CritSec );

	_ASSERTE( m_RefCount > 0 );

	m_RefCount-- ;

	if( ( m_RefCount == 0 )                      // not referenced
		&&  ( m_ToKill == TRUE ) ){            // kill request
			LeaveCriticalSection( &m_CritSec );
			delete this;                              // remove from memory
			return -1 ;
	}
	i = m_RefCount ;

	LeaveCriticalSection( &m_CritSec );
	return i;
}


//=====================================================================================
//  Kill the class instance.
//  If the RefCount is > 0 then only the kill request flag is set.
//  Returns the current RefCount or -1 if it was killed.
//=====================================================================================
int  DaGenericGroup::Kill( void )
{
	long                    i, newi;
	HRESULT                 res;
	DaGenericItem            *theItem;  


	EnterCriticalSection( &m_CritSec );

	// remove update notification if there is one
	SetActiveState( FALSE);

	// kill the items of this group
	EnterCriticalSection( &m_ItemsCritSec );
	res = m_oaItems.First( &i );
	while ( SUCCEEDED( res ) ) {
		res = m_oaItems.GetElem( i, &theItem );
		if ( theItem->Killed() == FALSE ) {
			theItem->Kill();
		}
		//res = m_oaItems.PutElem( i, NULL );
		res = m_oaItems.Next( i, &newi );
		i = newi;
	}
	LeaveCriticalSection( &m_ItemsCritSec );

	// !!!??? ev. force delete of all COM items!
	// so that the server will shutdown even if there are 
	// bad clients

	if( m_RefCount ) {            // still referenced
		m_ToKill = TRUE ;          // set kill request flag
	}
	else {
		LeaveCriticalSection( &m_CritSec );
		delete this ;
		return -1 ;
	}  
	i = m_RefCount ;

	LeaveCriticalSection( &m_CritSec );
	return i;
}

//=====================================================================================
// tells whether class instance was killed or not!
//=====================================================================================
BOOL DaGenericGroup::Killed()
{
	BOOL b;

	EnterCriticalSection( &m_CritSec );
	b = m_ToKill;
	LeaveCriticalSection( &m_CritSec );
	return b;
}



//=====================================================================================
// Set Group Name
// --------------
// Duplicate name must be controlled outside
// and the critical section for accessing a group name
// (that is daGenericServer_->m_GroupsCritSec) must be entered
// before calling this method
//=====================================================================================
HRESULT DaGenericGroup::set_Name( LPWSTR Name )
{
	// free old name
	if ( m_Name != NULL ) {
		WSTRFree( m_Name, NULL);
	}
	// copy new name
	m_Name = WSTRClone( Name, NULL );
	if ( m_Name != NULL ) {
		return S_OK;
	} else {
		return E_OUTOFMEMORY;
	}

}




//=====================================================================================
// Get Item
//    Get and Nail an item so that it cannot be deleted until
//    ReleaseGenericItem is called
//=====================================================================================
HRESULT DaGenericGroup::GetGenericItem( long ItemServerHandle, DaGenericItem **item )
{
	HRESULT res;

	// 
	EnterCriticalSection( &m_ItemsCritSec );

	res = m_oaItems.GetElem( ItemServerHandle, item);
	if( ( FAILED( res ) ) || ((*item)->Killed() == TRUE) ) {
		// invalid handle or group has to be removed
		res = OPC_E_INVALIDHANDLE;
		goto GetGenericItemExit0;
	}
	// increment ref count so that group cannot be deleted during access
	(*item)->Attach() ;
	res = S_OK;

GetGenericItemExit0:
	LeaveCriticalSection( &m_ItemsCritSec );
	return res;
}



//=====================================================================================
// ReleaseGenericItem
// Remove nail from item
//=====================================================================================
HRESULT DaGenericGroup::ReleaseGenericItem( long ItemServerHandle )
{
	DaGenericItem *item;
	long res;

	EnterCriticalSection( &m_ItemsCritSec );
	res = m_oaItems.GetElem( ItemServerHandle, &item);
	if( FAILED( res ) ) {
		res = OPC_E_INVALIDHANDLE;
		goto ReleaseGenericItemExit0;
	}
	// at this point the item could be makked as ToKill! (meaning ToDelete)
	item->Detach();
	res = S_OK;

ReleaseGenericItemExit0:
	LeaveCriticalSection( &m_ItemsCritSec );
	return res;
}




//=====================================================================================
// RemoveGenericItemNoLock
// -----------------------
// Removes an item from the group
// m_ItemsCritSec must be entered outside
//=====================================================================================
HRESULT DaGenericGroup::RemoveGenericItemNoLock( long ItemServerHandle )
{
	DaGenericItem*  pGItem;

	m_oaItems.GetElem( ItemServerHandle, &pGItem );
	if ((pGItem == NULL) || pGItem->Killed()) {
		return OPC_E_INVALIDHANDLE;
	}
	pGItem->Kill();
	return S_OK;
}




//=========================================================================
// GetCOMItem
// create if necessary and get the COM Item related to the GenericItem
// access synch must be done outside!
// Inputs:
//    hServerHandle: the handle of the item of the group
// Output:
//    created: is FALSE if a COM item was already created 
//             for the DaGenericItem with hServerHandle
//             if created is TRUE then there is no reference 
//             on the COM item and the caller has the right
//             (within the m_ItemsCritSec) to delete the COMItem
//=========================================================================
HRESULT DaGenericGroup::GetCOMItem(
								  long           hServerHandle,
								  DaGenericItem   *item,
								  DaItem     **COMItem,
								  int           *created
								  )
{
	HRESULT hres;


	*created = FALSE;

	hres = m_oaCOMItems.GetElem( hServerHandle, COMItem );
	if ( SUCCEEDED( hres ) ) {
		goto GetCOMItemBadExit0;
	}

	// Construct the COM object
	hres = CComObject<DaItem>::CreateInstance( (CComObject<DaItem>**)COMItem );
	if (FAILED( hres )) {
		goto GetCOMItemBadExit0;
	}

	// Initialize the COM object
	hres = (*COMItem)->Create( this, hServerHandle, item);
	if ( FAILED( hres ) ) {
		goto GetCOMItemBadExit1;
	}

	// Insert the COM Object in the groups list
	hres = m_oaCOMItems.PutElem( hServerHandle, *COMItem );
	if ( FAILED( hres ) ) {
		goto GetCOMItemBadExit1;
	}

	*created = TRUE;
	return S_OK;

GetCOMItemBadExit1:
	delete *COMItem;

GetCOMItemBadExit0:
	return hres;
}




//=========================================================================
// InternalRead      Read the given array of items.
// ------------
// This method is called from multiple threads 
//       - serving different clients
//       - handling sync/async read, refresh, update for the same client
// Non-reentrant parts must therefore be protected !
// if this method fails no item values initialisation is done!
//
// Note :
//    -  ppDItems may contain NULL values ...
//       these are the items eliminated before read ...
//       pItemValues[x].vDataValue.vt contains the requested data type
//       or VT_EMPTY if the canonical data type should be used.
//                                              
// return:
//    S_OK if succeeded for all items; otherwise S_FALSE.
//    Do not return other codes !
//=========================================================================
HRESULT DaGenericGroup::InternalRead(
									DWORD             dwSource,   // Cache / Device
									DWORD             numItems, 
									DaDeviceItem    ** ppItems,
									OPCITEMSTATE   *  pItemValues,
									HRESULT        *  errors,
									BOOL           *  pfPhyval /* = NULL */ )
{
	DWORD       i;
	VARIANT     *pvValue;
	DaDeviceItem *pDItem;
	HRESULT     hresReturn, hres;

	hresReturn = S_OK;

	if (dwSource == OPC_DS_DEVICE) {             // Refresh the chache for the requested items
		hresReturn = m_pServerHandler->OnRefreshInputCache( OPC_REFRESH_CLIENT, numItems, ppItems, errors );
		_ASSERTE( SUCCEEDED( hresReturn ) );      // Must return S_OK or S_FALSE
	}
	// handle read/write locking
	// Note : get_ItemValue() reads from cache.
	m_pServerHandler->readWriteLock_.BeginReading();

	for (i=0 ; i<numItems; i++) {

		pDItem = ppItems[i];

		if (pDItem == NULL) {                     // No refresh requested for this item
			continue;
		}
		if (FAILED( errors[i] )) {
			continue;                              // Cache refresh from device for this item failed
		}

		pvValue = &pItemValues[i].vDataValue;
		// Use the canonical data type if requested data
		// type is VT_EMPTY.
		if (V_VT( pvValue ) == VT_EMPTY) {
			V_VT( pvValue ) = pDItem->get_CanonicalDataType();
		}

		// Read from cache
		hres = pDItem->get_ItemValue(
			pvValue,                      // Value
			&pItemValues[i].wQuality,     // Quality
			&pItemValues[i].ftTimeStamp );// Timestamp

		pItemValues[i].wReserved = 0;             // initialize reserve field, Version 1.3

		if (FAILED( hres )) {
			errors[i] = hres;
			VariantClear( pvValue );
			hresReturn = S_FALSE;
		}
	}
	// handle read/write locking
	m_pServerHandler->readWriteLock_.EndReading();

	_ASSERTE( SUCCEEDED( hresReturn ) );         // Must return S_OK or S_FALSE
	return hresReturn;
}



//=========================================================================
// InternalWrite    Write the given array of items
// -------------
// This method is called from multiple threads 
//       - serving different clients
//       - handling sync/async write for the same client
// Non-reentrant parts must therefore be protected !
//                                     
// Note :   ppDItems may contain NULL values ...
//          these are the items eliminated before read ...
//
// !! Attention !!
//          This function can release Items from ppDItems[] so
//          the caller must check the pointers before the release.
//
// parameter:
//    [in]     numItems  Number of elements in ppItems[], pValues[]
//                         and errors[]
//    [in,out] ppItems     array [numItems] of pointer to device items
//    [in]     pValues     array [numItems] of VARIANT
//                         (which contains the values)
//    [in,out] errors     array [numItems] of HRESULT
//
// return:
//    S_OK if succeeded for all items; otherwise S_FALSE.
//    Do not return other codes !
//=========================================================================
HRESULT DaGenericGroup::InternalWriteVQT(
										DWORD             numItems,
										DaDeviceItem    ** ppItems,
										OPCITEMVQT     *  pVQTs,
										HRESULT        *  errors,
										BOOL           *  pfPhyval /* = NULL */ )
{
	return m_pServer->InternalWriteVQT(
		numItems, 
		ppItems,
		pVQTs,
		errors,
		pfPhyval,
		m_hServerGroupHandle );
}



//=====================================================================================
//  Get the group active mode state
//=====================================================================================
BOOL   DaGenericGroup::GetActiveState( void )
{
	if( Killed() ) {
		return FALSE ;
	}
	return m_Active ;
}



//=====================================================================================
// Change the group active mode state
//=====================================================================================
HRESULT DaGenericGroup::SetActiveState( BOOL NewState )
{
	EnterCriticalSection( &m_CritSec );

	if (NewState == m_Active) {
		LeaveCriticalSection( &m_CritSec );
		return S_OK;
	}

	// Update to the new state.
	m_Active = NewState;

	if (NewState) {
		ResetKeepAliveCounter();
	}

	// Handle the active state of all attached Device Items (don't change the
	// state of Generic Items).
	HRESULT        hres;
	DaGenericItem*  pGItem;
	long           idx;

	// Do it for all Generic Items of this group.
	EnterCriticalSection( &m_ItemsCritSec );
	EnterCriticalSection( &m_UpdateRateCritSec );
	hres = m_oaItems.First( &idx );
	while (SUCCEEDED( hres )) {
		hres = m_oaItems.GetElem( idx, &pGItem );
		if (pGItem) {

			if (pGItem->get_Active()) {            // Set state active for all Device Items attached to a              
				// Generic Item with active state and force as ubscription callback.
				if (NewState) {                     // Group state changed to active state.
					pGItem->AttachActiveCountOfDeviceItem();
					pGItem->ResetLastRead();         // Force a subscription callback
					m_TickCount = 1;                 // by next cycle.
				}
				else {
					pGItem->DetachActiveCountOfDeviceItem();
				}
			}
		}
		hres = m_oaItems.Next( idx, &idx );
	}
	LeaveCriticalSection( &m_UpdateRateCritSec );
	LeaveCriticalSection( &m_ItemsCritSec );
	LeaveCriticalSection( &m_CritSec );
	return S_OK;
}




//=====================================================================================
// make the group public
//=====================================================================================
HRESULT DaGenericGroup::MakePublic( long PGHandle )
{
	EnterCriticalSection( &m_CritSec );
	m_bPublicGroup = TRUE;
	m_hPublicGroupHandle = PGHandle;
	LeaveCriticalSection( &m_CritSec );
	return S_OK;
}

//=====================================================================================
// tells whether group public or not
// if public also returns the public group handle
//=====================================================================================
BOOL DaGenericGroup::GetPublicInfo( long *pPGHandle )
{
	BOOL b;

	EnterCriticalSection( &m_CritSec );
	b = m_bPublicGroup;
	if ( b == TRUE ) {
		*pPGHandle = m_hPublicGroupHandle;
	}
	LeaveCriticalSection( &m_CritSec );
	return b;
}


//=====================================================================================
// Set requested update rate and revise Update Rate for this group
//
// This function modifies the following data members:
//    m_RequestedUpdateRate:  The requested Update Rate
//    updateRate_:    The revised Update Rate
//    m_TickCount:            Number of required update cycles until next update
//    m_Ticks:                Number of required update cycles between two updates
//=====================================================================================
HRESULT DaGenericGroup::ReviseUpdateRate(
										long RequestedUpdateRate
										)
{
	HRESULT hr = S_OK;

	// cannot change update rate during callback
	EnterCriticalSection( &m_UpdateRateCritSec );

	try {
		m_RequestedUpdateRate = RequestedUpdateRate;
		hr = m_pServerHandler->ReviseUpdateRate( (DWORD)RequestedUpdateRate, (DWORD*)&m_RevisedUpdateRate );
		_OPC_CHECK_HR( hr );

		DWORD dwNewTicks;
		{
			long lNewTicks = m_RevisedUpdateRate / GetActualBaseUpdateRate();
			if (lNewTicks <= 0) throw E_FAIL;
			dwNewTicks = lNewTicks;
		}        

		if (dwNewTicks == m_Ticks) throw S_OK;

		DWORD dwTicksSinceLastUpdate = m_Ticks - m_TickCount;

		if (dwNewTicks <= dwTicksSinceLastUpdate) {
			m_TickCount = 1;                    // Force an update by next cycle
		}
		else {
			m_TickCount = dwNewTicks - dwTicksSinceLastUpdate;
		}

		m_Ticks = dwNewTicks;

	}
	catch (HRESULT hrEx) { hr = hrEx; }
	catch (...) { hr = E_FAIL; }

	LeaveCriticalSection( &m_UpdateRateCritSec );
	return hr;
}




//=====================================================================================
//=====================================================================================
DWORD DaGenericGroup::GetActualBaseUpdateRate( )
{
	DWORD temp;

	EnterCriticalSection( &m_UpdateRateCritSec );
	temp = m_ActualBaseUpdateRate;
	LeaveCriticalSection( &m_UpdateRateCritSec );

	return temp;
}



//=====================================================================================
//=====================================================================================
HRESULT DaGenericGroup::SetActualBaseUpdateRate( DWORD BaseUpdateRate )
{
	EnterCriticalSection( &m_UpdateRateCritSec );
	m_ActualBaseUpdateRate = BaseUpdateRate;
	LeaveCriticalSection( &m_UpdateRateCritSec );

	return S_OK;
}



//=====================================================================================
// KeepAliveTime
// -------------
//    Returns the currently active keep-alive time for the subscription
//    of this group.
//=====================================================================================
DWORD DaGenericGroup::KeepAliveTime()
{
	m_csKeepAlive.Lock();

	DWORD dwTime = m_dwKeepAliveTime;

	m_csKeepAlive.Unlock();
	return dwTime;
}



//=====================================================================================
// SetKeepAlive
// ------------
//    Sets the keep-alive time for the subscription of this group.
//=====================================================================================
HRESULT DaGenericGroup::SetKeepAlive(
									DWORD             dwKeepAliveTime,
									DWORD          *  pdwRevisedKeepAliveTime )
{
	m_csKeepAlive.Lock();

	HRESULT hr = S_OK;
	if (dwKeepAliveTime) {
		// We can use the same function to calculate the revised keep-alive
		// time as used to calculate the revised update rate
		hr = m_pServerHandler->ReviseUpdateRate( dwKeepAliveTime, &m_dwKeepAliveTime );
	}
	else {
		m_dwKeepAliveTime = 0;                    // Inactivate keep-alive callbacks
	}
	*pdwRevisedKeepAliveTime = m_dwKeepAliveTime;
	m_dwKeepAliveCount = m_dwKeepAliveTime / GetActualBaseUpdateRate();

	m_csKeepAlive.Unlock();
	if (FAILED( hr )) return hr;

	return (dwKeepAliveTime == m_dwKeepAliveTime) ? S_OK : OPC_S_UNSUPPORTEDRATE;
}



//=====================================================================================
// ResetKeepAliveCounter
//=====================================================================================
void DaGenericGroup::ResetKeepAliveCounter()
{
	m_csKeepAlive.Lock();
	m_dwKeepAliveCount = m_dwKeepAliveTime / GetActualBaseUpdateRate();
	m_csKeepAlive.Unlock();
}



//=====================================================================================
// Resets the Last Read Value and Quality of all Generic Items of the Group. 
//=====================================================================================
void DaGenericGroup::ResetLastReadOfAllGenericItems( void )
{
	HRESULT        hres;
	DaGenericItem*  pGItem;
	long           i;

	EnterCriticalSection( &m_ItemsCritSec );     // while accessing generic items array do not
	// allow add and delete group items.

	hres = m_oaItems.First( &i );
	while (SUCCEEDED( hres )) {
		m_oaItems.GetElem( i, &pGItem ) ;
		if (pGItem) {
			pGItem->ResetLastRead();
		}
		hres = m_oaItems.Next( i, &i );
	}

	LeaveCriticalSection( &m_ItemsCritSec );     // access to generic item array finished.
}



//=========================================================================
// GetDItemsAndStates                                             PROTECTED
// ------------------
//    Use this function if you need the device items with read access 
//    associated with the active items in the group.
//
// Parameters:
//
// OUT
//    ppErrors                Array of S_OK error codes for each item.
//                            The size is dwCountOut and the memory for the
//                            array is allocated from the global COM memory
//                            by this function.
//
//    pdwCountOut             Number of the elements in the arrays ppErrors,
//                            ppDItems and ppItemStates.
//
//    pppDItems               Array of device item pointers.
//                            The size is dwCountOut and the memory for the
//                            array is allocated from the heap.
//
//    ppItemStates            Array of OPCITEMSTATE. 
//                            NOTE : This function sets only the members 
//                            hClient and vDataValue.vt !
//                            The size is dwCountOut and the memory for the
//                            array is allocated from the heap.
//
//    ppfPhyvalItems          Array of booleans which marks items added
//                            with their physical value.
//                            The size is dwCountOut and the memory for the
//                            array is allocated from the heap.
//
// Return Code:
//    S_OK                    Function succeeded for all items.
//    E_XXX                   Function failed. All OUT parameters are
//                            NULL pointers
//
//=========================================================================
HRESULT DaGenericGroup::GetDItemsAndStates(
	// OUT
	/* [out][dwCount] */          HRESULT        ** ppErrors,
	/* [out] */                   DWORD          *  pdwCountOut,
	/* [out][*pdwCountOut] */     DaDeviceItem    ***pppDItems,
	/* [out][*pdwCountOut] */     OPCITEMSTATE   ** ppItemStates
	)
{
	_ASSERTE( ppErrors );
	_ASSERTE( pdwCountOut );
	_ASSERTE( pppDItems );
	_ASSERTE( ppItemStates );

	HRESULT hr = S_OK;

	CFixOutArray< HRESULT >             aErr;    // Use global COM memory
	CFixOutArray< DaDeviceItem*, FALSE > apDItems;// Use heap memory
	CFixOutArray< OPCITEMSTATE, FALSE > aIStates;// Use heap memory

	*pdwCountOut = 0;

	EnterCriticalSection( &m_ItemsCritSec );     // All active items are requested

	try {
		DWORD dwCount = m_oaItems.TotElem();

		aErr.Init( dwCount, ppErrors, S_OK );     // Allocate and initialize the result arrays
		apDItems.Init( dwCount, pppDItems, NULL );

		OPCITEMSTATE StateInit = { 0, 0, 0, 0, 0 };
		aIStates.Init( dwCount, ppItemStates, StateInit );

		DaGenericItem*  pGItem;
		DWORD          i;

		// All active items are requested
		i = 0;
		long idx;
		hr = m_oaItems.First( &idx );
		while (SUCCEEDED( hr )) {
			m_oaItems.GetElem( idx, &pGItem );
			if (pGItem && pGItem->get_Active()) {
				// Get the attached Device Item
				if (pGItem->AttachDeviceItem( &apDItems[i] ) > 0) {
					// Test if the item has the requested access rigth
					if (apDItems[i]->HasReadAccess()) {

						aIStates[i].hClient = pGItem->get_ClientHandle();
						V_VT(&aIStates[i].vDataValue) = pGItem->get_RequestedDataType();

						i++;                          // One additional Item successfully added
					}
					else {
						pGItem->DetachDeviceItem();   // Detach the Device Item and release the Generic Item
					}
				}
			}
			hr = m_oaItems.Next( idx, &idx );
		}
		*pdwCountOut = i;
		hr = S_OK;
	}
	catch (HRESULT hrEx) {
		aErr.Cleanup();
		apDItems.Cleanup();
		aIStates.Cleanup();
		*pdwCountOut = 0;
		hr= hrEx;
	}

	LeaveCriticalSection( &m_ItemsCritSec );
	return hr;

} // GetDItemsAndStates



//=========================================================================
// GetDItemsAndStates                                             PROTECTED
// ------------------
//    Use this function if you need the device items associated with 
//    the specified server item handles.
//
// Parameters:
//
// IN
//    dwCount                 Number of server item handles in phServer.
//
//    phServer                Array of server item handles.
//
//    pdwMaxAges              Array of Max Age definitions which
//                            corresponds to the server item handles.
//                            This parameter may be a NULL pointer if
//                            Max Age handling is not required.
//
//    dwAccessRightsFilter    Filter based on the DaAccessRights bit mask
//                            (OPC_READABLE or OPC_WRITEABLE).
//                            0 indicates no filtering.
// OUT
//    ppErrors                Array of errors for each requested item.
//                            The size is dwCount and the memory for the
//                            array is allocated from the global COM memory
//                            by this function.
//
//    pdwCountOut             Number of the elements in the arrays ppDItems
//                            ,ppItemStates and ppdwMaxAges (if specified).
//                            This counter also indicates the number of
//                            S_OK error codes in ppErrors.
//
//    pppDItems               Array of device item pointers.
//                            The size is dwCountOut and the memory for the
//                            array is allocated from the heap.
//
//    ppItemStates            Array of OPCITEMSTATE. 
//                            NOTE : This function sets only the members 
//                            hClient and vDataValue.vt !
//                            The size is dwCountOut and the memory for the
//                            array is allocated from the heap.
//
//    ppdwMaxAges             Array of the corresponding Max Age
//                            definitions. This parameter may be NULL and
//                            is only used if pdwMaxAges specifies the Max
//                            Age definitions which corresponds to the
//                            specified server item handles.
//                            The size is dwCountOut and the memory for the
//                            array is allocated from the heap.
// 
//    ppfPhyvalItems          Array of booleans which marks items added
//                            with their physical value.
//                            The size is dwCountOut and the memory for the
//                            array is allocated from the heap.
//
// Return Code:
//    S_OK                    Function succeeded for all items.
//    S_FALSE                 Function succeeded for some items. Not all
//                            item error codes are S_OK.
//                            dwCount differs from dwCountOut.
//    E_XXX                   Function failed. All OUT parameters are
//                            NULL pointers
//
//=========================================================================
HRESULT DaGenericGroup::GetDItemsAndStates(
	// IN
	/* [in] */                    DWORD             dwCount,
	/* [in][dwCount] */           OPCHANDLE      *  phServer,
	/* [in][dwCount] */           DWORD          *  pdwMaxAges,
	/* [in] */                    DWORD             dwAccessRightsFilter,
	// OUT
	/* [out][dwCount] */          HRESULT        ** ppErrors,
	/* [out] */                   DWORD          *  pdwCountOut,
	/* [out][*pdwCountOut] */     DaDeviceItem    ***pppDItems,
	/* [out][*pdwCountOut] */     OPCITEMSTATE   ** ppItemStates,
	/* [out][*pdwCountOut] */     DWORD          ** ppdwMaxAges
	)
{
	_ASSERTE( ppErrors );
	_ASSERTE( pdwCountOut );
	_ASSERTE( pppDItems );
	_ASSERTE( ppItemStates );
	// pdwMaxAge may be NULL
	if (pdwMaxAges) {
		_ASSERTE( ppdwMaxAges );                  // Must be specified if pdwMaxAge is not NULL
	}

	HRESULT hr = S_OK;

	CFixOutArray< HRESULT >             aErr;    // Use global COM memory
	CFixOutArray< DaDeviceItem*, FALSE > apDItems;// Use heap memory
	CFixOutArray< OPCITEMSTATE, FALSE > aIStates;// Use heap memory
	CFixOutArray< DWORD, FALSE >        aMaxAges;// Use heap memory

	*pdwCountOut = 0;
	try {
		aErr.Init( dwCount, ppErrors, S_OK );     // Allocate and initialize the result arrays
		apDItems.Init( dwCount, pppDItems, NULL );

		if (pdwMaxAges) {
			aMaxAges.Init( dwCount, ppdwMaxAges, 0 );
		}

		OPCITEMSTATE StateInit = { 0, 0, 0, 0, 0 };
		aIStates.Init( dwCount, ppItemStates, StateInit );

		DaGenericItem*  pGItem;
		DWORD          dwCurrentAccessRights;
		BOOL           fAccessRightsFit;
		DWORD          i;

		// Handle all requested items
		// The requested items are specified with their item server handle
		for (i=0; i<dwCount; i++) {
			// Get generic item with the specified server item handle
			if (SUCCEEDED( GetGenericItem( phServer[i], &pGItem ) )) {
				// Get the attached device item
				if (pGItem->AttachDeviceItem( &apDItems[*pdwCountOut] ) > 0) {
					// Test if the item has the requested access rigth
					fAccessRightsFit = TRUE;
					if (dwAccessRightsFilter) {
						apDItems[*pdwCountOut]->get_AccessRights( &dwCurrentAccessRights );
						if ((dwCurrentAccessRights & dwAccessRightsFilter ) == 0) {
							fAccessRightsFit = FALSE;
						}
					}
					if (!fAccessRightsFit) {
						pGItem->DetachDeviceItem();   // Detach the device item and release the generic item
						apDItems[*pdwCountOut] = NULL;
						aErr[i] = OPC_E_BADRIGHTS;
					}
					else {
						aIStates[*pdwCountOut].hClient               = pGItem->get_ClientHandle();
						V_VT( &aIStates[*pdwCountOut].vDataValue )   = pGItem->get_RequestedDataType();

						if (pdwMaxAges) {
							aMaxAges[*pdwCountOut] = pdwMaxAges[i];
						}

						(*pdwCountOut)++;
					}
				}
				else {
					aErr[i] = OPC_E_UNKNOWNITEMID;   // No longer available in the server address space
				}                                   // or the the kill flag is set

				ReleaseGenericItem( phServer[i] );  // This generic item is no longer used
			}
			else {
				aErr[i] = OPC_E_INVALIDHANDLE;      // There is no item with the specified server item handle
			}

			if (FAILED( aErr[i] )) {
				hr = S_FALSE;                       // Not all item related error codes are S_OK
			}

		} // Handle all requested items

		// Option: shrink outgoing arrays if S_FALSE is returned
	}
	catch (HRESULT hrEx) {
		aErr.Cleanup();
		apDItems.Cleanup();
		aIStates.Cleanup();
		aMaxAges.Cleanup();
		*pdwCountOut = 0;
		hr = hrEx;
	}

	return hr;

} // GetDItemsAndStates



//=========================================================================
// GetDItems                                                      PROTECTED
// ---------
//    Use this function if you need the device items and the Requested
//    Data Types associated with the specified server item handles.
//
//    Note: The returned Arrays has the same size like the Array with
//    the item handles. The position of the returned Device Item pointers
//    corresponds with the position of the server handle.
//
// Parameters:
//
// IN
//    dwCount                 Number of server item handles in phServer.
//
//    phServer                Array of server item handles.
//
//    dwAccessRightsFilter    Filter based on the DaAccessRights bit mask
//                            (OPC_READABLE or OPC_WRITEABLE).
//                            0 indicates no filtering. 
// IN / OUT
//    pvValues                Array of initialized Variants or NULL.
//                            On return the .vt members contains the
//                            Requested Data Type of the Items (if
//                            not NULL).
// OUT
//    pdwNumOfValidItems      The number of returned Device Item pointers
//                            (and also S_OK values in ppErrors).
//
//    ppErrors                Array of errors for each item.
//                            The size is dwCount and the memory for the
//                            array is allocated from the global COM memory
//                            by this function.
//
//    pppDItems               Array of device item pointers.
//                            The size is dwCountOut and the memory for the
//                            array is allocated from the heap.
//
//    ppfPhyvalItems          Array of booleans which marks items added
//                            with their physical value.
//                            The size is dwCount and the memory for the
//                            array is allocated from the heap.
//
// Return Code:
//    S_OK                    Function succeeded for all items.
//    S_FALSE                 Function succeeded for some items. Not all
//                            item error codes are S_OK.
//                            dwCount differs from pdwNumOfValidItems.
//    E_XXX                   Function failed. All OUT parameters are
//                            NULL pointers
//
//=========================================================================
HRESULT DaGenericGroup::GetDItems(
								 // IN
								 /* [in] */                    DWORD             dwCount,
								 /* [in][dwCount] */           OPCHANDLE      *  phServer,
								 /* [in] */                    DWORD             dwAccessRightsFilter,
								 /* [in][out] */               VARIANT        *  pvValues,
								 // OUT
								 /* [out] */                   DWORD          *  pdwNumOfValidItems,
								 /* [out][dwCount] */          HRESULT        ** ppErrors,
								 /* [out][dwCount] */          DaDeviceItem    ***pppDItems
								 )
{
	// pvValues may be NULL
	_ASSERTE( pdwNumOfValidItems );
	_ASSERTE( ppErrors );
	_ASSERTE( pppDItems );   

	HRESULT hr = S_OK;

	CFixOutArray< HRESULT >             aErr;    // Use global COM memory
	CFixOutArray< DaDeviceItem*, FALSE > apDItems;// Use heap memory

	*pdwNumOfValidItems = 0;
	try {
		apDItems.Init( dwCount, pppDItems, NULL );// Allocate and initialize the result arrays
		aErr.Init( dwCount, ppErrors, S_OK );

		DaGenericItem*  pGItem;
		DWORD          dwCurrentAccessRights;
		BOOL           fAccessRightsFit;
		DWORD          i;

		// Handle all requested items
		// The requested items are specified with their item server handle
		for (i=0; i<dwCount; i++) {
			// Get generic item with the specified server item handle
			if (SUCCEEDED( GetGenericItem( phServer[i], &pGItem ) )) {
				// Get the attached device item
				if (pGItem->AttachDeviceItem( &apDItems[i] ) > 0) {
					// Test if the item has the requested access rigths
					fAccessRightsFit = TRUE;
					if (dwAccessRightsFilter) {
						apDItems[i]->get_AccessRights( &dwCurrentAccessRights );
						if ((dwCurrentAccessRights & dwAccessRightsFilter ) == 0) {
							fAccessRightsFit = FALSE;
						}
					}
					if (!fAccessRightsFit) {
						pGItem->DetachDeviceItem();   // Detach the device item and release the generic item
						apDItems[i] = NULL;
						aErr[i] = OPC_E_BADRIGHTS;
					}
					else {
						if (pvValues) {
							V_VT( &pvValues[i] ) = pGItem->get_RequestedDataType();
						}
						(*pdwNumOfValidItems)++;
					}
				}
				else {
					aErr[i] = OPC_E_UNKNOWNITEMID;   // No longer available in the server address space
				}                                   // or the the kill flag is set

				ReleaseGenericItem( phServer[i] );  // This generic item is no longer used
			}
			else {
				aErr[i] = OPC_E_INVALIDHANDLE;      // There is no item with the specified server item handle
			}

			if (FAILED( aErr[i] )) {
				hr = S_FALSE;                       // Not all item related error codes are S_OK
			}

		} // Handle all requested items

	}
	catch (HRESULT hrEx) {
		aErr.Cleanup();
		apDItems.Cleanup();
		*pdwNumOfValidItems = 0;
		hr = hrEx;
	}

	return hr;

} // GetDItems



//=========================================================================
// CopyVariantArrayForValidItems                                  PROTECTED
// -----------------------------
//    Builds a copy of an array of VARIANTs. Only the values with a
//    succeeded code in the corresponding error array are copied.
//
//    Sample :
//          IN                               
//              src value   corresp. error
//             +-----------+-----------------+
//             | 100       | S_OK            |
//             | 200       | OPC_E_BADRIGHTS |
//             | 300       | S_OK            |
//             +-----------+-----------------+
//          OUT       
//              dest value
//             +-----------+
//             | 100       |
//             | 300       |
//             +-----------|
//
// Parameters:
//
// IN
//    dwCountAll              Number of entries in the arrays errors and
//                            pvarSrc.
//    errors                 Array of corresponding errors for each
//                            value in the source value array.
//    pvarSrc                 Array of VARIANT with the source values.
//
//    dwCountValid            Number of the elements in the array errors
//                            with succeeded code. Size of the output
//                            array ppvarDest.
// OUT
//    ppvarDest               Allocated array with  he copied VARIANTs.
//                            The memory for the array is allocated
//                            from the heap.
//
// Return Code:
//    S_OK                    Function succeeded for all items.
//    E_XXX                   Function failed. All OUT parameters are
//                            NULL pointers
//
//=========================================================================
HRESULT DaGenericGroup::CopyVariantArrayForValidItems(
	/* [in] */                    DWORD             dwCountAll,
	/* [in][dwCountAll] */        HRESULT        *  errors,
	/* [in][dwCountAll] */        VARIANT        *  pvarSrc,
	/* [in] */                    DWORD             dwCountValid,
	/* [out][dwCountValid] */     VARIANT        ** ppvarDest )
{
	*ppvarDest = new VARIANT[ dwCountValid ];
	if (*ppvarDest == NULL) return E_OUTOFMEMORY;

	VARIANT* pvarDest = *ppvarDest;
	DWORD    i = dwCountAll;
	DWORD    dwCount = 0;
	HRESULT  hr = S_OK;

	while (i--) {
		if (SUCCEEDED( *errors )) {
			VariantInit( pvarDest );
			hr = VariantCopy( pvarDest, pvarSrc );
			if (FAILED( hr )) {
				break;
			}
			dwCount++;
			pvarDest++;
		}
		errors++;
		pvarSrc++;
	}

	if (FAILED( hr )) {                          // Copy of a value failed
		while (dwCount--) {                       // Clear all previous copied values
			pvarDest--;
			VariantClear( pvarDest );
		}
		delete [] (*ppvarDest);
		*ppvarDest = NULL;
	}
	return hr;
}



//=========================================================================
// CopyVQTArrayForValidItems                                      PROTECTED
// -------------------------
//    Builds a copy of an array of VQTs. Only the VQTs with a
//    succeeded code in the corresponding error array are copied.
//
//    For further comments see function CopyVariantArrayForValidItems().
//    This function has the same functionality but uses VQTs instead of
//    Variants.
//=========================================================================
HRESULT DaGenericGroup::CopyVQTArrayForValidItems(
	/* [in] */                    DWORD             dwCountAll,
	/* [in][dwCountAll] */        HRESULT        *  errors,
	/* [in][dwCountAll] */        OPCITEMVQT     *  pVQTsSrc,
	/* [in] */                    DWORD             dwCountValid,
	/* [out][dwCountValid] */     OPCITEMVQT     ** ppVQTsDest )
{
	*ppVQTsDest = new OPCITEMVQT[ dwCountValid ];
	if (*ppVQTsDest == NULL) return E_OUTOFMEMORY;

	OPCITEMVQT* pVQTDest = *ppVQTsDest;
	DWORD       i = dwCountAll;
	DWORD       dwCount = 0;
	HRESULT     hr = S_OK;

	while (i--) {
		if (SUCCEEDED( *errors )) {              // A VQT which needs to be copied

			*pVQTDest = *pVQTsSrc;                 // VQT copy if simple data types

			// Deep copy of the Variant value
			VariantInit( &pVQTDest->vDataValue );
			hr = VariantCopy( &pVQTDest->vDataValue, &pVQTsSrc->vDataValue );
			if (FAILED( hr )) {
				break;
			}
			dwCount++;
			pVQTDest++;
		}
		errors++;
		pVQTsSrc++;
	}

	if (FAILED( hr )) {                          // Copy of a value failed
		while (dwCount--) {                       // Clear all previous copied values
			pVQTDest--;
			VariantClear( &pVQTDest->vDataValue );
		}
		delete [] (*ppVQTsDest);
		*ppVQTsDest = NULL;
	}
	return hr;
}
//DOM-IGNORE-END
