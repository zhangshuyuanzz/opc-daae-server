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
#include "DaComServer.h"
#include "DaGenericServer.h"
#include "DaBaseServer.h"
#include "UtilityFuncs.h"
#include "IClassicBaseNodeManager.h" 

//=========================================================================
// Constructor
//=========================================================================
DaBaseServer::DaBaseServer(void)
{
    created_ = FALSE;
    baseUpdateRate_ = 0;
    name_ = NULL;
    instanceIndex_ = 0;
    InitializeCriticalSection(&criticalSection_);
    InitializeCriticalSection(&serversCriticalSection_);
}

//=========================================================================
// Destructor
//=========================================================================
DaBaseServer::~DaBaseServer()
{
    if (name_) {
        WSTRFree(name_, NULL);
    }

	long     idx;
	DaItemProperty* itemProperty;
	HRESULT  hres = itemProperties_.First(&idx);
	while (SUCCEEDED(hres)) {
		itemProperties_.GetElem(idx, &itemProperty);
		if (itemProperty) {
			itemProperty->~DaItemProperty();
			itemProperty = NULL;
		}
		hres = itemProperties_.Next(idx, &idx);
	}

	itemProperties_.~OpenArray();
	
    DeleteCriticalSection(&serversCriticalSection_);
    DeleteCriticalSection(&criticalSection_);
}


HRESULT DaBaseServer::Create(int instanceIndex,
    LPCWSTR instanceName,
    DWORD updateRate)
{
    _ASSERTE(updateRate > 0);

    if (created_ == TRUE) {
        return E_FAIL;
    }

    if (readWriteLock_.Initialize() == FALSE) {
        return E_FAIL;
    }

    HRESULT hres = SetName(instanceName);
    if (FAILED(hres)) return hres;

    instanceIndex_ = instanceIndex;
    baseUpdateRate_ = updateRate;

    // Create the list of the supported Standard Item Properties.
    hres = SETUP_ITEMPROPERTIES_FROM_MAP();
    if (FAILED(hres)) {
        return hres;
    }

    created_ = TRUE;
    return S_OK;
}


//=========================================================================
// GetItemProperty
// ---------------
//    Gets the Item Property of the specified ID.
//=========================================================================
HRESULT DaBaseServer::GetItemProperty(DWORD dwPropID, DaItemProperty** ppProp)
{
    _ASSERTE(created_);

    long     idx;
    HRESULT  hres = itemProperties_.First(&idx);
    while (SUCCEEDED(hres)) {
        itemProperties_.GetElem(idx, ppProp);
        if (ppProp) {
            if ((*ppProp)->dwPropID == dwPropID)
                return S_OK;                        // Specified property found.
        }
        hres = itemProperties_.Next(idx, &idx);
    }
    *ppProp = NULL;                              // There is no property with the specified ID.
    return hres;
}


//=========================================================================
// FireShutdownRequest
// -------------------
//    Fires a 'Shutdown Request' to all subscribed clients.
//
// Parameters:
//    szReason                A text string witch indicates the reason for
//                            the shutdown or a NULL pointer if there is 
//                            no reason provided.
//                            This parameter is optional. 
//=========================================================================
void DaBaseServer::FireShutdownRequest(LPCWSTR szReason /* = NULL */)
{
    HRESULT           hres;
    long              idx;
    DaGenericServer*   pSrv;

    EnterCriticalSection(&serversCriticalSection_);

    hres = servers_.First(&idx);              // fire request to all clients
    while (SUCCEEDED(hres)) {
        servers_.GetElem(idx, &pSrv);
        if (pSrv && !pSrv->Killed()) {            // inform only active servers
            pSrv->FireShutdownRequest(szReason);
        }
        hres = servers_.Next(idx, &idx);
    }

    LeaveCriticalSection(&serversCriticalSection_);
}

void DaBaseServer::GetClients(int * numClientHandles, void* ** clientHandles, LPWSTR ** clientNames)
{
    HRESULT             hres;
    long                idx;
    DaGenericServer*    pSrv;
    int                 counter = 0;

    EnterCriticalSection(&serversCriticalSection_);

    *numClientHandles = servers_.TotElem();

    if (*numClientHandles > 0)
    {
        void** connectedClients;
        LPWSTR* connectedClientNames;

        connectedClients = new void*[*numClientHandles];
        connectedClientNames = new LPWSTR[*numClientHandles];

        hres = servers_.First(&idx);              
        while (SUCCEEDED(hres)) {
            servers_.GetElem(idx, &pSrv);
            if (pSrv && !pSrv->Killed()) {        
                connectedClients[counter] = pSrv;
                connectedClientNames[counter] = pSrv->m_pCOpcSrv->ClientName.Copy();
                counter++;
            }
            hres = servers_.Next(idx, &idx);
        }
        *clientHandles = connectedClients;
        *clientNames = connectedClientNames;
    }
    LeaveCriticalSection(&serversCriticalSection_);

}

void DaBaseServer::GetGroups(void * clientHandle, int * numGroupHandles, void* ** groupHandles, LPWSTR ** groupNames)
{
    long                idx;
    int                 counter = 0;
    HRESULT             res;
    DaGenericGroup     *group;
    int                 handles = 0;

    DaGenericServer*    pSrv = static_cast<DaGenericServer*>(clientHandle);

    // don't allow deletion or creation of groups while
    // counting
    EnterCriticalSection(&pSrv->m_GroupsCritSec);

    pSrv->RefreshPublicGroups();

    handles = servers_.TotElem();

    void** activeGroups;
    LPWSTR* activeGroupNames;

    activeGroups = new void*[handles];
    activeGroupNames = new LPWSTR[handles];

    // go through all groups!
    res = pSrv->m_GroupList.First(&idx);
    while (SUCCEEDED(res)) {
        pSrv->m_GroupList.GetElem(idx, &group);
        if (group->Killed() == FALSE) {
            activeGroups[counter] = group;
            activeGroupNames[counter] = new WCHAR[wcslen(group->m_Name) + 1];
            if (activeGroupNames[counter]) {
                wcscpy(activeGroupNames[counter], group->m_Name);
            }
            counter++;
        }
        res = pSrv->m_GroupList.Next(idx, &idx);
    }
    if (counter > 0) {
        *numGroupHandles = handles;
        *groupHandles = activeGroups;
        *groupNames = activeGroupNames;
    }

    LeaveCriticalSection(&pSrv->m_GroupsCritSec);
}

void DaBaseServer::GetGroupState(void * groupHandle, IClassicBaseNodeManager::DaGroupState * groupInformation)
{
    DaGenericGroup*    group = static_cast<DaGenericGroup*>(groupHandle);

    groupInformation->GroupName = group->m_Name;
    groupInformation->ClientGroupHandle = group->m_hClientGroupHandle;
    groupInformation->DataChangeEnabled = group->m_fCallbackEnable;
    groupInformation->LocaleId = group->m_dwLCID;
    groupInformation->PercentDeadband = group->m_PercentDeadband;
    groupInformation->UpdateRate = group->m_RevisedUpdateRate;
}

void DaBaseServer::GetItemStates(void * groupHandle, int * numDaItemStates, IClassicBaseNodeManager::DaItemState* * daItemStates)
{
    DaGenericGroup*    group = static_cast<DaGenericGroup*>(groupHandle);
    HRESULT        hrRet;                       // Returned result
    DaGenericItem* item;
    DaDeviceItem**  ppDItems = NULL;            // Array of Device Item Pointers
    OPCITEMSTATE*  pItemStates = NULL;          // Array of OPC Item States
    OPCITEMVQT*    pVQTsToWrite = NULL;         // Array of Values, Qualities and TimeStamps to write
    DWORD          dwCountValid = 0;            // Number of valid items
    long            ih, newih;

    DWORD dwCount = group->m_oaItems.TotElem();

    if (dwCount == 0)
        return;

    ppDItems = new DaDeviceItem*[dwCount];    // create item array object
    if (ppDItems == NULL) {
        return;
    }

    // we do not allow changes to the item list
    // while building arrays of device items, etc.
    EnterCriticalSection(&(group->m_ItemsCritSec));

    hrRet = group->m_oaItems.First(&ih);
    while (SUCCEEDED(hrRet)) {
        group->m_oaItems.GetElem(ih, &item);

        if (item->get_Active() == TRUE) {
            if (item->AttachDeviceItem(&ppDItems[dwCountValid]) > 0) {
                // DeviceItem attached
                dwCountValid++;
            }
        }
        hrRet = group->m_oaItems.Next(ih, &newih);
        ih = newih;
    }

    LeaveCriticalSection(&group->m_ItemsCritSec);

    if (dwCountValid > 0) {
		*numDaItemStates = dwCountValid;

		*daItemStates = new IClassicBaseNodeManager::DaItemState[dwCountValid];    // create item array object
		if (daItemStates == NULL) {
            return;
        }

        for (DWORD i = 0; i<dwCountValid; i++) {
            LPWSTR name;
            ppDItems[i]->get_ItemIDCopy(&name);
			daItemStates[i]->ItemName = name;
            ppDItems[i]->get_AccessPath(&name);
			daItemStates[i]->AccessPath = name;
			ppDItems[i]->get_AccessRights(&daItemStates[i]->AccessRights);
			daItemStates[i]->DeviceItemHandle = ppDItems[i];
            if (ppDItems[i] != NULL) {
                ppDItems[i]->Detach();
            }
        }
    }
    delete[] ppDItems;
}

//=========================================================================
// Standard Revise Update Rate
// ---------------------------
// this standard implementation can be overriden!
//=========================================================================
HRESULT DaBaseServer::ReviseUpdateRate(
    DWORD RequestedUpdateRate,
    DWORD *RevisedUpdateRate)
{
    DWORD n;
    DWORD temp;

    _ASSERTE(created_ == TRUE);

    EnterCriticalSection(&criticalSection_);
    temp = baseUpdateRate_;
    LeaveCriticalSection(&criticalSection_);

    _ASSERTE(temp > 0);

    // round 'RequestedUpdateRate'
    // to the next multiple of  'baseUpdateRate_'
    n = (RequestedUpdateRate + temp - 1) / temp;
    if (n == 0) {       // set to minimum rate
        n = 1;
    }
    *RevisedUpdateRate = n * temp;
    return S_OK;
}


//=========================================================================
// Notify Update to Server instances
// ---------------------------------
//=========================================================================
HRESULT DaBaseServer::UpdateServerClassInstances()
{
    long              idx, nidx;
    HRESULT           res;
    DaGenericServer    *serv;

    _ASSERTE(created_ == TRUE);

    EnterCriticalSection(&serversCriticalSection_);

    // inform all server threads to update
    // their respective clients 
    res = servers_.First(&idx);
    while (SUCCEEDED(res)) {
        servers_.GetElem(idx, &serv);

        if (serv->Killed() == FALSE) {
            // inform only active servers (not zombies)
            SetEvent(serv->m_hUpdateEvent);
        }

        res = servers_.Next(idx, &nidx);
        idx = nidx;
    }
    LeaveCriticalSection(&serversCriticalSection_);

    return S_OK;
}




//=========================================================================
// Server object enters
// --------------------
// adds a server instance to this class handler's list;
// this method is called when a client connects 
//=========================================================================
HRESULT DaBaseServer::AddServerToList(DaGenericServer *pServer)
{
    HRESULT           res;
    long              idx, nidx;
    DaGenericServer    *serv;

    _ASSERTE(created_ == TRUE);
    _ASSERTE(pServer != NULL);

    EnterCriticalSection(&serversCriticalSection_);

    // check if not already in list
    res = servers_.First(&idx);
    while (SUCCEEDED(res)) {

        servers_.GetElem(idx, &serv);
        if (serv == pServer) {
            // it's already in list
            LeaveCriticalSection(&serversCriticalSection_);
            return E_FAIL;
        }

        res = servers_.Next(idx, &nidx);
        idx = nidx;
    }

    // get a free index in the open array
    idx = servers_.New();

    // insert the server inthe list
    res = servers_.PutElem(idx, pServer);

    LeaveCriticalSection(&serversCriticalSection_);

    return res;
}




//=========================================================================
// Server object leaves
// --------------------
// removes a server instance from this class handler's list;
// this method is called when a client has properly disconnected,
// that is freed all the COM interfaces it owns
//=========================================================================
HRESULT DaBaseServer::RemoveServerFromList(DaGenericServer *pServer)
{
    HRESULT           res;
    long              idx, nidx;
    DaGenericServer    *serv;

    _ASSERTE(created_ == TRUE);
    _ASSERTE(pServer != NULL);

    EnterCriticalSection(&serversCriticalSection_);

    res = servers_.First(&idx);
    while (SUCCEEDED(res)) {

        servers_.GetElem(idx, &serv);
        if (serv == pServer) {
            // found => remove from the list
            servers_.PutElem(idx, NULL);

            LeaveCriticalSection(&serversCriticalSection_);
            return S_OK;
        }

        res = servers_.Next(idx, &nidx);
        idx = nidx;
    }

    LeaveCriticalSection(&serversCriticalSection_);

    return E_FAIL;
}


HRESULT DaBaseServer::QueryInterfaceOnServerFromList(
    SRVINSTHANDLE  server, // Specify the server
    BOOL     &     serverExist,
    REFIID         refId,
    LPVOID   *     unkown)
{
    DaGenericServer*   pSrvInst = NULL;
    HRESULT           hres = E_FAIL;
    long              idx;

    _ASSERTE(unkown != NULL);          // Must not be NULL

    *unkown = NULL;                      // Initialize results
    serverExist = FALSE;
    // protect server list while searching
    EnterCriticalSection(&serversCriticalSection_);

    hres = servers_.First(&idx);
    while (hres == S_OK) {             // search in all instances
        servers_.GetElem(idx, &pSrvInst);

        if (pSrvInst) {                  // server instance found

            if (pSrvInst->m_pCOpcSrv->serverInstanceHandle_ == server) {
                // server has the requested Instance Handle
                serverExist = TRUE;
                hres = ((IOPCServer *)pSrvInst->m_pCOpcSrv)->QueryInterface(refId, unkown);
                break;
            }
        }
        // get next server instance
        hres = servers_.Next(idx, &idx);
    }
    // release server list protection
    LeaveCriticalSection(&serversCriticalSection_);

    return hres;
}




//=========================================================================
// Sets the base update rate of the server
// ---------------------------------------
//=========================================================================
HRESULT DaBaseServer::SetBaseUpdateRate(DWORD BaseUpdateRate)
{
    _ASSERTE(created_ == TRUE);

    if (BaseUpdateRate == 0) {
        return E_FAIL;
    }

    EnterCriticalSection(&criticalSection_);
    baseUpdateRate_ = BaseUpdateRate;
    LeaveCriticalSection(&criticalSection_);
    return S_OK;
}



//=========================================================================
// gets the base update rate of the server
// ---------------------------------------
//=========================================================================
DWORD DaBaseServer::GetBaseUpdateRate(void)
{
    DWORD temp;

    _ASSERTE(created_ == TRUE);

    EnterCriticalSection(&criticalSection_);
    temp = baseUpdateRate_;
    LeaveCriticalSection(&criticalSection_);

    return temp;
}



//=========================================================================
// Set the Name of the Server
// --------------------------
//    Sets a new Server Name.
//    The access to the Server Name is protected with a critical section.
//
//=========================================================================
HRESULT DaBaseServer::SetName(LPCWSTR pSrvName)
{
    if (pSrvName == NULL) {
        return E_INVALIDARG;
    }

    HRESULT hres = S_OK;
    EnterCriticalSection(&criticalSection_);

    LPWSTR pOldName = name_;

    name_ = WSTRClone(pSrvName, NULL);
    if (name_ == NULL) {
        name_ = pOldName;                        // restore the old name
        hres = E_OUTOFMEMORY;
    }
    else {
        if (pOldName) {
            WSTRFree(pOldName, NULL);            // release the old name
        }
    }

    LeaveCriticalSection(&criticalSection_);
    return hres;
}



//=========================================================================
// Get the Name of the Server
// --------------------------
//    Returns a copy of the Server Name.
//    The access to the Server Name is protected with a critical section.
//
//       pIMalloc          Allocates string with new() if NULL; otherwise
//                         with pIMalloc->Alloc().
//
//=========================================================================
HRESULT DaBaseServer::GetName(LPWSTR *pSrvName, IMalloc *pIMalloc /*=NULL*/)
{
    if (pSrvName == NULL) {
        return E_INVALIDARG;
    }

    HRESULT hres = S_OK;
    EnterCriticalSection(&criticalSection_);

    if (name_) {
        *pSrvName = WSTRClone(name_, pIMalloc);
        if (*pSrvName == NULL) {
            hres = E_OUTOFMEMORY;
        }
    }
    else {
        *pSrvName = NULL;
    }
    LeaveCriticalSection(&criticalSection_);
    return hres;
}

//DOM-IGNORE-END
