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
#include <process.h>
#include <cstdio>
#include "DaGenericServer.h"
#include "UtilityFuncs.h"
#include "DaComServer.h"
#include "DaPublicGroup.h"

//=========================================================================
// Constructor
//=========================================================================
DaGenericServer::DaGenericServer()
{
    m_Created = FALSE;

    m_pServerHandler = nullptr
    ;
    m_pCOpcSrv = nullptr;
    m_FilterCriteria = nullptr;
    m_hUpdateThread = nullptr;

    try
    {
        InitializeCriticalSection(&m_CritSec);
        InitializeCriticalSection(&m_GroupsCritSec);
        InitializeCriticalSection(&m_UpdateRateCritSec);
    }
    catch (...)
    {
    }

    CriticalSectionCOMGroupList.Initialize();
}


//=========================================================================
// Initializer
//=========================================================================
HRESULT DaGenericServer::Create(DaBaseServer* pServerClassHandler, DaComBaseServer* pOPCServerBase)
{
    HRESULT res;

    _ASSERTE(pServerClassHandler != NULL);
    _ASSERTE(pOPCServerBase != NULL);

    if (m_Created == TRUE) {
        // already created
        return E_FAIL;
    }

    // fill position zero with no group
    // so that zero won't be used as a valid group handle!
    res = m_GroupList.PutElem(0, nullptr);
    if (FAILED(res)) {
        return res;
    }
    // actual filter criteria
    m_FilterCriteria = SysAllocString(L"");
    if (m_FilterCriteria == nullptr) {
        return E_OUTOFMEMORY;
    }
    // server class handler
    m_pServerHandler = pServerClassHandler;

    // OPC server class base
    m_pCOpcSrv = pOPCServerBase;

    // get the actual update rate 
    // from which the groups 'tick count'  information will be calculated
    m_ActualBaseUpdateRate = pServerClassHandler->GetBaseUpdateRate();

    m_ToKill = FALSE;
    m_RefCount = 0;

    // default enumeration type
    m_Enum_Type = OPC_GROUPNAME_ENUM;

    CoFileTimeNow(&m_StartTime);
    memset(&m_LastUpdateTime, 0, sizeof(m_LastUpdateTime));  // 1.Jan 1970, 00:00:00

       // automation group enumeration settings
    m_Enum_Scope = OPC_ENUM_ALL;
    m_Enum_Type = OPC_GROUP_ENUM;

    // automation browse server address space enumerartions
    m_EnumeratingItemIDs = TRUE;
    // if m_EnumeratingItemIDs == FALSE => enumerating Access Paths
    m_BrowseItemIDForAccessPath = nullptr;

    m_BrowseFilterType = OPC_FLAT;
    m_DataTypeFilter = VT_EMPTY;
    m_AccessRightsFilter = 0;

    // Create the update event
    TCHAR szEventName[30];                       // Name must be unique
    _stprintf_s(szEventName,  30, _T("UpdateEvent_0x%X"), this);

    m_hUpdateEvent = CreateEvent(nullptr,          // Cannot inherit handle.
                                 FALSE,         // Automatically reset.
                                 FALSE,         // Nonsignaled initial state.
                                 szEventName);

    if (m_hUpdateEvent == nullptr) {
        res = E_FAIL;                             // Cannot create an event.
        goto CreateExit2;
    }

    res = CreateWaitForUpdateThread();
    if (FAILED(res)) {
        goto CreateExit3;
    }

    res = m_BrowseData.Create(pServerClassHandler);
    if (FAILED(res)) {
        goto CreateExit4;
    }

    m_Created = TRUE;

    res = m_pServerHandler->AddServerToList(this);
    if (FAILED(res)) {
        goto CreateExit4;
    }

    Attach();                                     // Now generic Server is used
    return S_OK;

CreateExit4:
    m_Created = FALSE;
    KillWaitForUpdateThread();

CreateExit3:
    CloseHandle(m_hUpdateEvent);
    m_hUpdateEvent = nullptr;

CreateExit2:
    SysFreeString(m_FilterCriteria);
    m_FilterCriteria = nullptr;

    return res;
}

//=========================================================================
// Destructor
//=========================================================================
DaGenericServer::~DaGenericServer(void)
{

    if (m_Created == TRUE) {

        m_pServerHandler->RemoveServerFromList(this);

        CloseHandle(m_hUpdateEvent);
        m_hUpdateEvent = nullptr;
    }

    if (m_FilterCriteria) {
        SysFreeString(m_FilterCriteria);
    }

    DeleteCriticalSection(&m_GroupsCritSec);
    DeleteCriticalSection(&m_CritSec);
    DeleteCriticalSection(&m_UpdateRateCritSec);
}


//=====================================================================================
//  Attach to class by incrementing the RefCount
//  Returns the current RefCount
//  or -1 if the ToKill flag is set
//=====================================================================================
int  DaGenericServer::Attach(void)
{
    long i;

    EnterCriticalSection(&m_CritSec);
    if (m_ToKill) {
        return -1;
    }
    i = m_RefCount++;
    LeaveCriticalSection(&m_CritSec);
    return i;
}


//=====================================================================================
//  Detach from class by decrementing the RefCount.
//  Kill the class instance if request pending.
//  Returns the current RefCount or -1 if it was killed.
//=====================================================================================
int  DaGenericServer::Detach(void)
{
    long i;

    EnterCriticalSection(&m_CritSec);

    _ASSERTE(m_RefCount > 0);

    m_RefCount--;

    if ((m_RefCount == 0)                      // not referenced
        && (m_ToKill == TRUE)) {            // kill request

        LeaveCriticalSection(&m_CritSec);
        delete this;                              // remove from memory
        return -1;
    }
    i = m_RefCount;
    LeaveCriticalSection(&m_CritSec);
    return i;
}


//=====================================================================================
//  Kill the class instance.
//  If the RefCount is > 0 then only the kill request flag is set.
//  Returns the current RefCount or -1 if it was killed.
//=====================================================================================
int  DaGenericServer::Kill(BOOL WithDetach)
{

    long           i;
    HRESULT        res;
    DaGenericGroup* theGroup;

    KillWaitForUpdateThread();

    EnterCriticalSection(&m_CritSec);

    // kill the groups of this server
    EnterCriticalSection(&m_GroupsCritSec);
    res = m_GroupList.First(&i);
    while (SUCCEEDED(res)) {
        res = m_GroupList.GetElem(i, &theGroup);
        if (theGroup->Killed() == FALSE) {
            theGroup->Kill();
        }
        res = m_GroupList.Next(i, &i);
    }
    LeaveCriticalSection(&m_GroupsCritSec);

    // !!!??? ev. force delete of all COM groups!
    // so that the server will shutdown even if there are 
    // bad clients

    //
    if (WithDetach) {
        if (m_RefCount > 0)
            --m_RefCount;
    }
    if (m_RefCount) {            // still referenced
        m_ToKill = TRUE;          // set kill request flag
    }
    else {
        LeaveCriticalSection(&m_CritSec);
        delete this;
        return -1;
    }
    i = m_RefCount;
    LeaveCriticalSection(&m_CritSec);
    return i;
}

//=========================================================================
// Killed()
// tells whether the class instance was killed (with Kill()) or not
//=========================================================================
BOOL DaGenericServer::Killed(void)
{
    BOOL b;

    EnterCriticalSection(&m_CritSec);
    b = m_ToKill;
    LeaveCriticalSection(&m_CritSec);
    return b;
}




//=========================================================================
// GetGenericGroup
//  returns the group associated with handle
//  if handle not valid (maybe group was deleted)
//  then return NULL
//=========================================================================
HRESULT DaGenericServer::GetGenericGroup(
    long            ServerGroupHandle,
    DaGenericGroup **group)
{
    long res;
    long pgh;

    DaPublicGroup         *pg;
    DaPublicGroupManager  *pgHandler;

    pgHandler = &(m_pServerHandler->publicGroups_);

    EnterCriticalSection(&m_GroupsCritSec);
    res = m_GroupList.GetElem(ServerGroupHandle, group);
    if (FAILED(res)) {
        res = OPC_E_INVALIDHANDLE;                // invalid handle
        goto GetGenericGroupExit0;
    }
    if ((*group)->Killed() == TRUE) {
        res = E_FAIL;                             // group has to be removed
        goto GetGenericGroupExit0;
    }

    // check if group is public
    if ((*group)->GetPublicInfo(&pgh) == TRUE) {
        // check if this public group still exists
        // !!!??? maybe this instruction can create DEADLOCK! control!
        res = pgHandler->GetGroup(pgh, &pg);
        if (FAILED(res)) {
            // public group isn't there
            // remove DaGenericGroup from server list and kill
            (*group)->Kill();
            res = OPC_E_INVALIDHANDLE;
            goto GetGenericGroupExit0;
        }
        pgHandler->ReleaseGroup(pgh);
    }
    // increment ref count so that group cannot be deleted
    // during access
    (*group)->Attach();
    res = S_OK;

GetGenericGroupExit0:
    LeaveCriticalSection(&m_GroupsCritSec);
    return res;
}




//=========================================================================
// ReleaseGenericGroup
//
//=========================================================================
HRESULT DaGenericServer::ReleaseGenericGroup(
    long ServerGroupHandle)
{
    DaGenericGroup *group;
    long res;

    EnterCriticalSection(&m_GroupsCritSec);
    res = m_GroupList.GetElem(ServerGroupHandle, &group);
    if (FAILED(res)) {
        res = OPC_E_INVALIDHANDLE;
        goto ReleaseGenericGroupExit0;
    }
    if (group->Detach() < 0) {         // killed due to request
        res = m_GroupList.PutElem(ServerGroupHandle, nullptr);
    }
    res = S_OK;

ReleaseGenericGroupExit0:
    LeaveCriticalSection(&m_GroupsCritSec);
    return res;
}


//=========================================================================
// RemoveGenericGroup
//=========================================================================
HRESULT DaGenericServer::RemoveGenericGroup(
    long ServerGroupHandle)
{
    HRESULT           hres = S_OK;
    DaGenericGroup*    pGGroup;

    EnterCriticalSection(&m_GroupsCritSec);

    m_GroupList.GetElem(ServerGroupHandle, &pGGroup);
    if ((pGGroup == nullptr) || pGGroup->Killed()) {
        hres = OPC_E_INVALIDHANDLE;
    }
    else {
        pGGroup->Kill();
    }

    LeaveCriticalSection(&m_GroupsCritSec);
    return hres;
}


//=========================================================================
// SearchGroup
//  searches a group between public or private groups
//  this function expects external 
//  collision control through the m_GroupsCritSec critical section
//=========================================================================
HRESULT DaGenericServer::SearchGroup(
    BOOL           Public,
    LPCWSTR        theName,
    DaGenericGroup  **theGroup,
    long           *sgh)
{
    long           i, size;
    long           pgh;
    DaGenericGroup  *group;
    HRESULT        res;

    if (theName != nullptr) {
        size = m_GroupList.Size();
        for (i = 0; i < size; i++) {
            res = m_GroupList.GetElem(i, &group);
            if ((SUCCEEDED(res))
                && (group->Killed() == FALSE)
                && (group->GetPublicInfo(&pgh) == Public)
                && (wcscmp(theName, group->m_Name) == 0)) {
                *sgh = i;
                *theGroup = group;
                return S_OK;
            }
        }
    }
    *theGroup = nullptr;
    return E_FAIL;
}



//=========================================================================
// ChangePrivateGroupName
//=========================================================================
HRESULT DaGenericServer::ChangePrivateGroupName(
    DaGenericGroup  *group,
    LPWSTR         Name
)
{

    HRESULT        res;
    long           sgh;
    long           pgh;
    DaGenericGroup  *foundGroup;

    if (group->GetPublicInfo(&pgh) == TRUE) {
        res = OPC_E_PUBLIC;
        goto ChangePrivateGroupNameExit0;
    }

    if (!Name || *Name == L'\0') {
        LOGFMTE("ChangePrivateGroupName() failed with invalid argument(s): Name is a NULL-Pointer or a NULL-String");
        res = E_INVALIDARG;
        goto ChangePrivateGroupNameExit0;
    }

    EnterCriticalSection(&m_GroupsCritSec);

    res = SearchGroup(FALSE, Name, &foundGroup, &sgh);
    if (SUCCEEDED(res)) {
        if (foundGroup == group) {                // Also succeeded bause it's the name
            goto ChangePrivateGroupNameExit1;      // of the own group.
        }
        res = OPC_E_DUPLICATENAME;
        goto ChangePrivateGroupNameExit1;
    }

    res = group->set_Name(Name);

ChangePrivateGroupNameExit1:
    LeaveCriticalSection(&(m_GroupsCritSec));

ChangePrivateGroupNameExit0:
    return res;
}




//=================================================================
// GetCOMGroup                              
//    Gets an interface of the associated COM group. If there
//    is no COM object then it will be created.
//=================================================================
HRESULT DaGenericServer::GetCOMGroup(
    DaGenericGroup*    pGGroup,
    REFIID            riid,
    LPUNKNOWN*        ppUnk)
{
    HRESULT                 hr;
    CComObject<DaGroup>*  pCOMGroup;

    *ppUnk = nullptr;

    CriticalSectionCOMGroupList.BeginReading();                        // Lock reading COM list

    hr = m_COMGroupList.GetElem(pGGroup->m_hServerGroupHandle, &pCOMGroup);
    if (SUCCEEDED(hr)) {                       // there already is a COM group
        hr = pCOMGroup->QueryInterface(riid, (LPVOID *)ppUnk);
        CriticalSectionCOMGroupList.EndReading();                      // Unlock reading COM list
        return hr;
    }
    CriticalSectionCOMGroupList.EndReading();                      // Unlock reading COM list


    CriticalSectionCOMGroupList.BeginWriting();                        // Lock writing COM list

    hr = m_COMGroupList.GetElem(pGGroup->m_hServerGroupHandle, &pCOMGroup);
    if (SUCCEEDED(hr)) {                       // there already is a COM group
        hr = pCOMGroup->QueryInterface(riid, (LPVOID *)ppUnk);
    }
    else {
        // create the COM Group
        hr = CComObject<DaGroup>::CreateInstance(&pCOMGroup);
        if (FAILED(hr)) {
            goto GetCOMGroupExit1;
        }
        // initialize the COM group
        hr = pCOMGroup->Create(this, pGGroup->m_hServerGroupHandle, pGGroup);
        if (FAILED(hr)) {                       // failed
            delete pCOMGroup;
            goto GetCOMGroupExit1;
        }
        // get the requested interface
        hr = pCOMGroup->QueryInterface(riid, (LPVOID *)ppUnk);
        if (FAILED(hr)) {
            delete pCOMGroup;
            goto GetCOMGroupExit1;
        }
        // store the created COM group in the list
        hr = m_COMGroupList.PutElem(pGGroup->m_hServerGroupHandle, pCOMGroup);
        if (FAILED(hr)) {                       // probably out of memory
            pCOMGroup->Release();                  // release interface and destroy object
            goto GetCOMGroupExit1;
        }
        // all succeeded
    }

GetCOMGroupExit1:
    CriticalSectionCOMGroupList.EndWriting();                      // Unlock writing COM list
    return hr;
}


//=================================================================
//  Internal Enumeration Helper Method to get an array of 
//  Interface-Ptrs to the requested groups.
//
//  Parameters:
//  riid:  Interface type requested: IID_IEnumUnknown or IID_IEnumString
//  scope: Indicates the class of groups to be enumerated
//         OPC_ENUM_PRIVATE_CONNECTIONS 
//               enumerates the private groups the client is connected to.
//         OPC_ENUM_PUBLIC_CONNECTIONS 
//               enumerates the public groups the client is connected to.
//         OPC_ENUM_ALL_CONNECTIONS 
//               enumerates all of the groups the client is connected to.
//         OPC_ENUM_PRIVATE 
//               enumerates all of the private groups created by the client (whether connected or not)
//         OPC_ENUM_PUBLIC 
//               enumerates all of the public groups available in the server (whether connected or not)
//         OPC_ENUM_ALL 
//               enumerates all private groups and all public groups (whether connected or not)
//
//  GroupList:   returned array of Interface Ptrs
//
//  GroupCount:  Number of returned ptrs
//
//=================================================================
//-----------------------------------------------------------------
// Notes:
//    m_GroupsCritSec locking is done outside!
//-----------------------------------------------------------------
HRESULT DaGenericServer::GetGrpList(
    REFIID         riid,
    OPCENUMSCOPE   scope,
    LPUNKNOWN    **GroupList,
    int           *GroupCount)
{
    long            sh, size, count;
    LPUNKNOWN       pUnk;
    DaGenericGroup  *group;
    CComObject<DaGroup>   *pCOMGroup;
    HRESULT         res;
    BOOL            bPublic;
    long pgh;

    _ASSERTE((riid == IID_IEnumUnknown) || (riid == IID_IEnumString));

    if ((scope != OPC_ENUM_PRIVATE)
        && (scope != OPC_ENUM_ALL)
        && (scope != OPC_ENUM_PUBLIC)
        && (scope != OPC_ENUM_PRIVATE_CONNECTIONS)
        && (scope != OPC_ENUM_ALL_CONNECTIONS)
        && (scope != OPC_ENUM_PUBLIC_CONNECTIONS)) {
        goto GetGrpListExit0;
    }

    // First public group list must be refreshed
    RefreshPublicGroups();

    size = m_GroupList.TotElem();       // highest group index
    if (size == 0) {                   // no groups
        res = S_OK;
        goto GetGrpListExit0;
    }
    *GroupList = new LPUNKNOWN[size];
    if (*GroupList == nullptr) {
        res = E_OUTOFMEMORY;
        goto GetGrpListExit0;
    }

    // go through all generic groups
    // and if enumerating connections
    // look if there is the corresponding COM group
    count = 0;
    res = m_GroupList.First(&sh);
    while (SUCCEEDED(res)) {
        res = m_GroupList.GetElem(sh, &group);
        if ((SUCCEEDED(res))
            && (group->Killed() == FALSE)) {
            bPublic = group->GetPublicInfo(&pgh);

            CriticalSectionCOMGroupList.BeginReading();                  // Lock reading COM list

            res = m_COMGroupList.GetElem(sh, &pCOMGroup);

            if (((scope == OPC_ENUM_PRIVATE) && (bPublic == FALSE))
                || ((scope == OPC_ENUM_PUBLIC) && (bPublic == TRUE))
                || (scope == OPC_ENUM_ALL)
                || ((SUCCEEDED(res)) && (scope == OPC_ENUM_PRIVATE_CONNECTIONS) && (bPublic == FALSE))
                || ((SUCCEEDED(res)) && (scope == OPC_ENUM_PUBLIC_CONNECTIONS) && (bPublic == TRUE))
                || ((SUCCEEDED(res)) && (scope == OPC_ENUM_ALL_CONNECTIONS))) {

                if (riid == IID_IEnumUnknown) {
                    // create a new COM group for this group
                    // and get a new interface for this group
                    res = GetCOMGroup(group, IID_IUnknown, &pUnk);
                    if (FAILED(res)) {           // error, ignore group
                        goto GetGrpListExit1;
                    }
                    (*GroupList)[count] = pUnk;

                }
                else {         // IID_IEnumString
                    (*GroupList)[count] = (LPUNKNOWN)WSTRClone(group->m_Name, nullptr);
                }

                // count the group
                ++count;
            }
            CriticalSectionCOMGroupList.EndReading();                // Unlock reading COM list

        }
        res = m_GroupList.Next(sh, &sh);
    }

    *GroupCount = count;
    return S_OK;


GetGrpListExit1:
    // release the COM items
    CriticalSectionCOMGroupList.EndReading();                      // Unlock reading COM list
    count--;
    while (count >= 0) {
        (*GroupList)[count]->Release();
        count--;
    }
    delete GroupList;

GetGrpListExit0:
    *GroupList = nullptr;            // init results
    *GroupCount = 0;
    return res;
}



//=================================================================
// free the group list allocated with GetGrpList
//=================================================================
HRESULT DaGenericServer::FreeGrpList(
    REFIID      riid,
    LPUNKNOWN  *GroupList,
    int         count)
{
    int   sh;

    if ((count == 0) && (GroupList == nullptr)) {
        return S_OK;
    }

    if (riid == IID_IEnumString) {
        // free the strings
        for (sh = 0; sh < count; sh++) {
            //delete ( LPOLESTR ) ( GroupList[sh] );     
            WSTRFree((LPWSTR)(GroupList[sh]), nullptr);
        }
    }
    else {
        sh = 0;
        while (sh < count) {
            GroupList[sh]->Release();
            sh++;
        }
    }
    delete GroupList;               // free the list
    return S_OK;
}


//=================================================================
// refreshes m_GroupList deleting no more existing public groups
// and inserting newly created ones!
// This proc assumes m_GroupList is already locked by the caller
//=================================================================
HRESULT DaGenericServer::RefreshPublicGroups()
{

    long                 sh;
    long                 res;
    DaGenericGroup        *group, *theGroup;
    DaPublicGroupManager  *pgHandler;
    DaPublicGroup         *pPG;
    long                 pgh;
    int                  found;
    long                 thePGh;
    BOOL                 bPublic;

    // start to enumerate all the groups
    // except 0 which is the invalid handle

    pgHandler = &(m_pServerHandler->publicGroups_);

    res = m_GroupList.Next(0, &sh);
    while (SUCCEEDED(res)) {
        // found a group
        res = m_GroupList.GetElem(sh, &group);
        if (group->Killed() == FALSE) {
            bPublic = group->GetPublicInfo(&pgh);
            if (bPublic == TRUE) {
                res = pgHandler->GetGroup(pgh, &pPG);
                if (FAILED(res)) {
                    // public group isn't there
                    // remove DaGenericGroup from server list and delete
                    m_GroupList.PutElem(sh, nullptr);
                    // invalid handle or group has to be removed
                    group->Kill();
                }
                else {
                    // there is a public group
                    pgHandler->ReleaseGroup(pgh);
                }
            }
        }

        res = m_GroupList.Next(sh, &sh);
    }

    // search in the public group list
    // if there is some group not in the servers m_GroupList create and add it!

    res = pgHandler->First(&pgh);
    while (SUCCEEDED(res)) {

        // get and nail the public group
        res = pgHandler->GetGroup(pgh, &pPG);
        if (SUCCEEDED(res)) {

            found = FALSE;
            res = m_GroupList.Next(0, &sh);
            while ((SUCCEEDED(res)) && (found == FALSE)) {
                res = m_GroupList.GetElem(sh, &group);

                bPublic = group->GetPublicInfo(&thePGh);
                if ((bPublic == TRUE) && (thePGh == pgh)) {
                    found = TRUE;
                }

                res = m_GroupList.Next(sh, &sh);
            }

            if (found == FALSE) {
                // there is a public group without a corresponding DaGenericGroup: 
                // so we have to create the DaGenericGroup for this client

                // create the group from the DaPublicGroup
                theGroup = new DaGenericGroup();
                if (theGroup == nullptr) {
                    // unnail the public group
                    pgHandler->ReleaseGroup(pgh);
                    res = E_OUTOFMEMORY;
                    goto RefreshPublicGroupsExit0;
                }
                // create the group from the DaPublicGroup
                res = theGroup->Create(
                    this,
                    pPG->m_Name,
                    pPG->m_InitialActive,
                    &pPG->m_InitialTimeBias,
                    pPG->m_InitialUpdateRate,
                    0,  // client group handle: default
                    &pPG->m_PercentDeadband,
                    0,      // LCID
                    TRUE,   // Public group
                    pgh,
                    &sh
                );

                if (FAILED(res)) {
                    delete theGroup;
                    // unnail the public group
                    pgHandler->ReleaseGroup(pgh);
                    res = E_OUTOFMEMORY;
                    goto RefreshPublicGroupsExit0;
                }
            }
            // unnail the public group
            pgHandler->ReleaseGroup(pgh);

        } // couldn't get the group

           // get the next in the list
        res = pgHandler->Next(pgh, &pgh);
    }

    res = S_OK;

RefreshPublicGroupsExit0:
    return res;

}



//=================================================================
// GetGroupUniqueName
// old: "returns a unique name for a private or public group"
// V1.0a: returns a unique name among private and public groups!
// the server group handle is used as a starting point ( to build the name )
// This proc assumes m_GroupList is already locked by the caller
//=================================================================
HRESULT DaGenericServer::GetGroupUniqueName(long ServerGroupHandle, WCHAR **theName, BOOL bPublic)
{
    WCHAR    wc[15];
    long     i;
    long     sgh;
    DaGenericGroup  *foundGroup;

    i = ServerGroupHandle;
    *theName = new wchar_t[32];
    if (*theName == nullptr) {
        return E_OUTOFMEMORY;
    }

    WCHAR defaultGroupName[] = L"Group_";

    wcscpy_s(*theName, wcslen(defaultGroupName)+1, defaultGroupName);
    _itow_s(i, wc, 10);
    wcscat(*theName, wc);

    while ((SUCCEEDED(SearchGroup(TRUE, *theName, &foundGroup, &sgh)))
        || (SUCCEEDED(SearchGroup(FALSE, *theName, &foundGroup, &sgh)))) {
        // it's already a public or private group name
        // generate another
        wcscpy_s(*theName, wcslen(*theName), defaultGroupName);
        _itow_s(i, wc, 10);
        wcscat(*theName, wc);
        i++;
    }

    return S_OK;
}



//=================================================================================
// Thread waiting for Update Event from server class handler
// ---------------------------------------------------------
//=================================================================================
unsigned __stdcall WaitForUpdateThread(void* pArg)
{
    DaGenericServer* serv = static_cast<DaGenericServer *>(pArg);
    _ASSERTE(serv != NULL);

    long              idx;
    HRESULT           res;
    DaGenericGroup    *group;

    while (serv->m_UpdateThreadToKill == FALSE) {
        // wait ServerClassHandler sends me update event
        WaitForSingleObject(serv->m_hUpdateEvent, INFINITE);

        // check if base update rate has changend
        // and if so recalc tick counts of groups
        serv->CheckBaseUpdateRateChanged();

        // update all groups belonging to this server instance
        // all access to the group list is protected by crit sec
        EnterCriticalSection(&serv->m_GroupsCritSec);
        res = serv->m_GroupList.First(&idx);

        while (SUCCEEDED(res)) {
            // get and nail the group
            res = serv->GetGenericGroup(idx, &group);

            LeaveCriticalSection(&serv->m_GroupsCritSec);

            if (SUCCEEDED(res)) {

                if (group->GetActiveState() == TRUE) {
                    // only active groups enter into account for update
                    EnterCriticalSection(&group->m_UpdateRateCritSec);
                    // recalc ticks if base update rate changed
                    serv->RecalcTicks(group);
                    // update only active groups
                    group->m_TickCount--;

                    if (group->m_TickCount == 0) {
                        // restart tick counter
                        group->m_TickCount = group->m_Ticks;
                        LeaveCriticalSection(&group->m_UpdateRateCritSec);

                        // it's group turn to update
                        res = group->UpdateNotify();

                    }
                    else {
                        LeaveCriticalSection(&group->m_UpdateRateCritSec);
                    }

                    //
                    // Keep Alive
                    //
                    if (group->m_fCallbackEnable) {

                        group->m_csKeepAlive.Lock();
                        if (group->KeepAliveTime()) {    // Activated keep-alive callbacks

                            CComObject<DaGroup>* pCOMGroup = nullptr;
                            IUnknown** ppCallback = nullptr;

                            serv->CriticalSectionCOMGroupList.BeginReading();   // lock reading
                            if (SUCCEEDED(serv->m_COMGroupList.GetElem(group->m_hServerGroupHandle, &pCOMGroup))) {
                                ppCallback = pCOMGroup->m_vec.begin(); // Check if there is a registered callback function

                                if (*ppCallback) {            // Enabled Callback exist

                                    group->m_dwKeepAliveCount--;
                                    if (group->m_dwKeepAliveCount == 0) {
                                        group->ResetKeepAliveCounter();
                                        HRESULT hrErr = S_OK;
                                        pCOMGroup->FireOnDataChange(0, nullptr, &hrErr);
                                    }
                                }
                            }
                            serv->CriticalSectionCOMGroupList.EndReading(); // unlock reading
                        }
                        group->m_csKeepAlive.Unlock();
                    }
                    //
                    // Keep Alive
                    //
                }

                // unnail the group
                serv->ReleaseGenericGroup(idx);
            }

            // get the index of the next group in the list
            EnterCriticalSection(&serv->m_GroupsCritSec);
            res = serv->m_GroupList.Next(idx, &idx);
        }
        LeaveCriticalSection(&serv->m_GroupsCritSec);

    }

    _endthreadex(0);
    return 0;

} // WaitForUpdateThread


//=================================================================================
// Create Update Thread
// --------------------
//=================================================================================
HRESULT DaGenericServer::CreateWaitForUpdateThread(void)
{
    unsigned uThreadID;                       // Thread identifier

    m_UpdateThreadToKill = FALSE;

    m_hUpdateThread = (HANDLE)_beginthreadex(
        nullptr,                // No thread security attributes
        0,                   // Default stack size  
        WaitForUpdateThread, // Pointer to thread function 
        this,                // Pass class to new thread for access to the AppHandler
        0,                   // Run thread immediately
        &uThreadID);        // Thread identifier


    if (m_hUpdateThread == nullptr) {                  // Cannot create the thread
        return HRESULT_FROM_WIN32(GetLastError());
    }
    return S_OK;
}



//=================================================================================
// Kill Update Thread
// ------------------
//=================================================================================
HRESULT DaGenericServer::KillWaitForUpdateThread(void)
{
    if (m_hUpdateThread) {

        m_UpdateThreadToKill = TRUE;

        // give a chance to exit if the thread is waiting
        SetEvent(m_hUpdateEvent);

        // Wait max 60 secs until the update thread has terminated.
        if (WaitForSingleObject(m_hUpdateThread, 60000) == WAIT_TIMEOUT) {
            TerminateThread(m_hUpdateThread, 1);
        }
        CloseHandle(m_hUpdateThread);
        m_hUpdateThread = nullptr;
    }
    return S_OK;
}



//=================================================================================
// Check Base Update Rate
// ----------------------
//=================================================================================
BOOL DaGenericServer::CheckBaseUpdateRateChanged(void)
{
    DWORD newrate;

    newrate = m_pServerHandler->GetBaseUpdateRate();

    EnterCriticalSection(&m_UpdateRateCritSec);
    if (newrate != m_ActualBaseUpdateRate) {
        // base update rate has changed!

        m_ActualBaseUpdateRate = newrate;
        LeaveCriticalSection(&m_UpdateRateCritSec);

        return TRUE;
    }
    LeaveCriticalSection(&m_UpdateRateCritSec);

    return FALSE;
}


//=================================================================================
// Recalculate Tick Count information for group
// --------------------------------------------
//=================================================================================
HRESULT DaGenericServer::RecalcTicks(DaGenericGroup *group)
{
    HRESULT res;
    DWORD rate;

    _ASSERTE(group != NULL);

    rate = GetActualBaseUpdateRate();

    if (rate != group->GetActualBaseUpdateRate()) {
        // set the new update rate 
        group->SetActualBaseUpdateRate(rate);
        // recalc tick information
        res = group->ReviseUpdateRate(group->m_RequestedUpdateRate);
    }
    else {
        res = S_OK;
    }

    return res;
}


//=================================================================================
// returns the actual base update rate
//=================================================================================
DWORD DaGenericServer::GetActualBaseUpdateRate(void)
{
    DWORD rate;

    EnterCriticalSection(&m_UpdateRateCritSec);
    rate = m_ActualBaseUpdateRate;
    LeaveCriticalSection(&m_UpdateRateCritSec);

    return rate;
}


//=================================================================================
// sets the actual base update rate
//=================================================================================
HRESULT DaGenericServer::SetActualBaseUpdateRate(DWORD BaseUpdateRate)
{
    EnterCriticalSection(&m_UpdateRateCritSec);
    m_ActualBaseUpdateRate = BaseUpdateRate;
    LeaveCriticalSection(&m_UpdateRateCritSec);

    return S_OK;
}


//=================================================================================
// Fires a 'Shutdown Request' to the subscribed client
//=================================================================================
void DaGenericServer::FireShutdownRequest(LPCWSTR szReason)
{
    _ASSERTE(m_Created);
    _ASSERTE(m_pCOpcSrv);

    m_pCOpcSrv->FireShutdownRequest(szReason);
}



//=================================================================================
// Calls the OnRefreshOutputDevices() function in the server specific part.
// The current client name is provided as 'Actor ID'.
// This function returns only S_OK or S_FALSE.
//=================================================================================
HRESULT DaGenericServer::OnRefreshOutputDevices(
    /* [in] */                    DWORD             dwNumOfItems,
    /* [in][dwNumOfItems] */      DaDeviceItem    ** ppDItems,
    /* [in][out][dwNumOfItems] */ OPCITEMVQT     *  pVQTs,
    /* [in][out][dwNumOfItems] */ HRESULT        *  errors)
{
    _ASSERTE(m_Created);
    _ASSERTE(m_pCOpcSrv);

    LPWSTR   szActorID;
    HRESULT  hrActorID = m_pCOpcSrv->GetClientName(&szActorID);

    if (FAILED(hrActorID)) {
        szActorID = L"Unknown";                   // Must not be NULL
    }

    HRESULT hrRefresh = m_pServerHandler->OnRefreshOutputDevices(
        OPC_REFRESH_CLIENT, szActorID,
        dwNumOfItems,
        ppDItems,
        pVQTs,
        errors);

    _ASSERTE(SUCCEEDED(hrRefresh));          // Must return S_OK or S_FALSE

    if (SUCCEEDED(hrActorID)) {                // Release ActorID if successfully returned
        ComFree(szActorID);
    }

    return hrRefresh;
}



//=========================================================================
// InternalReadMaxAge
// ------------------
//    Read the given array of items. The MaxAge value will determine
//    where the data is obtained.
//
// This method is called from multiple threads 
//       - serving different clients
//       - handling sync/async read, refresh, update for the same client
// Non-reentrant parts must therefore be protected !
//
// Note :
//    -  ppDItems may contain NULL values ...
//       these are the items eliminated before read ...
//    -  pItemValues[x].vDataValue.vt contains the requested data type
//       or VT_EMPTY if the canonical data type should be used.
//    -  Since this function must return S_OK or S_FALSE a 
//       temporary buffer (ppTemporaryBuffer) and the current time
//       (pftNow) must be allocated/specified outside of this function
//       by the caller.
//    -  The size of the ppTemporaryBuffer must be identically with the
//       size of the ppDItems buffer.
//                                              
// return:
//    S_OK if succeeded for all items; otherwise S_FALSE.
//    Do not return other codes !
//=========================================================================
HRESULT DaGenericServer::InternalReadMaxAge(
    /* [in] */                    LPFILETIME        pftNow,
    /* [in] */                    DWORD             dwNumOfItems,
    /* [in][dwNumOfItems] */      DaDeviceItem    ** ppDItems,            // Contains the Device Item pointers, Arrays size is dwNumOfItems
    /* [in][dwNumOfItems] */      DaDeviceItem    ** ppTemporaryBuffer,   // A temporary buffer, must be provided by the caller
    /* [in][dwNumOfItems] */      DWORD          *  pdwMaxAges,
    /* [in][out][dwNumOfItems] */ VARIANT        *  pvValues,
    /* [in][out][dwNumOfItems] */ WORD           *  pwQualities,
    /* [in][out][dwNumOfItems] */ FILETIME       *  pftTimeStamps,
    /* [in][out][dwNumOfItems] */ HRESULT        *  errors,
    /* [in][dwNumOfItems] */      BOOL           *  pfPhyval /* = NULL */)
{
    // Contains the pointers to the Device Items for
    // which the value needs to be updated from the Device.
    DaDeviceItem**  pDItemsToReadFromDevice = ppTemporaryBuffer;
    HRESULT        hrRet = S_OK;

    DWORD i;
    DWORD dwNumOfItemsToReadFromDevice = 0;
    for (i = 0; i < dwNumOfItems; i++) {         // Check each requested Item

        pDItemsToReadFromDevice[i] = nullptr;
        DaDeviceItem* pDItem = ppDItems[i];
        if (pDItem) {

            // Max Age Handling

            if (pdwMaxAges[i] == 0) {              // Read always from device
                pDItemsToReadFromDevice[i] = pDItem;
                dwNumOfItemsToReadFromDevice++;
            }
            else if (pdwMaxAges[i] != 0xFFFFFFFF) {
                // Compare Max Age if not disabled

// Subtract Max Age from the current time
                ULONGLONG* pulNow = (ULONGLONG*)pftNow;
                ULONGLONG* pulAge = (ULONGLONG*)&pdwMaxAges[i];
                ULONGLONG  ulCompareTime = *pulNow - *pulAge;

                if (pDItem->IsTimeStampOlderThan((FILETIME*)&ulCompareTime)) {
                    pDItemsToReadFromDevice[i] = pDItem;
                    dwNumOfItemsToReadFromDevice++;
                }
            }
        }
    }


    //
    // Read from the Device for those Items which it's required.
    //
    if (dwNumOfItemsToReadFromDevice) {
        HRESULT hrRefresh = m_pServerHandler->OnRefreshInputCache(
            OPC_REFRESH_CLIENT,
            dwNumOfItems,
            pDItemsToReadFromDevice,
            errors);

        _ASSERTE(SUCCEEDED(hrRefresh));       // Must return S_OK or S_FALSE
    }


    //
    // Now the Data Cache is up to date and the values can be read
    //
    m_pServerHandler->readWriteLock_.BeginReading();

    for (i = 0; i < dwNumOfItems; i++) {         // Check each requested Item

        if (ppDItems[i]) {                        // Valid Item

            if (FAILED(errors[i])) {            // Read from Device failed
                hrRet = S_FALSE;
            }
            else {
                // Use the canonical data type if requested data
                // type is VT_EMPTY.
                if (V_VT(&pvValues[i]) == VT_EMPTY) {
                    V_VT(&pvValues[i]) = ppDItems[i]->get_CanonicalDataType();
                }
                // Read from Cache
                errors[i] = ppDItems[i]->get_ItemValue(
                    &pvValues[i],        // Value
                    &pwQualities[i],     // Quality
                    &pftTimeStamps[i]); // Timestamp

                if (FAILED(errors[i])) {
                    hrRet = S_FALSE;
                }
            }
        }
    }

    m_pServerHandler->readWriteLock_.EndReading();

    _ASSERTE(SUCCEEDED(hrRet));              // Must return S_OK or S_FALSE
    return hrRet;
}



//=========================================================================
// InternalWriteVQT
// ----------------
//    See comments of DaGenericGroup::InternalWrite() in
//    module DaGenericGroup.cpp.
//=========================================================================
HRESULT DaGenericServer::InternalWriteVQT(
    /* [in] */                    DWORD             dwNumOfItems,
    /* [in][dwNumOfItems] */      DaDeviceItem    ** ppDItems,
    /* [in][out][dwNumOfItems] */ OPCITEMVQT     *  pVQTs,
    /* [in][out][dwNumOfItems] */ HRESULT        *  errors,
    /* [in][dwNumOfItems] */      BOOL           *  pfPhyval /* = NULL */,
    /* [in] */                    long              hServerGroupHandle /* = -1 */)
{
    DaDeviceItem*   pDItem = nullptr;
    HRESULT        hr = S_OK;
    HRESULT        hrRet = S_OK;

    DWORD dwNumOfItemsToWrite = 0;               // Counts the number of valid item pointers in ppItems[]
    for (DWORD i = 0; i < dwNumOfItems; ++i) {   // Write value into cache

        pDItem = ppDItems[i];
        if (pDItem == nullptr) {                     // No refresh requested for this item
            continue;
        }

        hr = S_OK;

        OPCITEMVQT* pVQT = &pVQTs[i];

        if (V_VT(&pVQT->vDataValue) != VT_EMPTY) {
            // The value to be written must have the format required by the HW (indicated by
            // the canonical data type)
            hr = pDItem->ChangeValueToCanonicalType(&pVQT->vDataValue);
        }
        else if (!pVQT->bQualitySpecified && !pVQT->bTimeStampSpecified) {
            hr = OPC_E_BADTYPE;                    // At least one value should be specified
        }

        if (FAILED(hr)) {
            errors[i] = hr;
            ppDItems[i] = nullptr;
            pDItem->Detach();                      // Do not refresh output for this item
            hrRet = S_FALSE;
        }
        else {
            dwNumOfItemsToWrite++;
        }
    }

    if (dwNumOfItemsToWrite) {

        HRESULT hrRefresh = OnRefreshOutputDevices(
            dwNumOfItems,
            ppDItems,
            pVQTs,
            errors
        );

        if (hrRefresh == S_FALSE && hrRet == S_OK) {
            hrRet = S_FALSE;
        }
    }
    else {
        hrRet = S_FALSE;
    }

    _ASSERTE(SUCCEEDED(hrRet));              // Must return S_OK or S_FALSE
    return hrRet;
}
//DOM-IGNORE-END
