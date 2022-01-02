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

#ifndef __GENERICSERVER_H_
#define __GENERICSERVER_H_

 //DOM-IGNORE-BEGIN

#include "DaBaseServer.h"
#include "DaComBaseServer.h"
#include "DaGroup.h"
#include "DaGenericGroup.h"
#include "DaBrowse.h"
#include "ReadWriteLock.h"

#define  OPC_GROUPNAME_ENUM   1     // an enumerator which iterates over group names (Default).
#define  OPC_GROUP_ENUM       2     // an enumerator which iterates over group objects

class DaGroup;
class DaGenericGroup;
class DaBaseServer;
class DaComBaseServer;

class DaGenericServer {
    // this is the thread associated to each instance of this
    // class and therefore virtually a member of it:
    // virtually because the operating system doesn't allow
    // class methods to be passed to CreateThread
    friend unsigned __stdcall WaitForUpdateThread(void* pArg);
    friend class DaBrowseData;

public:

    //--------------------------------------------------------------
    // Constructor
    // -----------
    // After construction of the instance call Create to properly
    // initialize the class
    //--------------------------------------------------------------
    DaGenericServer(void);

    //--------------------------------------------------------------
    // Creator 
    // -------
    // Must be called to properly initialize the constructed instance,
    // if it fails then delete the instance immediatly
    //--------------------------------------------------------------
    HRESULT Create(DaBaseServer *pServerClassHandler, DaComBaseServer* pOPCServerBase);

    //--------------------------------------------------------------
    // Destructor
    //--------------------------------------------------------------
    ~DaGenericServer(void);

    //--------------------------------------------------------------
    //  Attach to class by incrementing the RefCount
    //  Returns the current RefCount
    //  or -1 if the ToKill flag is set
    //--------------------------------------------------------------
    int  Attach(void);

    //--------------------------------------------------------------
    //  Detach from class by decrementing the RefCount.
    //  Kill the class instance if request pending.
    //  Returns the current RefCount or -1 if it was killed.
    //--------------------------------------------------------------
    int  Detach(void);

    //--------------------------------------------------------------
    //  Kill the class instance.
    //  If the RefCount is > 0 then only the kill request flag is set.
    //  Returns the current RefCount or -1 if it was killed.
    //  If WithDeatch == TRUE then RefCound is first decremented.
    //--------------------------------------------------------------
    int  Kill(BOOL WithDetach);

    //--------------------------------------------------------------
    // Tells whether the instance was killed
    //--------------------------------------------------------------
    BOOL Killed(void);


    //==================================================== Member Variables
       // set if instance successfully created
    BOOL m_Created;

    // used to by the server instance to "nail" the group
    // so that this instance cannot be deleted while threads are running 
    // and accessing it
    long m_RefCount;

    // when a group has to be deleted set this bit TRUE
    BOOL m_ToKill;

    // This pointer must be set by a COM Server Class inheriting from this class.
    // It contains all the data shared by all server instances
    // of that Server Class
    DaBaseServer *m_pServerHandler;

    // Link to assosiated DaComServer class
    DaComBaseServer      *m_pCOpcSrv;

    // CGeneric groups
    OpenArray<DaGenericGroup *> m_GroupList;

    // access to the group arrays is synchronized ...
    // use this CS only for m_GroupList
    CRITICAL_SECTION m_GroupsCritSec;

    // associated COM groups ...
    // the fact that a CGeneric group exists doesn't imply
    // that the COM Group exists too, because the existance of
    // a COM Object is managed by COM ...
    // important is that when a COM Group is deleted by COM
    // the destructor removes the COM group from this list
    OpenArray<CComObject<DaGroup>*> m_COMGroupList;

    // the time this server instance was started
    FILETIME m_StartTime;

    // the last time the server sent values to the client
    FILETIME m_LastUpdateTime;

    // Enumerator type settings for automation interface
    OPCENUMSCOPE  m_Enum_Scope; // Indicates the class of groups to be enumerated
                            // OPC_ENUM_PRIVATE_CONNECTIONS  enumerates the private groups the client is connected to (Default).
                            // OPC_ENUM_PUBLIC_CONNECTIONS   enumerates the public groups the client is connected to.
                            // OPC_ENUM_ALL CONNECTIONS      enumerates all of the groups the client is connected to.
                            // OPC_ENUM_PRIVATE              enumerates all of the private groups created by the client (whether connected or not)
                            // OPC_ENUM_PUBLIC               enumerates all of the public groups available in the server (whether connected or not)
                            // OPC_ENUM_ALL                  enumerates all private groups and all public groups (whether connected or not)
    short    m_Enum_Type;  // The interface requested. 
                            // OPC_GROUPNAME_ENUM - an enumerator which iterates over group names (Default).
                            // OPC_GROUP_ENUM       an enumerator which iterates over group objects


       // Browse Address Space Enumerations (automation interface):
       // TRUE:  enumerating AccessPaths of item m_BrowseItemIDForAccessPath
       // FALSE: enumerating ItemIDs from the actual Browse Position of the server
    BOOL        m_EnumeratingItemIDs;

    // Data stored for Access Path enumeration
    LPWSTR      m_BrowseItemIDForAccessPath;

    // Data stored for item ID enumeration
    long        m_BrowseFilterType;
    LPWSTR      m_FilterCriteria;
    short       m_DataTypeFilter;
    long        m_AccessRightsFilter;

    // Current Browse Info for the server instance to browse in the address space
    DaBrowseData m_BrowseData;

    // Critical section for synchronisation of member variable access
    CRITICAL_SECTION m_CritSec;

public:

    // critical section
    HRESULT GetGenericGroup(long ServerGroupHandle, DaGenericGroup **group);
    HRESULT ReleaseGenericGroup(long ServerGroupHandle);
    HRESULT RemoveGenericGroup(long ServerGroupHandle);
    HRESULT ChangePrivateGroupName(DaGenericGroup *group, LPWSTR Name);

    // get a interface from a COM group from a DaGenericGroup
    // If the COM group does not exist it will be created
    HRESULT GetCOMGroup(
        DaGenericGroup*    pGGroup,
        REFIID            riid,
        LPUNKNOWN*        ppUnk
    );

    // no critical section
    HRESULT SearchGroup(
        BOOL     Public,
        LPCWSTR  theName,
        DaGenericGroup  **theGroup,
        long     *sgh
    );

    HRESULT GetGroupUniqueName(long ServerGroupHandle, WCHAR **theName, BOOL bPublic);

    HRESULT OnRefreshOutputDevices(
        /* [in] */                    DWORD             dwNumOfItems,
        /* [in][dwNumOfItems] */      DaDeviceItem    ** ppDItems,
        /* [in][out][dwNumOfItems] */ OPCITEMVQT     *  pVQTs,
        /* [in][out][dwNumOfItems] */ HRESULT        *  errors);

    HRESULT InternalReadMaxAge(
        /* [in] */                    LPFILETIME        pftNow,
        /* [in] */                    DWORD             dwNumOfItems,
        /* [in][dwNumOfItems] */      DaDeviceItem    ** ppDItems,         // Contains the Device Item pointers, Arrays size is dwNumOfItems
        /* [in][out][dwNumOfItems] */ DaDeviceItem    ** ppTemporaryBuffer,// A temporary buffer, must be provided by the caller
        /* [in][dwNumOfItems] */      DWORD          *  pdwMaxAges,
        /* [in][out][dwNumOfItems] */ VARIANT        *  pvValues,
        /* [in][out][dwNumOfItems] */ WORD           *  pwQualities,
        /* [in][out][dwNumOfItems] */ FILETIME       *  pftTimeStamps,
        /* [in][out][dwNumOfItems] */ HRESULT        *  errors,
        /* [in][dwNumOfItems] */      BOOL           *  pfPhyval = NULL);

    HRESULT InternalWriteVQT(
        /* [in] */                    DWORD             dwNumOfItems,
        /* [in][dwNumOfItems] */      DaDeviceItem    ** ppDItems,
        /* [in][out][dwNumOfItems] */ OPCITEMVQT     *  pVQTs,
        /* [in][out][dwNumOfItems] */ HRESULT        *  errors,
        /* [in][dwNumOfItems] */      BOOL           *  pfPhyval = NULL,
        /* [in] */                    long              hServerGroupHandle = -1);

public:
    HRESULT GetGrpList(REFIID riid, OPCENUMSCOPE scope,
        LPUNKNOWN **GroupList, int *GroupCount);

    HRESULT FreeGrpList(REFIID riid, LPUNKNOWN *GroupList, int count);

    HRESULT RefreshPublicGroups();

    //--------------------------------------------------------------
    // returns the actual base update rate
    // from which   m_ticks  is calculated
    //--------------------------------------------------------------
    DWORD GetActualBaseUpdateRate(void);

    //--------------------------------------------------------------
    // sets the actual base update rate
    //--------------------------------------------------------------
    HRESULT SetActualBaseUpdateRate(DWORD BaseUpdateRate);

    //--------------------------------------------------------------
    // Fires a 'Shutdown Request' to the subscribed client
    //--------------------------------------------------------------
    void FireShutdownRequest(LPCWSTR szReason);

    // protects access to m_COMGroupList
    ReadWriteLock CriticalSectionCOMGroupList;


    //////////////////////////
    // Update notification  //
    //////////////////////////
public:
    // this member is used by the server class handler
    // associated with this class ( daBaseServer_ )
    // to inform to send group update to the clients
    HANDLE      m_hUpdateEvent;

private:
    // the update rate for which the ticks count limits
    // of the groups are calculated
    // used by groups to set their actual value
    DWORD       m_ActualBaseUpdateRate;

    // protects access to  m_ActualBaseUpdateRate
    CRITICAL_SECTION m_UpdateRateCritSec;

    // handle of the refresh thread
    // associated with this class which has to send
    // updates to clients through IAdviseSink (custom)
    // or some Callback interface (automation)
    HANDLE      m_hUpdateThread;

    // tells whether the thread has to stop
    // or go on updating
    BOOL        m_UpdateThreadToKill;

    // creates the thread waiting for the  m_hUpdateEvent
    // and sending update callbacks to the clients
    HRESULT CreateWaitForUpdateThread(void);

    // destroys the thread waiting for the  m_hUpdateEvent
    HRESULT KillWaitForUpdateThread(void);

    // checks if update rate of server class handler has changed
    // and sets it to the new rate
    BOOL CheckBaseUpdateRateChanged(void);

    // recalculates tick count information for a group
    HRESULT RecalcTicks(DaGenericGroup *group);

};
//DOM-IGNORE-END

#endif // __GENERICSERVER_H_

