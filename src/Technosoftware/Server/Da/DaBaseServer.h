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

#ifndef __SERVERCLASSHANDLER_H_
#define __SERVERCLASSHANDLER_H_


#include "DaPublicGroupManager.h"
#include "DaServerInstanceHandle.h"
#include "ReadWriteLock.h"
#include "DaItemProperty.h"
#include "DaDeviceItem.h" 
#include "IClassicBaseNodeManager.h" 

/**
 * @typedef enum tagOPC_REFRESH_REASON
 *
 * @brief   Reasons for calling the cache refresh function RefreshOutputDevices() /
 *          RefreshInputCache()
 */

typedef enum tagOPC_REFRESH_REASON {
    OPC_REFRESH_INIT,
    OPC_REFRESH_CLIENT,
    OPC_REFRESH_PERIODIC
} OPC_REFRESH_REASON;

/**
 * @typedef enum tagOPC_VALIDATE_REQUEST
 *
 * @brief   Requested Validate Results.
 */

typedef enum tagOPC_VALIDATE_REQUEST {
    OPC_VALIDATEREQ_ITEMRESULTS,
    OPC_VALIDATEREQ_DEVICEITEMS
} OPC_VALIDATE_REQUEST;


/////////////////////////////////////////////////////////////////////////////////
//
// All the server instances of the same COM-OPC server class share 
// a pointer to an instance of this class (a derivate).
// No global var and procedures are used because there could  
// be more server objects in the EXE having if not different behaviour
// at least a different initialization and data sources.
// For each COM-OPC server class a different instance of a DaBaseServer
// derivate should be created ( else they would have exactly the same behaviour ) 
//
/////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////
// MACROs to define Property Maps
/////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////
// MACRO BEGIN_ITEMPROPERTY_MAP()
/////////////////////////////////////////////////////////////////////////////////
#define BEGIN_ITEMPROPERTY_MAP() \
   class CPropDefs { \
      public:\
               HRESULT Create( DaBaseServer* pServerHandler ) { \
                  static OPCITEMPROPERTY_OPC aPropDef[] = {


/////////////////////////////////////////////////////////////////////////////////
// MACRO ITEMPROPERTY_ENTRY()
/////////////////////////////////////////////////////////////////////////////////
#define ITEMPROPERTY_ENTRY( ID, VT, DESCR )  { ID, VT, DESCR },


/////////////////////////////////////////////////////////////////////////////////
// MACRO END_ITEMPROPERTY_MAP()
/////////////////////////////////////////////////////////////////////////////////
#define END_ITEMPROPERTY_MAP()               { 0,  VT_EMPTY, NULL} };\
                  DaItemProperty* pProp = (DaItemProperty *)aPropDef;\
                  HRESULT hres;\
                  while (pProp->pwszDescr) {\
                     hres = pServerHandler->AddItemProperty( pProp++ );\
                     if (FAILED( hres ))  return hres;\
                  }\
                  return S_OK;\
               }\
   };\
   public:\
      CPropDefs   _PropMap;


/////////////////////////////////////////////////////////////////////////////////
// MACRO SETUP_ITEMPROPERTIES_FROM_MAP()
/////////////////////////////////////////////////////////////////////////////////
#define SETUP_ITEMPROPERTIES_FROM_MAP() _PropMap.Create( this )

class DaDeviceItem;
class DaGenericServer;

/**
 * @class	DaBaseServer
 *
 * @brief	A Data Access Server must have one instance of a DaBaseServer derived class. This
 * 			object is the main interface between the generic server part and the application
 * 			specific part for DA functionality.
 *
 * 			Those functions which are called from the generic server part are prefixed with On
 * 			and must be implemented in the application specific part.
 *
 * 			A good place to create a Data Server Handler is the global interface function
 * 			OnInitializeServer(). The Handler object can be deleted in function
 * 			OnTerminateServer().
 */

class DaBaseServer {

    /**
     * @brief	this friend declaration allows the class DaGenericServer to add itself to the list of
     * 			this handler using AddServerToList  and  RemoveServerFromList which is private and
     * 			should not be accessed by the specific part of the server.
     */

    friend DaGenericServer;

    /**
     * @fn	DaBaseServer::BEGIN_ITEMPROPERTY_MAP() ITEMPROPERTY_ENTRY(OPC_PROPERTY_DATATYPE, VT_I2, OPC_PROPERTY_DESC_DATATYPE) ITEMPROPERTY_ENTRY(OPC_PROPERTY_VALUE, VT_EMPTY, OPC_PROPERTY_DESC_VALUE) ITEMPROPERTY_ENTRY(OPC_PROPERTY_QUALITY, VT_I2, OPC_PROPERTY_DESC_QUALITY) ITEMPROPERTY_ENTRY(OPC_PROPERTY_TIMESTAMP, VT_DATE, OPC_PROPERTY_DESC_TIMESTAMP) ITEMPROPERTY_ENTRY(OPC_PROPERTY_ACCESS_RIGHTS, VT_I4, OPC_PROPERTY_DESC_ACCESS_RIGHTS) ITEMPROPERTY_ENTRY(OPC_PROPERTY_SCAN_RATE, VT_R4, OPC_PROPERTY_DESC_SCAN_RATE) ITEMPROPERTY_ENTRY(OPC_PROPERTY_EU_TYPE, VT_I4, OPC_PROPERTY_DESC_EU_TYPE) ITEMPROPERTY_ENTRY(OPC_PROPERTY_EU_INFO, VT_ARRAY | VT_BSTR, OPC_PROPERTY_DESC_EU_INFO) #ifndef _NO_DEADBAND_HANDLING ITEMPROPERTY_ENTRY(OPC_PROPERTY_HIGH_EU, VT_R8, OPC_PROPERTY_DESC_HIGH_EU) ITEMPROPERTY_ENTRY(OPC_PROPERTY_LOW_EU, VT_R8, OPC_PROPERTY_DESC_LOW_EU) #endif END_ITEMPROPERTY_MAP() public;
     *
     * @brief	Property IDs supported by default.
     */

    BEGIN_ITEMPROPERTY_MAP()
        // OPC Specific Properties
        ITEMPROPERTY_ENTRY(OPC_PROPERTY_DATATYPE, VT_I2, OPC_PROPERTY_DESC_DATATYPE)
        ITEMPROPERTY_ENTRY(OPC_PROPERTY_VALUE, VT_EMPTY, OPC_PROPERTY_DESC_VALUE)
        ITEMPROPERTY_ENTRY(OPC_PROPERTY_QUALITY, VT_I2, OPC_PROPERTY_DESC_QUALITY)
        ITEMPROPERTY_ENTRY(OPC_PROPERTY_TIMESTAMP, VT_DATE, OPC_PROPERTY_DESC_TIMESTAMP)
        ITEMPROPERTY_ENTRY(OPC_PROPERTY_ACCESS_RIGHTS, VT_I4, OPC_PROPERTY_DESC_ACCESS_RIGHTS)
        ITEMPROPERTY_ENTRY(OPC_PROPERTY_SCAN_RATE, VT_R4, OPC_PROPERTY_DESC_SCAN_RATE)
        ITEMPROPERTY_ENTRY(OPC_PROPERTY_EU_TYPE, VT_I4, OPC_PROPERTY_DESC_EU_TYPE)
        ITEMPROPERTY_ENTRY(OPC_PROPERTY_EU_INFO, VT_ARRAY | VT_BSTR, OPC_PROPERTY_DESC_EU_INFO)

        // Recommended Properties
#ifndef _NO_DEADBAND_HANDLING
        ITEMPROPERTY_ENTRY(OPC_PROPERTY_HIGH_EU, VT_R8, OPC_PROPERTY_DESC_HIGH_EU)
        ITEMPROPERTY_ENTRY(OPC_PROPERTY_LOW_EU, VT_R8, OPC_PROPERTY_DESC_LOW_EU)
#endif

    END_ITEMPROPERTY_MAP()

public:

    /**
     * @fn	DaBaseServer::DaBaseServer(void);
     *
     * @brief	constructor.
     */

    DaBaseServer(void);

    /**
     * @fn	DaBaseServer::~DaBaseServer();
     *
     * @brief	destructor.
     */

protected:
    ~DaBaseServer();

public:
    /**
    * @fn	HRESULT DaBaseServer::Create(int instanceIndex, LPCWSTR instanceName, DWORD updateRate);
    *
    * @brief	initializes the class after construction.
    *
    * @param	instanceIndex	Zero-based index of the instance.
    * @param	instanceName 	Name of the instance.
    * @param	updateRate   	The update rate.
    *
    * @return	A hResult.
    */

    HRESULT Create(int instanceIndex, LPCWSTR instanceName, DWORD updateRate);

    /**
     * @brief   the common public groups created by default when the server loads. Instance is
     *          initialized in in Create().
     */

    DaPublicGroupManager  publicGroups_;

    /**
     * @fn  inline int DaBaseServer::InstanceIndex() const
     *
     * @brief   Name and Index of the Server Instance.
     *
     * @return  An int.
     */

    inline int InstanceIndex() const { return instanceIndex_; }

    HRESULT GetName(LPWSTR *pSrvName, IMalloc *pIMalloc = NULL);
    HRESULT SetName(LPCWSTR pSrvName);

    ////////////////////////////////////////////////////////////
    // Start of functions located in the server-specific part //
    ////////////////////////////////////////////////////////////
public:

    ////////////////////////////////
    // Common DA Server Functions //
    ////////////////////////////////

    virtual HRESULT OnGetServerState(
        /*[out]                             */ DWORD                    & bandWidth,
        /*[out]		                        */ OPCSERVERSTATE           & serverState,
        /*[out]		                        */ LPWSTR		            & vendor) = 0;

    /**
     * @fn  virtual HRESULT DaBaseServer::OnCreateCustomServerData( LPVOID * customData)
     *
     * @brief   used to add custom specific client related data.
     *
     * @param [in,out]  customData  If non-null, information describing the custom.
     *
     * @return  A hResult.
     */

    virtual HRESULT OnCreateCustomServerData(
        /*[out                              */ LPVOID                   * customData)
    {
        return S_OK;
    }

    /**
     * @fn  virtual HRESULT DaBaseServer::OnDestroyCustomServerData( LPVOID * customData)
     *
     * @brief   used to release custom specific client related data.
     *
     * @param [in,out]  customData  If non-null, information describing the custom.
     *
     * @return  A hResult.
     */

    virtual HRESULT OnDestroyCustomServerData(
        /*[in, out                          */ LPVOID                   * customData)
    {
        return S_OK;
    }

    ///////////////////////
    // Validate Function //
    ///////////////////////

    /**
     * @fn  virtual HRESULT DaBaseServer::OnValidateItems( OPC_VALIDATE_REQUEST validateRequest, BOOL blobUpdate, DWORD numItems, OPCITEMDEF * itemDefinitions, DaDeviceItem ** daDeviceItems, OPCITEMRESULT * itemResults, HRESULT * errors) = 0;
     *
     * @brief   Checks the ItemID, the Requested Data Type, etc. and returns the OPCITEMRESULT
     *          Structures or the Device Items (depening on the validateRequest Parameter).
     *
     * @param   validateRequest         The validate request.
     * @param   blobUpdate              true to BLOB update.
     * @param   numItems                Number of items.
     * @param [in,out]  itemDefinitions If non-null, the item defs.
     * @param [in,out]  daDeviceItems   If non-null, the d item ptrs.
     * @param [in,out]  itemResults     If non-null, the opc item results.
     * @param [in,out]  errors          If non-null, the errors.
     *
     * @return  A hResult.
     */

    virtual HRESULT OnValidateItems(
        /*[in]                              */ OPC_VALIDATE_REQUEST       validateRequest,
        /*[in]                              */ BOOL                       blobUpdate,
        /*[in]                              */ DWORD                      numItems,
        /*[in,  size_is(numItems)]          */ OPCITEMDEF               * itemDefinitions,
        /*[out, size_is(numItems)]          */ DaDeviceItem            ** daDeviceItems,
        /*[out, size_is(numItems)]          */ OPCITEMRESULT            * itemResults,
        /*[out, size_is(numItems)]          */ HRESULT                  * errors) = 0;

    //////////////////////////////////
    // Data Cache Refresh Functions //
    //////////////////////////////////

    /**
     * @brief   Use the members of this object to lock and unlock critical sections to allow the
     *          reading and writing of consistent data, avoiding interferences with some Refresh
     *          threads.
     */

    ReadWriteLock           readWriteLock_;


    virtual HRESULT OnRefreshInputCache(
        /*[in]                              */ OPC_REFRESH_REASON         dwReason,
        /*[in]                              */ DWORD                      numItems,
        /*[in, size_is(,numItems)]          */ DaDeviceItem            ** pDevItemPtr,
        /*[in, size_is(,numItems)]          */ HRESULT                  * errors) = 0;

    virtual HRESULT OnRefreshOutputDevices(
        /*[in]                              */ OPC_REFRESH_REASON         dwReason,
        /*[in, string]                      */ LPWSTR                     szActorID,
        /*[in]                              */ DWORD                      numItems,
        /*[in, size_is(,numItems)]          */ DaDeviceItem            ** pDevItemPtr,
        /*[in, size_is(,numItems)]          */ OPCITEMVQT               * pItemVQTs,
        /*[in, size_is(,numItems)]          */ HRESULT                  * errors) = 0;

    ///////////////////////////////////////////
    // Server Address Space Browse Functions //
    ///////////////////////////////////////////

    virtual OPCNAMESPACETYPE OnBrowseOrganization() = 0;

    virtual HRESULT OnBrowseChangeAddressSpacePosition(
        /*[in]         */                   OPCBROWSEDIRECTION
        dwBrowseDirection,
        /*[in, string] */                   LPCWSTR        szString,
        /*[in,out]     */                   LPWSTR      *  szActualPosition,
        /*[in,out]     */                   LPVOID      *  customData) = 0;

    virtual HRESULT OnBrowseGetFullItemIdentifier(
        /*[in]         */                   LPWSTR         szActualPosition,
        /*[in]         */                   LPWSTR         szItemDataID,
        /*[out, string]*/                   LPWSTR      *  szItemID,
        /*[in,out]     */                   LPVOID      *  customData) = 0;

    virtual HRESULT OnBrowseItemIdentifiers(
        /*[in]         */                   LPWSTR         szActualPosition,
        /*[in]         */                   OPCBROWSETYPE  dwBrowseFilterType,
        /*[in, string] */                   LPWSTR         szFilterCriteria,
        /*[in]         */                   VARTYPE        vtDataTypeFilter,
        /*[in]         */                   DWORD          dwAccessRightsFilter,
        /*[out]        */                   DWORD       *  pNrItemIDs,
        /*[out, string]*/                   LPWSTR      ** ppItemIDs,
        /*[in,out]     */                   LPVOID      *  customData) = 0;

    virtual HRESULT OnBrowseAccessPaths(
        /*[in]         */                   LPWSTR         szItemID,
        /*[out]        */                   DWORD       *  pNrAccessPaths,
        /*[out, string]*/                   LPWSTR      ** szAccessPaths,
        /*[in,out]     */                   LPVOID      *  customData)
    {
        return E_NOTIMPL;
    }

    /////////////////////////////
    // Item Property Functions //
    /////////////////////////////

    virtual HRESULT OnQueryItemProperties(
        /*[in]         */                   LPCWSTR        szItemID,
        /*[out]        */                   LPDWORD        pdwCount,
        /*[out]        */                   LPDWORD     *  ppdwPropIDs,
        /*[out]        */                   LPVOID      *  ppCookie) = 0;

    virtual HRESULT OnGetItemProperty(
        /*[in]         */                   LPCWSTR        szItemID,
        /*[in]         */                   DWORD          dwPropID,
        /*[out]        */                   LPVARIANT      pvPropData,
        /*[in]         */                   LPVOID         pCookie) = 0;

    virtual HRESULT OnLookupItemID(
        /*[in]         */                   LPCWSTR        szItemID,
        /*[in]         */                   DWORD          dwPropID,
        /*[out]        */                   LPWSTR      *  pszNewItemID,
        /*[in]         */                   LPVOID         pCookie)
    {
        return E_FAIL;
    }

    virtual HRESULT OnReleasePropertyCookie(
        /*[in]         */                   LPVOID         pCookie)
    {
        return S_OK;
    }

    //////////////////////////////////////////////////////////
    // End of functions located in the server-specific part //
    //////////////////////////////////////////////////////////

       //////////////////////////////////
       // Update Notification          //
       //////////////////////////////////
private:

    /**
     * @brief	this is the base Update Rate of the server and must be set when creating an instance
     * 			of this class (derived class)
     * 			should be a non zero value (like all update rates are supposed to be). This is used
     * 			by the standard implementation of ReviseUpdateRate and by the
     * 			UpdateServerClassInstances method (it's in millisec)
     */

    DWORD baseUpdateRate_;

    /** @brief	critical section for accessing members of this class. */
    CRITICAL_SECTION criticalSection_;

public:

    /**
     * @fn	HRESULT DaBaseServer::SetBaseUpdateRate(DWORD baseUpdateRate);
     *
     * @brief	sets the base update rate of the server.
     *
     * @param	baseUpdateRate	The base update rate.
     *
     * @return	A hResult.
     */

    HRESULT SetBaseUpdateRate(DWORD baseUpdateRate);

    /**
     * @fn	DWORD DaBaseServer::GetBaseUpdateRate(void);
     *
     * @brief	gets the base update rate of the server.
     *
     * @return	The base update rate.
     */

    DWORD GetBaseUpdateRate(void);

    /**
     * @fn	virtual HRESULT DaBaseServer::ReviseUpdateRate( DWORD requestedUpdateRate, DWORD *revisedUpdateRate);
     *
     * @brief	revises the update rate according to the some application specific rate: in theory
     * 			(OPC Spec.)  RevisedUpdateRate >= RequestedUpdateRate should be true;
     * 			RevisedUpdateRate > 0  must always be true.
     *
     * @param	requestedUpdateRate		 	The requested update rate.
     * @param [in,out]	revisedUpdateRate	If non-null, the revised update rate.
     *
     * @return	A hResult.
     */

    virtual HRESULT ReviseUpdateRate(
        DWORD requestedUpdateRate,
        DWORD *revisedUpdateRate);

    /**
     * @fn	HRESULT DaBaseServer::UpdateServerClassInstances();
     *
     * @brief	this method should be called each  baseUpdateRate_  millisec to trigger the advise
     * 			mechanism to the clients;
     * 			it's not virtual because only methods of this class can access the list of server
     * 			attached to this handler (servers_), application specific derivations won't have
     * 			access to it and cannot therefore update the clients.
     *
     * @return	A hResult.
     */

    HRESULT UpdateServerClassInstances();


    /**
    * @fn	HRESULT DaBaseServer::QueryInterfaceOnServerFromList( SRVINSTHANDLE server, BOOL & serverExist, REFIID refId, LPVOID * unkown)
    *
    * @brief	Executes QueryInterface on a specified server from list
    * 			Test if in this class handler's list a server with the specified Server Instance
    * 			Handle exist. If the server exist QueryInterface with the defined parameters will be
    * 			executed.
    *
    * 			results :
    *
    * 			serverExist = FALSE / HRESULT = E_FAIL;     Server doesn't exist in list
    * 			serverExist = TRUE  / HRESULT = Result from QueryInterface()
    *
    * @param	server			   	The server.
    * @param [in,out]	serverExist	The server exist.
    * @param	refId			   	Identifier for the reference.
    * @param [in,out]	unkown	   	If non-null, the unkown.
    *
    * @return	The interface on server from list.
    */

    HRESULT QueryInterfaceOnServerFromList(SRVINSTHANDLE  server,
        BOOL     &     serverExist,
        REFIID         refId,
        LPVOID   *     unkown);


    // Returns the item property definition with the specified ID
    HRESULT GetItemProperty(DWORD dwPropID, DaItemProperty** ppProp);

    // Add a new item property definition
    inline
        HRESULT AddItemProperty(DaItemProperty* pProp)
    {
        return itemProperties_.PutElem(itemProperties_.New(), pProp);
    }

    // Fires a 'Shutdown Request' to all subscribed clients
    void FireShutdownRequest(LPCWSTR szReason = NULL);

    void GetClients(int * numClientHandles, void* ** clientHandles, LPWSTR ** clientNames);

    void GetGroups(void * clientHandle, int * numGroupHandles, void* ** groupHandles, LPWSTR ** groupNames);

    void GetGroupState(void * groupHandle, IClassicBaseNodeManager::DaGroupState * groupState);

	void GetItemStates(void * groupHandle, int * numDaItemStates, IClassicBaseNodeManager::DaItemState* * daItemStates);

private:
    /** @brief	Index of the Server Instance. */
    int   instanceIndex_;

    /** @brief	tells whether the class instance has been successfully created. */
    BOOL     created_;

    /** @brief	Name of the Server Instance. */
    LPWSTR   name_;

    /**
    * @brief	this is the list of the Server instances (one for each client connected)
    * 			attached to the actual server class handler instance.
    */

    OpenArray<DaGenericServer *> servers_;

    /** @brief   critical section for accessing the list of server instances. */
    CRITICAL_SECTION serversCriticalSection_;

    /**
     * @fn	HRESULT DaBaseServer::AddServerToList(DaGenericServer *server);
     *
     * @brief	adds a server instance to this class handler's list;
     * 			this method is called when a client connects.
     *
     * @param [in,out]	server	If non-null, the server.
     *
     * @return	A hResult.
     */

    HRESULT AddServerToList(DaGenericServer *server);

    /**
     * @fn	HRESULT DaBaseServer::RemoveServerFromList(DaGenericServer *server);
     *
     * @brief	removes a server instance from this class handler's list;
     * 			this method is called when a client has properly disconnected, that is freed all the
     * 			COM interfaces it owns.
     *
     * @param [in,out]	server	If non-null, the server.
     *
     * @return	A hResult.
     */

    HRESULT RemoveServerFromList(DaGenericServer *server);

    /** @brief	this list conatains all Item Properties which may be attached to an item. */
    OpenArray<DaItemProperty *> itemProperties_;
};

#endif // __SERVERCLASSHANDLER_
