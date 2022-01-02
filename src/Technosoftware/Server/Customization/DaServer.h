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

#ifndef __DASERVER_H_
#define __DASERVER_H_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "DaBaseServer.h"
#include "DaAddressSpace.h"


//-----------------------------------------------------------------------------
// MACRO
//-----------------------------------------------------------------------------
#define CHECK_RESULT(f) {HRESULT hres = f; if (FAILED( hres )) throw hres;}
#define CHECK_PTR(p) {if ((p)== NULL) throw E_OUTOFMEMORY;}
#define CHECK_ADD(f) {if (!(f)) throw E_OUTOFMEMORY;}
#define CHECK_BOOL(f) {if (!(f)) throw E_FAIL;}


//-----------------------------------------------------------------------------
// Structure with Server Registry Definitions of the DLL.
// Do not change or delete !
//-----------------------------------------------------------------------------
typedef struct _tagSERVER_REGDEF {
	WCHAR * ClsidServer; // CLSID of current Server
      WCHAR  * ClsidApp;                        // CLSID of current Server Application
	WCHAR * PrgidServer; // Version independent Prog.Id.
      WCHAR  * PrgidCurrServer;                 // Prog.Id. of current Server
      WCHAR  * NameServer;                      // Friendly name of server
      WCHAR  * NameCurrServer;                  // Friendly name of current server version
      WCHAR  * CompanyName;       				// Server vendor name.
} ServerRegDefs;

#include "IClassicBaseNodeManager.h"


//-----------------------------------------------------------------------------
// CLASS DaDeviceCache
//-----------------------------------------------------------------------------
// Utility class used by OnRefreshOutputDevices()
class DaDeviceCache
{
	// Constrution
public:
   DaDeviceCache();
	HRESULT Create(DWORD dwSize);

	// Destruction
   ~DaDeviceCache();

	// Attributes
public:
   void**      m_paHandles;
   OPCITEMVQT* m_paItemVQTs;
   HRESULT*    m_paErrors;
   WORD*       m_paQualities;
   FILETIME*   m_paTimeStamps;

	// Implementation
protected:
   void Cleanup();
};


//-----------------------------------------------------------------------------
// CLASS DeviceItem
//-----------------------------------------------------------------------------
// My Server Specific Item Definition
class DeviceItem : public DaDeviceItem
{
	// Construction
public:
   DeviceItem();
   inline HRESULT Create(                       // Initializer
		LPWSTR szItemID,
                     DWORD          dwAccessRights,
                     LPVARIANT      pvValue,
                     BOOL           fActive        = TRUE,
                     DWORD          dwBlobSize     = 0,
                     LPBYTE         pBlob          = NULL,
                     LPWSTR         szAccessPath   = NULL,
                     OPCEUTYPE      eEUType        = OPC_NOENUM,
                     LPVARIANT      pvEUInfo       = NULL
                 ) {

		return DaDeviceItem::Create(szItemID,
                                    dwAccessRights,
                                    pvValue,
                                    fActive,
                                    dwBlobSize,
                                    pBlob,
                                    szAccessPath,
                                    eEUType,
			pvEUInfo);
   }

	// Destruction
   ~DeviceItem();

   HRESULT AttachActiveCount(void);
   HRESULT DetachActiveCount(void);

};


//-----------------------------------------------------------------------------
// CLASS DaServer
//-----------------------------------------------------------------------------
// My Server Specific Data Access Implementation
class DaServer : public DaBaseServer
{
	friend unsigned __stdcall NotifyUpdateThread(LPVOID pAttr);

	// Construction
public:
   DaServer();
	HRESULT Create(DWORD dwUpdateRate, LPCTSTR szDLLParams, LPCTSTR szInstanceNamebool);

	// Destruction
   ~DaServer();

	// Attributes
public:
   inline OPCSERVERSTATE ServerState() const { return m_dwServerState; }
   inline void SetServerState(OPCSERVERSTATE serverState) { m_dwServerState = serverState; m_fServerStateChanged = TRUE; }
   inline DWORD          BandWidth() const { return m_dwBandWith; }
   inline DaBranch* SASRoot() { return &m_SASRoot; }



	// Operations
public:
	HRESULT AddDeviceItem(DeviceItem* pDItem);
	HRESULT RemoveDeviceItem(DeviceItem* pDItem);
	void DeleteDeviceItem(DeviceItem* pDItem);

         // array of items supported by this server
   CSimplePtrArray<DeviceItem*>  m_arServerItems;

      // Use the members of this object to lock and unlock
      // critical sections to allow consistent access of item
      // list (m_ServerItems), avoiding interferences with
      // some refresh threads.
   ReadWriteLock                    m_ItemListLock;
   ReadWriteLock                    m_OnRequestItemsLock;

	// Implementation
protected:

   /////////////////////////////////////////////////////////////////////////////
   // Implementation of the pure virtual functions of the DaBaseServer //
   /////////////////////////////////////////////////////////////////////////////

      ////////////////////////////////
      // Common DA Server Functions //
      ////////////////////////////////

   HRESULT OnGetServerState(
         /*[out]        */          DWORD          & bandWidth,
		 /*[out]		*/			OPCSERVERSTATE & serverState,
		/*[out] */ LPWSTR & vendor);

      // use to add/release custom specific client related data
	HRESULT OnCreateCustomServerData(LPVOID* customData);

   //*************************************************************************
   //* INFO: The following function is declared as virtual and
   //*       implementation is therefore optional.
   //*       The following code is only here for informational purpose
   //*       and if no server specific handling is needed the code
   //*       can savely be deleted.
   //*************************************************************************
	HRESULT OnDestroyCustomServerData(LPVOID* customData);

		///////////////////////
		// Validate Function //
		///////////////////////

   HRESULT OnValidateItems(
		/*[in]                        */     OPC_VALIDATE_REQUEST
                                                                 validateRequest,
        /*[in]                        */     BOOL                blobUpdate,
		/*[in]                        */     DWORD               numItems,
		/*[in,  size_is(numItems)]  */     OPCITEMDEF       *  itemDefinitions,
		/*[out, size_is(numItems)]  */     DaDeviceItem      ** daDeviceItems,
        /*[out, size_is(numItems)]  */     OPCITEMRESULT    *  itemResults,
		/*[out, size_is(numItems)]  */     HRESULT          *  errors
      );


		///////////////////////
		// Refresh Functions //
		///////////////////////

   HRESULT OnRefreshInputCache(
			/*[in]                        */    OPC_REFRESH_REASON   dwReason,
         /*[in]                        */    DWORD                numItems,
         /*[in, size_is(,numItems)]  */    DaDeviceItem       ** pDevItemPtr,
			/*[in, size_is(,numItems)]  */    HRESULT           *  errors
      );

   HRESULT OnRefreshOutputDevices(
			/*[in]                        */    OPC_REFRESH_REASON   dwReason,
         /*[in, string]                */    LPWSTR               szActorID,
         /*[in]                        */    DWORD                numItems,
         /*[in, size_is(,numItems)]  */    DaDeviceItem       ** pDevItemPtr,
			/*[in, size_is(,numItems)]  */    OPCITEMVQT        *  pItemVQTs,
			/*[in, size_is(,numItems)]  */    HRESULT           *  errors
      );

      ///////////////////////////////////////////
      // Server Address Space Browse Functions //
      ///////////////////////////////////////////

   OPCNAMESPACETYPE OnBrowseOrganization();

   HRESULT OnBrowseChangeAddressSpacePosition(
         /*[in]         */          OPCBROWSEDIRECTION
		dwBrowseDirection,
         /*[in, string] */          LPCWSTR        szString,
         /*[in,out]     */          LPWSTR      *  szActualPosition,
		/*[in,out] */ LPVOID * customData);


   HRESULT OnBrowseGetFullItemIdentifier(
         /*[in]         */          LPWSTR         szActualPosition,
         /*[in]         */          LPWSTR         szItemDataID,
         /*[out, string]*/          LPWSTR      *  szItemID,
		/*[in,out] */ LPVOID * customData);

   HRESULT OnBrowseItemIdentifiers(
         /*[in]         */          LPWSTR         szActualPosition,
         /*[in]         */          OPCBROWSETYPE  dwBrowseFilterType,
		/*[in, string] */ LPWSTR szFilterCriteria,
		/*[in] */ VARTYPE vtDataTypeFilter,
         /*[in]         */          DWORD          dwAccessRightsFilter,
         /*[out]        */          DWORD       *  pNrItemIds,
         /*[out, string]*/          LPWSTR      ** ppItemIds,
		/*[in,out] */ LPVOID * customData);

   //*************************************************************************
   //* INFO: The following function is declared as virtual and
   //*       implementation is therefore optional.
   //*       The following code is only here for informational purpose
   //*       and if no server specific handling is needed the code
   //*       can savely be deleted.
   //*************************************************************************
   // HRESULT OnBrowseAccessPaths(
   //       /*[in]         */          LPWSTR         szItemID,
	// /*[out] */ DWORD * pNrAccessPaths,
   //       /*[out, string]*/          LPWSTR      ** szAccessPaths,
   //       /*[in,out]     */          LPVOID      *  customData );

      /////////////////////////////
      // Item Property Functions //
      /////////////////////////////

   HRESULT OnQueryItemProperties(
         /*[in]         */          LPCWSTR        szItemID,
         /*[out]        */          LPDWORD        pdwCount,
         /*[out]        */          LPDWORD     *  ppdwPropIDs,
		/*[out] */ LPVOID * ppCookie);

   HRESULT OnGetItemProperty(
         /*[in]         */          LPCWSTR        szItemID,
         /*[in]         */          DWORD          dwPropID,
         /*[out]        */          LPVARIANT      pvPropData,
		/*[in] */ LPVOID pCookie);

   //*************************************************************************
   //* INFO: The following function is declared as virtual and
   //*       implementation is therefore optional.
   //*       The following code is only here for informational purpose
   //*       and if no server specific handling is needed the code
   //*       can savely be deleted.
   //*************************************************************************
   // HRESULT OnLookupItemId(
   //       /*[in]         */          LPCWSTR        szItemID,
   //       /*[in]         */          DWORD          dwPropID,
   //       /*[out]        */          LPWSTR      *  pszNewItemId,
   //       /*[in]         */          LPVOID         pCookie );

   HRESULT OnReleasePropertyCookie(
		/*[in] */ LPVOID pCookie);

   //////////////////////////////////////////////////////////////
   // Implementation internal functions (application specific) //
   //////////////////////////////////////////////////////////////

   HRESULT DeleteServerItems();
   HRESULT CreateUpdateThread();
   HRESULT KillUpdateThread();
	HRESULT FindDeviceItem(LPCWSTR szItemID, DaDeviceItem** ppDItem);


      // Handle of the Refresh Thread
   HANDLE               m_hUpdateThread;
      // Signal-Handle to terminate the threads.
   HANDLE               m_hTerminateThraedsEvent;

   OPCSERVERSTATE       m_dwServerState;
   BOOL			m_fServerStateChanged;
   DWORD                m_dwBandWith;

      // Root of the hierarchical Server Address Space
   DaBranch           m_SASRoot;

   BOOL                 m_fCreated;
};

// The Global Data Server Handler
extern DaServer* gpDataServer;

#endif // __DASERVER_H_
