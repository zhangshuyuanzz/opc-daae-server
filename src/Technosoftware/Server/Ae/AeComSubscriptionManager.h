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

#ifndef __EVENTSUBSCRIPTIONMGT_H
#define __EVENTSUBSCRIPTIONMGT_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "UtilityDefs.h"                        // for CSimplePtrArray<>

class AeBaseServer;
class AeComBaseServer;
class AeSource;
class AeEvent;
class WideString;


//-----------------------------------------------------------------------
// DEFINES
//-----------------------------------------------------------------------
//    Minimum value of low severity
//    The spec V1.02 is uncertain if the low severity is 0 or 1
//    WARNING: Spec says different things
#define  MIN_LOW_SEVERITY  1


//-----------------------------------------------------------------------
// CLASS
//-----------------------------------------------------------------------
class AeComSubscriptionManager : public IOPCEventSubscriptionMgt
{
// Construction
public:
   AeComSubscriptionManager();
   HRESULT  Create(
                  /* [in] */                    AeComBaseServer*  pServer,
                  /* [in] */                    BOOL           bActive,
                  /* [in] */                    DWORD          dwBufferTime,
                  /* [in] */                    DWORD          dwMaxSize,
                  /* [in] */                    OPCHANDLE      hClientSubscription,
                  /* [out] */                   DWORD       *  pdwRevisedBufferTime,
                  /* [out] */                   DWORD       *  pdwRevisedMaxSize );

// Destruction
public:
   ~AeComSubscriptionManager();
   void CleanupSubscription();

// Attributes
public:
   inline BOOL IsCancelRefreshActive() { return m_fCancelRefresh; }
   static DWORD DEFAULT_SIZE_EVENT_BUFFER;

// Operations
public:
   ///////////////////////////////////////////////////////////////////////////
   ////////////////////////// IOPCEventSubscriptionMgt ///////////////////////
   ///////////////////////////////////////////////////////////////////////////

   STDMETHODIMP SetFilter(
                  /* [in] */                    DWORD          dwEventType,
                  /* [in] */                    DWORD          dwNumCategories,
                  /* [size_is][in] */           DWORD       *  pdwEventCategories,
                  /* [in] */                    DWORD          dwLowSeverity,
                  /* [in] */                    DWORD          dwHighSeverity,
                  /* [in] */                    DWORD          dwNumAreas,
                  /* [size_is][in] */           LPWSTR      *  pszAreaList,
                  /* [in] */                    DWORD          dwNumSources,
                  /* [size_is][in] */           LPWSTR      *  pszSourceList
                  );
        
   STDMETHODIMP GetFilter(
                  /* [out] */                   DWORD       *  pdwEventType,
                  /* [out] */                   DWORD       *  pdwNumCategories,
                  /* [size_is][size_is][out] */ DWORD       ** ppdwEventCategories,
                  /* [out] */                   DWORD       *  pdwLowSeverity,
                  /* [out] */                   DWORD       *  pdwHighSeverity,
                  /* [out] */                   DWORD       *  pdwNumAreas,
                  /* [size_is][size_is][out] */ LPWSTR      ** ppszAreaList,
                  /* [out] */                   DWORD       *  pdwNumSources,
                  /* [size_is][size_is][out] */ LPWSTR      ** ppszSourceList
                  );
        
   STDMETHODIMP SelectReturnedAttributes(
                  /* [in] */                    DWORD          dwEventCategory,
                  /* [in] */                    DWORD          dwCount,
                  /* [size_is][in] */           DWORD       *  dwAttributeIDs
                  );
        
   STDMETHODIMP GetReturnedAttributes(
                  /* [in] */                    DWORD          dwEventCategory,
                  /* [out] */                   DWORD       *  pdwCount,
                  /* [size_is][size_is][out] */ DWORD       ** ppdwAttributeIDs
                  );
        
   STDMETHODIMP Refresh(
                  /* [in] */                    DWORD          dwConnection
                  );
        
   STDMETHODIMP CancelRefresh(
                  /* [in] */                    DWORD          dwConnection
                  );
        
   STDMETHODIMP GetState(
                  /* [out] */                   BOOL        *  pbActive,
                  /* [out] */                   DWORD       *  pdwBufferTime,
                  /* [out] */                   DWORD       *  pdwMaxSize,
                  /* [out] */                   OPCHANDLE   *  phClientSubscription
                  );
        
   STDMETHODIMP SetState(
                  /* [in][unique] */            BOOL        *  pbActive,
                  /* [in][unique] */            DWORD       *  pdwBufferTime,
                  /* [in][unique] */            DWORD       *  pdwMaxSize,
                  /* [in] */                    OPCHANDLE      hClientSubscription,
                  /* [out] */                   DWORD       *  pdwRevisedBufferTime,
                  /* [out] */                   DWORD       *  pdwRevisedMaxSize
                  );

   ///////////////////////////////////////////////////////////////////////////

      // Handles new events for this subscription
   HRESULT ProcessEvents( DWORD dwNumOfEvents, AeEvent** ppEvents );

// Impmementation
protected:
   // States
   BOOL                          m_fActive;
   DWORD                         m_dwBufferTime;
   DWORD                         m_dwMaxSize;
   OPCHANDLE                     m_hClientSubscription;
                                                // Use the critical section m_csStatesAndEventBuffer
                                                // to lock/unlock access to the 'States' data members.
   // Filters
   DWORD                         m_dwEventType;
   DWORD                         m_dwLowSeverity;
   DWORD                         m_dwHighSeverity;
   CSimpleArray<DWORD>           m_arCatIDs;
   CSimplePtrArray<WideString*> m_arAreas;
   CSimplePtrArray<WideString*> m_arSources;
   CComAutoCriticalSection       m_csFilters;   // lock/unlock all filters

   // Event Bufer
   CSimpleValQueue<AeEvent*>    m_EventBuffer;
   CComAutoCriticalSection       m_csStatesAndEventBuffer;
                                                // lock/unlock the event buffer
                                                // and all states
   // Selected Attributes
   typedef CSimpleArray<DWORD> AttrIDArray, *LPATTRIDARRAY;
   CSimpleMap<DWORD, LPATTRIDARRAY> m_mapSelectedAttrIDs;

   // Threads
   friend unsigned __stdcall EventRefreshThreadHandler( void* pCreator );
   HANDLE      m_hRefreshThread;                // Refresh Thread handle
   BOOL        m_fCancelRefresh;

   friend unsigned __stdcall EventNotificationThreadHandler( void* pCreator );
   HANDLE      m_hNotificationThread;           // Event Notification Thread
   HANDLE      m_hTerminateNotificationThread;
   HANDLE      m_hSendBufferedEvents;
   HRESULT     SendBufferedEvents();


   // Pure virtual function, must be implemented by derived classes
   virtual  HRESULT FireOnEvent( DWORD dwNumOfEvents, ONEVENTSTRUCT* pEvent,
                                 BOOL fRefresh = FALSE, BOOL fLastRefresh = FALSE ) = 0;

   BOOL     IsEventPassingFilters( AeEvent* pOnEvent );
   inline   HRESULT RefreshLastUpdateTime();

private:
   AeComBaseServer*              m_pServer;
   AeBaseServer*  m_pServerHandler;
};
//DOM-IGNORE-END

#endif // __EVENTSUBSCRIPTIONMGT_H
