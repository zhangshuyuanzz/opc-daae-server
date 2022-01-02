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
//-----------------------------------------------------------------------
// INLCUDE
//-----------------------------------------------------------------------
#include "stdafx.h"
#include <process.h>
#include "CoreMain.h"
#include "WideString.h"
#include "FixOutArray.h"
#include "AeSource.h"
#include "AeCategory.h"
#include "AeComSubscription.h"
#include "AeEvent.h"
#include "AeComSubscriptionManager.h"
#include "AeAttribute.h"
#include "MatchPattern.h"

//-------------------------------------------------------------------------
// STATIC MEMBERS
//-------------------------------------------------------------------------
DWORD AeComSubscriptionManager::DEFAULT_SIZE_EVENT_BUFFER = 100;


//-------------------------------------------------------------------------
// CODE
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AeComSubscriptionManager::AeComSubscriptionManager()
{
   m_pServer = NULL;
   m_pServerHandler = NULL;
   m_hRefreshThread = 0;
   m_fCancelRefresh = FALSE;
   m_hNotificationThread = 0;
   m_hTerminateNotificationThread = NULL;
   m_hSendBufferedEvents = NULL;

   // Set Default Filter
   m_dwEventType     = OPC_ALL_EVENTS;
   m_dwLowSeverity   = MIN_LOW_SEVERITY;
   m_dwHighSeverity  = 1000;

   // Set Default State
   m_fActive            = FALSE;
   m_dwBufferTime       = 0;
   m_dwMaxSize          = 0;
   m_hClientSubscription= NULL;
}



//=========================================================================
// Initializer
// -----------
//    Must be called after construction.
//    Initializes this instance and sets the default values. This instance 
//    is added to he list of the server and prevents deletion of
//    the server.
//=========================================================================
HRESULT AeComSubscriptionManager::Create(
   /* [in] */                    AeComBaseServer*  pServer,
   /* [in] */                    BOOL           bActive,
   /* [in] */                    DWORD          dwBufferTime,
   /* [in] */                    DWORD          dwMaxSize,
   /* [in] */                    OPCHANDLE      hClientSubscription,
   /* [out] */                   DWORD       *  pdwRevisedBufferTime,
   /* [out] */                   DWORD       *  pdwRevisedMaxSize )
{
   _ASSERTE( pServer != NULL );                 // Must not be NULL

   m_pServerHandler = ::OnGetAeBaseServer();
   if (m_pServerHandler == NULL) {
      return E_FAIL;
   }

   m_hTerminateNotificationThread = CreateEvent(
                           NULL,                // pointer to security attributes
                           FALSE,               // Automatically reset
                           FALSE,               // Nonsignaled initial state
                           NULL);               // pointer to event-object name 

   if (!m_hTerminateNotificationThread) {
      return HRESULT_FROM_WIN32( GetLastError() );
   }

   m_hSendBufferedEvents = CreateEvent(
                           NULL,                // pointer to security attributes
                           FALSE,               // Automatically reset
                           FALSE,               // Nonsignaled initial state
                           NULL);               // pointer to event-object name 

   if (!m_hSendBufferedEvents) {
      return HRESULT_FROM_WIN32( GetLastError() );
   }

   m_csStatesAndEventBuffer.Lock();
   HRESULT hres = m_EventBuffer.PreAllocate( DEFAULT_SIZE_EVENT_BUFFER );
   m_csStatesAndEventBuffer.Unlock();

   if (SUCCEEDED( hres )) {
      hres = SetState(  &bActive, &dwBufferTime, &dwMaxSize, hClientSubscription,
                        pdwRevisedBufferTime, pdwRevisedMaxSize );

      ResetEvent( m_hSendBufferedEvents );      // May be set into signaled state by SetState()
   }

   if (SUCCEEDED( hres )) {
      unsigned uCreationFlags = 0;
      unsigned uThreadAddr;

      m_hNotificationThread = (HANDLE)_beginthreadex(
                           NULL,                // No thread security attributes
                           0,                   // Default stack size  
                                                // Pointer to thread function 
                           EventNotificationThreadHandler,
                           this,                // Pass class to new thread 
                           uCreationFlags,      // Into ready state 
                           &uThreadAddr );      // Address of the thread
      
      hres = m_hNotificationThread ? S_OK : HRESULT_FROM_WIN32( GetLastError() );
   }

   if (SUCCEEDED( hres )) {
      hres = pServer->AddSubscriptionToList( this );
   }
   if (SUCCEEDED( hres )) {
      pServer->AddRef();                        // Prevents deletion of the server
      m_pServer = pServer;                      // Initialize member daGenericServer_ only
   }                                            // if AddSubscriptionToList() is successfully.
                                                // So the destructor can determine if the
                                                // remove function must be called.
   return hres;
}



//=========================================================================
// Destructor
// ----------
//    Do all cleanup stuff in function CleanupSubscription() !
//=========================================================================
AeComSubscriptionManager::~AeComSubscriptionManager()
{
}



//=========================================================================
// CleanupSubscription
// -------------------
//    Removes this insance from the list of the server and 
//    permits deletion of the server. Release all resources used by this
//    class instance.
//
//    Important:
//    Do all cleanup stuff in this function and not in the destructor !
//    Otherwise it's possible that a C Run-Time Error R6025 occurs.
//    This function is called by the function FinalRelease().
//
//    R6025:
//    The pure virtual funtion FireOnEvent() is called by the Event
//    Notification thread. If the virtual function is called from a base
//    class during the destruction of the base class then a 
//    C Run-Time Error R6025 occurs.
//
//    Therefore all the cleanup stuff should be done by this function.
//=========================================================================
void AeComSubscriptionManager::CleanupSubscription()
{
   int   i;
   BOOL  fReleaseServerRef = FALSE;

   // Remove current subscription
   if (m_pServer) {
      if (SUCCEEDED( m_pServer->RemoveSubscriptionFromList( this ) )) {
         fReleaseServerRef = TRUE;              // Permits deletion of the server but execute
      }                                         // it as the last action of the cleanup function.
   }

   if (m_hRefreshThread) {
      CancelRefresh( 0 );
   }

   if (m_hNotificationThread) {

      // Set the signal to shutdown the refresh thread.
      if (m_hTerminateNotificationThread) {
         SetEvent( m_hTerminateNotificationThread );
      }

      // Wait max 30 secs until the notification thread has terminated.
      if (WaitForSingleObject( m_hNotificationThread, 30000 ) == WAIT_TIMEOUT) {
         TerminateThread( m_hNotificationThread, 1 );
         CloseHandle( m_hNotificationThread );
      }
   }

   if (m_hTerminateNotificationThread) {
      CloseHandle( m_hTerminateNotificationThread );
   }

   if (m_hSendBufferedEvents) {
      CloseHandle( m_hSendBufferedEvents );
   }

   // Release non-fired events
   m_csStatesAndEventBuffer.Lock();                      
   for (DWORD z=0; z < m_EventBuffer.GetSize(); z++) {
      m_EventBuffer[z]->Release();
   }
   m_EventBuffer.RemoveAll();
   m_csStatesAndEventBuffer.Unlock();

   // Cleanup map of Attribute ID arrays
   AttrIDArray* parAttrIDs;
   for (i=0; i < m_mapSelectedAttrIDs.GetSize(); i++) {
      parAttrIDs = m_mapSelectedAttrIDs.m_aVal[i];
      if (parAttrIDs) delete parAttrIDs;
   }

   if (fReleaseServerRef) {
      m_pServer->Release();                     // Permits deletion of the server
   }
}



//-------------------------------------------------------------------------
// OPERATIONS
//-------------------------------------------------------------------------

///////////////////////////////////////////////////////////////////////////
////////////////////////// IOPCEventSubscriptionMgt ///////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCEventSubscriptionMgt::SetFilter                            INTERFACE
// -----------------------------------
//    Sets the filtering criteria to be used for tis event subscription.
//=========================================================================
STDMETHODIMP AeComSubscriptionManager::SetFilter(
   /* [in] */                    DWORD          dwEventType,
   /* [in] */                    DWORD          dwNumCategories,
   /* [size_is][in] */           DWORD       *  pdwEventCategories,
   /* [in] */                    DWORD          dwLowSeverity,
   /* [in] */                    DWORD          dwHighSeverity,
   /* [in] */                    DWORD          dwNumAreas,
   /* [size_is][in] */           LPWSTR      *  pszAreaList,
   /* [in] */                    DWORD          dwNumSources,
   /* [size_is][in] */           LPWSTR      *  pszSourceList )
{
   // Check the parameters
   DWORD    i;

   if (dwEventType == 0) {
      return E_INVALIDARG;
   }
   if (dwEventType & ~OPC_ALL_EVENTS) {
      return E_INVALIDARG;
   }
   if (dwHighSeverity < MIN_LOW_SEVERITY || dwHighSeverity > 1000) {
      return E_INVALIDARG;
   }

   // 23-jan-2002 MT V3.2
   // V1.03 of the OPC AE specification defines :
   //    The Server is responsible for mapping its internal severity levels
   //    to evenly span the 1..1000 range. Clients that wish to receive events
   //    of all severities should set dwLowSeverity=1 and dwHighSeverity=1000.
   if ((dwLowSeverity < MIN_LOW_SEVERITY) || (dwLowSeverity > dwHighSeverity)) {
      return E_INVALIDARG;
   }

   for (i=0; i < dwNumCategories; i++) {
      if (!m_pServerHandler->ExistEventCategory( pdwEventCategories[i] )) {
         return E_INVALIDARG;
      }
   }

   for (i=0; i < dwNumAreas; i++) {
      if (!m_pServerHandler->ExistArea( pszAreaList[i] )) {
         return E_INVALIDARG;
      }
   }

   for (i=0; i < dwNumSources; i++) {
      if (!m_pServerHandler->ExistSource( pszSourceList[i] )) {
         return E_INVALIDARG;
      }
   }

   if (m_hRefreshThread) {
      return OPC_E_BUSY;
   }

   // Set all new filters
   DWORD          hres = S_OK;
   WideString*   pwsz;

   m_csFilters.Lock();
   try {

      m_dwEventType     = dwEventType;
      m_dwLowSeverity   = dwLowSeverity;
      m_dwHighSeverity  = dwHighSeverity;

      m_arCatIDs.RemoveAll();
      if (dwNumCategories) {
         for (i=0; i < dwNumCategories; i++) {
            if (!m_arCatIDs.Add( pdwEventCategories[i] ))
               throw E_OUTOFMEMORY;
         }
      }

      m_arAreas.RemoveAll();
      for (i=0; i < dwNumAreas; i++) {
         pwsz = new WideString;
         if (!pwsz) throw E_OUTOFMEMORY;
         hres = pwsz->SetString( pszAreaList[i] );
         if (FAILED( hres )) throw hres;
         if (!m_arAreas.Add( pwsz )) throw E_OUTOFMEMORY;
      }

      m_arSources.RemoveAll();
      for (i=0; i < dwNumSources; i++) {
         pwsz = new WideString;
         if (!pwsz) throw E_OUTOFMEMORY;
         hres = pwsz->SetString( pszSourceList[i] );
         if (FAILED( hres )) throw hres;
         if (!m_arSources.Add( pwsz )) throw E_OUTOFMEMORY;
      }

      hres = S_OK;
   }
   catch (HRESULT hresEx) {
      hres = hresEx;
   }
   m_csFilters.Unlock();

   return hres;
}


     
//=========================================================================
// IOPCEventSubscriptionMgt::GetFilter                            INTERFACE
// -----------------------------------
//    Returns the current filter of this event subscription.
//=========================================================================
STDMETHODIMP AeComSubscriptionManager::GetFilter(
   /* [out] */                   DWORD       *  pdwEventType,
   /* [out] */                   DWORD       *  pdwNumCategories,
   /* [size_is][size_is][out] */ DWORD       ** ppdwEventCategories,
   /* [out] */                   DWORD       *  pdwLowSeverity,
   /* [out] */                   DWORD       *  pdwHighSeverity,
   /* [out] */                   DWORD       *  pdwNumAreas,
   /* [size_is][size_is][out] */ LPWSTR      ** ppszAreaList,
   /* [out] */                   DWORD       *  pdwNumSources,
   /* [size_is][size_is][out] */ LPWSTR      ** ppszSourceList )
{
   m_csFilters.Lock();

   *pdwEventType     = m_dwEventType;
   *pdwLowSeverity   = m_dwLowSeverity;
   *pdwHighSeverity  = m_dwHighSeverity;

   *pdwNumCategories = m_arCatIDs.GetSize();
   *pdwNumAreas      = m_arAreas.GetSize();
   *pdwNumSources    = m_arSources.GetSize();

   *ppdwEventCategories = NULL;                 // Note : Proxy/Stub checks if the pointers are NULL
   *ppszAreaList        = NULL;
   *ppszSourceList      = NULL;

   CFixOutArray< DWORD >   aCatIDs;
   CFixOutArray< LPWSTR >  aAreas;
   CFixOutArray< LPWSTR >  aSources;

   DWORD    i;
   HRESULT  hres = S_OK;

   try {                                        // Calculate the size of the result arrays
      // 22-jan-2002 MT
      // Return a NULL instead of an empty array for Categories, Areas and Sources
      // (changed in OPC AE spec 1.03).
      if (*pdwNumCategories)  aCatIDs.Init( *pdwNumCategories, ppdwEventCategories );
      if (*pdwNumAreas)       aAreas.Init( *pdwNumAreas ,ppszAreaList );
      if (*pdwNumSources)     aSources.Init( *pdwNumSources, ppszSourceList );

      for (i=0; i<*pdwNumCategories; i++) {
         aCatIDs[i] = m_arCatIDs[i];
      }
      for (i=0; i<*pdwNumAreas; i++) {
         aAreas[i] = m_arAreas[i]->CopyCOM();
         if (!aAreas[i]) throw E_OUTOFMEMORY;
      }
      for (i=0; i<*pdwNumSources; i++) {
         aSources[i] = m_arSources[i]->CopyCOM();
         if (!aSources[i]) throw E_OUTOFMEMORY;
      }
   }
   catch (HRESULT hresEx) {
      aCatIDs.Cleanup();
      aAreas.Cleanup();
      aSources.Cleanup();
      hres = hresEx;
   }

   m_csFilters.Unlock();
   return hres;
}


     
//=========================================================================
// IOPCEventSubscriptionMgt::SelectReturnedAttributes             INTERFACE
// --------------------------------------------------
//    Sets the attributes to be returned with event notifications.
//=========================================================================
STDMETHODIMP AeComSubscriptionManager::SelectReturnedAttributes(
   /* [in] */                    DWORD          dwEventCategory,
   /* [in] */                    DWORD          dwCount,
   /* [size_is][in] */           DWORD       *  dwAttributeIDs )
{
   DWORD i;

   if (!m_pServerHandler->ExistEventCategory( dwEventCategory )) {
      return E_INVALIDARG;
   }
   for (i=0; i < dwCount; i++) {
      if (!m_pServerHandler->ExistEventAttribute( dwEventCategory, dwAttributeIDs[i] )) {
         return E_INVALIDARG;
      }
   }
                                                // get the array of IDs associated with the specified category
   LPATTRIDARRAY parAttrIDs = m_mapSelectedAttrIDs.Lookup( dwEventCategory );
   if (parAttrIDs) {
      parAttrIDs->RemoveAll();                  // remove all currently selected attribute IDs
   }
   else {                                       // first selection for the specified category
      parAttrIDs = ::new AttrIDArray;           // ::new guarantees that the new operator of global scope is used
      if (!parAttrIDs) return E_OUTOFMEMORY;

      if (!m_mapSelectedAttrIDs.Add( dwEventCategory, parAttrIDs )) {
         delete parAttrIDs;
         return E_OUTOFMEMORY;
      }
   }
                                                // add all specified attribute IDs to the array
   for (i=0; i < dwCount; i++) {
      if (!parAttrIDs->Add( dwAttributeIDs[i] )) {
         m_mapSelectedAttrIDs.Remove( dwEventCategory );
         delete parAttrIDs;
         return E_OUTOFMEMORY;
      }
   }
   return S_OK;
}


     
//=========================================================================
// IOPCEventSubscriptionMgt::GetReturnedAttributes                INTERFACE
// -----------------------------------------------
//    Gets the ID of attributes which are currently specified to be
//    returned with event notifications.
//=========================================================================
STDMETHODIMP AeComSubscriptionManager::GetReturnedAttributes(
   /* [in] */                    DWORD          dwEventCategory,
   /* [out] */                   DWORD       *  pdwCount,
   /* [size_is][size_is][out] */ DWORD       ** ppdwAttributeIDs )
{
   if (!m_pServerHandler->ExistEventCategory( dwEventCategory )) {
      return E_INVALIDARG;
   }
                                                // get the array of IDs associated with the specified category
   AttrIDArray* parAttrIDs = m_mapSelectedAttrIDs.Lookup( dwEventCategory );

                                                // the count is 0 if there is no array for the specified category
   *pdwCount = (parAttrIDs) ? parAttrIDs->GetSize() : 0;
   *ppdwAttributeIDs = ComAlloc<DWORD>(*pdwCount);
   if (*ppdwAttributeIDs == NULL) {
      *pdwCount = 0;
      return E_OUTOFMEMORY;
   }
   for (DWORD i=0; i < *pdwCount ; i++) {       // return the IDs of all selected attributes of the specified category
      (*ppdwAttributeIDs)[i] = (*parAttrIDs)[i];
   }
   return S_OK;
}



//=========================================================================
// IOPCEventSubscriptionMgt::Refresh                              INTERFACE
// ---------------------------------
//    Forces a refresh for all active and inactive, unacknowledged
//    conditions.
//    This function creates the Event Refresh Thread.
//=========================================================================
STDMETHODIMP AeComSubscriptionManager::Refresh(
   /* [in] */                    DWORD          dwConnection )
{
   HRESULT hres;

   if (m_hRefreshThread) {
      hres = OPC_E_BUSY;                        // Another refresh is in progress
   }
   else {

      unsigned uCreationFlags = 0;
      unsigned uThreadAddr;
      
      m_hRefreshThread = (HANDLE)_beginthreadex(
                           NULL,                // No thread security attributes
                           0,                   // Default stack size  
                                                // Pointer to thread function 
                           EventRefreshThreadHandler,
                           this,                // Pass class to new thread 
                           uCreationFlags,      // Into ready state 
                           &uThreadAddr );      // Address of the thread
      
      hres = m_hRefreshThread ? S_OK : HRESULT_FROM_WIN32( GetLastError() );
   }
   return hres;
}



//=========================================================================
// IOPCEventSubscriptionMgt::CancelRefresh                        INTERFACE
// ---------------------------------------
//    Cancels a refresh in progress for the event subscription.
//    This function kills the Event Refresh Thread.
//=========================================================================
STDMETHODIMP AeComSubscriptionManager::CancelRefresh(
   /* [in] */                    DWORD          dwConnection )
{
   if (m_hRefreshThread) {

      // Set the signal to shutdown the refresh thread.
      m_fCancelRefresh = TRUE;

      // Wait max 30 secs until the refresh thread has terminated.
      if (WaitForSingleObject( m_hRefreshThread, 30000 ) == WAIT_TIMEOUT) {
         TerminateThread( m_hRefreshThread, 1 );
         CloseHandle( m_hRefreshThread );
      }

      // Send allways the final callback with count=0
      AeSubscribedEvent  dummyevent;             // Note : Ptr for events must not be NULL
      FireOnEvent( 0, &dummyevent, TRUE, TRUE );

      m_fCancelRefresh = FALSE;
      m_hRefreshThread = 0;
   }
   else {
      return E_FAIL;
   }
   return S_OK;
}



//=========================================================================
// IOPCEventSubscriptionMgt::GetState                             INTERFACE
// ----------------------------------
//    Gets the current state of the subscription.
//=========================================================================
STDMETHODIMP AeComSubscriptionManager::GetState(
   /* [out] */                   BOOL        *  pbActive,
   /* [out] */                   DWORD       *  pdwBufferTime,
   /* [out] */                   DWORD       *  pdwMaxSize,
   /* [out] */                   OPCHANDLE   *  phClientSubscription )
{
                                                // Note : Proxy/Stub checks if the pointers are NULL
   m_csStatesAndEventBuffer.Lock();
   *pbActive               = m_fActive;
   *pdwBufferTime          = m_dwBufferTime;
   *pdwMaxSize             = m_dwMaxSize;
   *phClientSubscription   = m_hClientSubscription;
   m_csStatesAndEventBuffer.Unlock();
   return S_OK;
}



//=========================================================================
// IOPCEventSubscriptionMgt::SetState                             INTERFACE
// ----------------------------------
//    Sets the new state of the subscription.
//=========================================================================
STDMETHODIMP AeComSubscriptionManager::SetState(
   /* [in][unique] */            BOOL        *  pbActive,
   /* [in][unique] */            DWORD       *  pdwBufferTime,
   /* [in][unique] */            DWORD       *  pdwMaxSize,
   /* [in] */                    OPCHANDLE      hClientSubscription,
   /* [out] */                   DWORD       *  pdwRevisedBufferTime,
   /* [out] */                   DWORD       *  pdwRevisedMaxSize )
{
   HRESULT  hres = S_OK;
   BOOL     fNotifyEventThread = FALSE;         // The notification thread must be notified 
                                                // if subscription states changes
   m_csStatesAndEventBuffer.Lock();

   if (pbActive && (m_fActive != *pbActive)) {
      m_fActive = *pbActive;
      if (!m_fActive) {                         // Remove all buffered events if the new state is inactive
         for (DWORD i=0; i < m_EventBuffer.GetSize(); i++) {
            m_EventBuffer[i]->Release();
         }
         m_EventBuffer.RemoveAll();
      }
      fNotifyEventThread = TRUE;
   }
   if (pdwBufferTime && (m_dwBufferTime != *pdwBufferTime)) {
      m_dwBufferTime = *pdwBufferTime;
      fNotifyEventThread = TRUE;
   }
   if (pdwMaxSize && (m_dwMaxSize != *pdwMaxSize)) {
      m_dwMaxSize = *pdwMaxSize;

      if ((m_dwMaxSize + 10) > DEFAULT_SIZE_EVENT_BUFFER) {
         hres = m_EventBuffer.PreAllocate( m_dwMaxSize + 10 );
      }
      if (SUCCEEDED( hres )) {
         if (m_EventBuffer.GetSize() >= m_dwMaxSize) {
            fNotifyEventThread = TRUE;
         }
      }
   }

   m_hClientSubscription = hClientSubscription;
   *pdwRevisedBufferTime = m_dwBufferTime;      // Note : Proxy/Stub checks if the pointers are NULL
   *pdwRevisedMaxSize =    m_dwMaxSize;
   
   if (fNotifyEventThread) {
      SetEvent( m_hSendBufferedEvents );        // Notify the Event Notification Thread
   }                                            // that there is an new buffer time

   m_csStatesAndEventBuffer.Unlock();
   return hres;
}



///////////////////////////////////////////////////////////////////////////

//=========================================================================
// ProcessEvents                                                     PUBLIC
// -------------
//    Add the events to the event buffer if the filters of the
//    subscription are passed. The Event Notification Thread is
//    activated if the events must be sent immediately
//    (BufferTime = 0 or Number of events in
//    buffer >= MaxSize if MaxSize is specified).
//=========================================================================
HRESULT AeComSubscriptionManager::ProcessEvents( DWORD dwNumOfEvents, AeEvent** ppEvents )
{
   DWORD       i;
   HRESULT     hres = S_OK;
   AeEvent*   pEvent;

   m_csStatesAndEventBuffer.Lock();

   if (m_fActive) {                             // Subscription must be active

      for (i=0; i < dwNumOfEvents; i++) {       // Handle all new events
         pEvent = ppEvents[i];
      
         if (!IsEventPassingFilters( pEvent ))
            continue;
      
         pEvent->AddRef();                      // Event passed all filters. Is
                                                // used once more
                                                
         hres = m_EventBuffer.Add( pEvent );    // Add Event to event buffer
         if (FAILED( hres ))                       
            break;                                 
      }                                            
                                                // Check if the events in the buffer
                                                // must be sent immediately
      if (m_dwBufferTime == 0 || 
         (m_dwMaxSize && (m_EventBuffer.GetSize() >= m_dwMaxSize))) {
         SetEvent( m_hSendBufferedEvents );     // Send the events immediately
      }
   } // Subscription is active

   m_csStatesAndEventBuffer.Unlock();
   return hres;
}



//-------------------------------------------------------------------------
// IMPLEMENTATTION
//-------------------------------------------------------------------------

//=========================================================================
// IsEventPassingFilters                                           INTERNAL
// ---------------------
//    Checks if the specified event passes the filters of
//    the subscription.
//    
// returns:
//    FALSE    event is filtered
//    TRUE     all filters passed
//=========================================================================
BOOL AeComSubscriptionManager::IsEventPassingFilters( AeEvent* pOnEvent )
{
   DWORD dwSize;
   BOOL  fPassed = TRUE;

   m_csFilters.Lock();
   try {
      //
      // Event Type & Severity Filters
      //
      if (!(pOnEvent->dwEventType & m_dwEventType)) throw FALSE;
      // V3.0
      // if (pOnEvent->dwSeverity < m_dwLowSeverity) throw FALSE;

      // V3.1, 0 means Severity Filter OFF
      if (m_dwLowSeverity) {
         if (pOnEvent->dwSeverity < m_dwLowSeverity) throw FALSE;
      }

      if (pOnEvent->dwSeverity > m_dwHighSeverity) throw FALSE;
      //
      // Category Filter
      //
      if (dwSize = m_arCatIDs.GetSize()) {
         if (m_arCatIDs.Find( pOnEvent->dwEventCategory ) == -1)
            throw FALSE;
      }
      //
      // Area Filter
      //
      if (dwSize = m_arAreas.GetSize()) {
         BOOL fAreaOK = FALSE;
         LPVARIANT pAreas = pOnEvent->LookupAttributeValue( ATTRID_AREAS );
         if (pAreas) {
                                                // Must be an array of BSTRs
            _ASSERTE( V_VT( pAreas ) == (VT_ARRAY | VT_BSTR) );  

            HRESULT     hres;
            BSTR HUGEP* pbstr;
                                                   // Get a pointer to the elements of the array.
            if (SUCCEEDED( SafeArrayAccessData( V_ARRAY( pAreas ), (void HUGEP**)&pbstr ) )) {
            
               do {
                  dwSize--;
                  for (DWORD i = 0; i < V_ARRAY( pAreas )->rgsabound->cElements; i++) {
                     if (wcscmp( (LPCWSTR)*m_arAreas[dwSize], pbstr[i]) == 0) {
                        fAreaOK = TRUE;
                        break;
                     }
                  }
               } while (dwSize && !fAreaOK);
            
               hres = SafeArrayUnaccessData( V_ARRAY( pAreas ) );
               _ASSERTE( SUCCEEDED( hres ) );
            }
         }
         if (!fAreaOK) throw FALSE;
      }
      //
      // Source Filter
      //
      if (dwSize = m_arSources.GetSize()) {
         BOOL fSourceOK = FALSE;
         do {
            dwSize--;
            if (MatchPattern( pOnEvent->szSource, (LPCWSTR)*m_arSources[dwSize] )) {
               fSourceOK = TRUE;
               break;
            }
         } while (dwSize);
         if (!fSourceOK) throw FALSE;
      }
   }
   catch(BOOL fPassedEx) {
      fPassed = fPassedEx;                   // not all filters passed
   }
   m_csFilters.Unlock();

   return fPassed;
}



//=========================================================================
// SendBufferedEvents                                              INTERNAL
// ------------------
//    Send events from the event buffer to the client.
//    If no MaxSize is defined then all events from the buffer are
//    sent; otherwise the maximum number of events limited by MaxSize.
//=========================================================================
HRESULT AeComSubscriptionManager::SendBufferedEvents()
{
   DWORD             dwNumOfEvents, i;
   AeEvent*         pEvent;
   AeSubscribedEvent* pSubscrEvents;
   AttrIDArray*      parAttrIDs;

   m_csStatesAndEventBuffer.Lock();
   if (m_dwMaxSize) {
      dwNumOfEvents = min( m_EventBuffer.GetSize(), m_dwMaxSize );
   }
   else {
      dwNumOfEvents = m_EventBuffer.GetSize();
   }
   m_csStatesAndEventBuffer.Unlock();

   if (dwNumOfEvents == 0) {                    // There is nothing to send
      return S_OK;
   }

   pSubscrEvents = new AeSubscribedEvent [dwNumOfEvents];
   if (!pSubscrEvents) {
      return E_OUTOFMEMORY;
   }

   m_csStatesAndEventBuffer.Lock();
   for (i=0; i < dwNumOfEvents; i++) {

      pEvent = m_EventBuffer[i];

      parAttrIDs = m_mapSelectedAttrIDs.Lookup( pEvent->dwEventCategory );

      if (parAttrIDs)                        // Create the events with the selected attributes
         pSubscrEvents[ i ].Create( pEvent, parAttrIDs->GetSize(), parAttrIDs->m_aT );
      else
         pSubscrEvents[ i ].Create( pEvent, 0, NULL );
   }
   m_csStatesAndEventBuffer.Unlock();

   FireOnEvent( dwNumOfEvents, pSubscrEvents );

   delete [] pSubscrEvents;

   m_csStatesAndEventBuffer.Lock();
   for (i=0; i < dwNumOfEvents; i++) {
      m_EventBuffer[i]->Release();
   }
   m_EventBuffer.RemoveFirstN( dwNumOfEvents );
   m_csStatesAndEventBuffer.Unlock();

   return S_OK;
}



//-------------------------------------------------------------------------
// THREAD HANDLERS
//-------------------------------------------------------------------------

//=========================================================================
// EventRefreshThreadHandler                                       INTERNAL
// -------------------------
//    Sends all active and inactive, unacknowledged conditions whose event
//    notification match the filter of the subscription.
//    Checks the 'cancel refresh' flag if in progress.
//    This thread must send an event notification in any case even if
//    an error occurs.
//=========================================================================
unsigned __stdcall EventRefreshThreadHandler( void* pCreator )
{
   AeComSubscriptionManager* pSubscr = static_cast<AeComSubscriptionManager *>(pCreator);
   _ASSERTE( pSubscr );

   int                     i;
   HRESULT                 hres;
   DWORD                   dwNumOfEvents = 0;   // Number of events to send
   AeSubscribedEvent*       pSubscrEvents = NULL;
   CSimpleArray<AeEvent*> arEventPtrs;
   AeSubscribedEvent        dummyevent;          // Note : Ptr for events must not be NULL

                                                // Get all active and inactive, unacknowledged conditions
                                                // This function checks the 'cancel refresh' flag
   hres = pSubscr->m_pServerHandler->GetEventsForRefresh( pSubscr, arEventPtrs );
   pSubscrEvents = new AeSubscribedEvent [arEventPtrs.GetSize()];

   if (FAILED( hres ) || !pSubscrEvents) {
                                                // send one final callback
      pSubscr->FireOnEvent( 0, &dummyevent, TRUE, TRUE );
   }
   else {
                                                // Filter the events
      for (i=0; i < arEventPtrs.GetSize(); i++) {

         if (pSubscr->m_fCancelRefresh)
            break;                              // Cancel Refresh Request

         if (pSubscr->IsEventPassingFilters( arEventPtrs[i] )) {

            AeComSubscriptionManager::AttrIDArray* parAttrIDs = pSubscr->m_mapSelectedAttrIDs.Lookup( arEventPtrs[i]->dwEventCategory );
                                          
            if (parAttrIDs)                     // Create the events with the selected attributes
               pSubscrEvents[ dwNumOfEvents ].Create( arEventPtrs[i], parAttrIDs->GetSize(), parAttrIDs->m_aT );
            else
               pSubscrEvents[ dwNumOfEvents ].Create( arEventPtrs[i], 0, NULL );

            dwNumOfEvents++;
         }

      }
      if (!pSubscr->m_fCancelRefresh) {
                                                // Send the events to the client
         if (pSubscr->m_dwMaxSize) {            // There is a maximum number of events for a single callback defined
         
            i = dwNumOfEvents;
            DWORD dwMaxSize = pSubscr->m_dwMaxSize;

            while (((DWORD)i > dwMaxSize) && !pSubscr->m_fCancelRefresh) {
               pSubscr->FireOnEvent( dwMaxSize,
                                     &pSubscrEvents[dwNumOfEvents-i], TRUE, FALSE );
               i -= dwMaxSize;
               // MaxSize value maybe changed during refresh
               dwMaxSize = pSubscr->m_dwMaxSize;
            }
            if (!pSubscr->m_fCancelRefresh) {
               pSubscr->FireOnEvent( i, &pSubscrEvents[dwNumOfEvents-i], TRUE, TRUE );
            }
         }
         else {                                 // No limitations for a single callback defined
            pSubscr->FireOnEvent( dwNumOfEvents, pSubscrEvents, TRUE, TRUE );
         }
      }
   }                                            // Release all
   for (i=0; i < arEventPtrs.GetSize(); i++) {
      arEventPtrs[i]->Release();
   }
   arEventPtrs.RemoveAll();
   if (pSubscrEvents) {
      delete [] pSubscrEvents;
   }

   CloseHandle( pSubscr->m_hRefreshThread );    // Thread handle must be closed explicitly
   pSubscr->m_fCancelRefresh = FALSE;
   pSubscr->m_hRefreshThread = 0;
   _endthreadex( 0 );                           // if _endthreadex() is used.
   return 0;                                    // Should never get here
}



//=========================================================================
// EventNotificationThreadHandler                                  INTERNAL
// ------------------------------
//    This thread send buffered events to the client if activated
//    by timeout or by functions for the event buffer handling.
//=========================================================================
unsigned __stdcall EventNotificationThreadHandler( void* pCreator )
{
   AeComSubscriptionManager* pSubscr = static_cast<AeComSubscriptionManager *>(pCreator);
   _ASSERTE( pSubscr );

   DWORD    dwState;
   HANDLE   ahObject[2];

   ahObject[0] = pSubscr->m_hTerminateNotificationThread;
   ahObject[1] = pSubscr->m_hSendBufferedEvents;

   DWORD dwTimeoutInterv;
   BOOL  fTerminate = FALSE;
   do {

      dwTimeoutInterv = INFINITE;
      pSubscr->m_csStatesAndEventBuffer.Lock();
      if (pSubscr->m_fActive && pSubscr->m_dwBufferTime) {
         dwTimeoutInterv = pSubscr->m_dwBufferTime;
      }
      pSubscr->m_csStatesAndEventBuffer.Unlock();

      dwState = WaitForMultipleObjects(
                     2,                         // number of handles in array
                     ahObject,                  // object-handle array
                     FALSE,                     // wait until one state is signaled
                     dwTimeoutInterv );         // time-out interval

      switch (dwState) {

         case WAIT_OBJECT_0:                    // Terminate
            fTerminate = TRUE;
            break;

         case WAIT_OBJECT_0+1:                  // New Buffer Time, new events if no
            pSubscr->SendBufferedEvents();      // BufferTime or MaxSize of events in buffer
            break;

         case WAIT_TIMEOUT:                     // BufferTime run out
            pSubscr->SendBufferedEvents();      // Send existing events
            break;

         case WAIT_FAILED:
            break;
      }
   } while (!fTerminate);


                                                // Thread handle must be closed explicitly
   CloseHandle( pSubscr->m_hNotificationThread );
   pSubscr->m_hNotificationThread = NULL;
   _endthreadex( 0 );                           // if _endthreadex() is used.
   return 0;                                    // Should never get here
}



//=========================================================================
// RefreshLastUpdateTime                                             INLINE
// ---------------------
//    Updates the 'Last Update Time' with the current time.
//=========================================================================
HRESULT AeComSubscriptionManager::RefreshLastUpdateTime()
{
   return CoFileTimeNow( &m_pServer->m_ftLastUpdateTime );
}



//-------------------------------------------------------------------------
// CODE AeComSubscription
//-------------------------------------------------------------------------

//=========================================================================
// FireOnEven                                                      INTERNAL
// ----------
//    Fires the specified number of events via the IOPCEventSink
//    interface to the subscribed client.
//=========================================================================
HRESULT AeComSubscription::FireOnEvent( DWORD dwNumOfEvents, ONEVENTSTRUCT* pEvent,
                                            BOOL fRefresh /*= FALSE */, BOOL fLastRefresh /*= FALSE */ )
{
   // Invoke the callback
   IOPCEventSink* pSink;
   HRESULT hres = GetEventSinkInterface( &pSink );
   if (SUCCEEDED( hres )) {

      hres = pSink->OnEvent(
                        m_hClientSubscription,
                        fRefresh,
                        fLastRefresh,
                        dwNumOfEvents,
                        pEvent );
                                             // Refresh the last update time even
      hres = RefreshLastUpdateTime();        // if OnEvent() failed.
      pSink->Release();
   }
   return hres;
}



//=========================================================================
// GetEventSinkInterface                                           INTERNAL
// ---------------------
//    Returns an IOPCEventSink interface pointer from the connection
//    point list or from the Global Interface Table.
//        
//    The GetEventSinkInterface() method calls IUnknown::AddRef() on the
//    pointer obtained in the ppSink parameter. It is the caller's
//    responsibility to call Release() on this pointer.
//
// Return Code:
//    S_OK                    All succeeded
//    CONNECT_E_NOCONNECTION  There is no connection established
//                            (IConnectionPoint::Advise not called)
//    E_xxx                   Error occured
//=========================================================================
HRESULT AeComSubscription::GetEventSinkInterface( IOPCEventSink** ppSink )
{
   HRESULT hres = CONNECT_E_NOCONNECTION;          // There is no registered event sink

   *ppSink = NULL;

   Lock();                                         // Lock the connection point list

   IUnknown** pp = m_vec.begin();                  // There can be only one registered event sink.
   if (*pp) {
      hres = core_generic_main.m_pGIT->GetInterfaceFromGlobal( m_dwCookieGITEventSink, IID_IOPCEventSink, (LPVOID*)ppSink );
   }

   Unlock();                                       // Unlock the connection point list
   return hres;
}



//=========================================================================
// IConnectionPoint::Advise                                       INTERFACE
// ------------------------
//    Overriden ATL implementation of IConnectionPoint::Advise().
//    This function also registers the client's advise sink in the Global
//    Interface Table if required.
//=========================================================================
STDMETHODIMP AeComSubscription::Advise( IUnknown* pUnkSink, DWORD* pdwCookie )
{
   Lock();                                // Lock the connection point list

                                          // Call the base class member
   HRESULT hres = IOPCEventSubscriptionConnectionPointImpl::Advise( pUnkSink, pdwCookie );
   if (SUCCEEDED( hres )) {
                                          // Register the callback interface in the global interface table
      hres = core_generic_main.m_pGIT->RegisterInterfaceInGlobal( pUnkSink, IID_IOPCEventSink, &m_dwCookieGITEventSink );
      if (FAILED( hres )) {
         IOPCEventSubscriptionConnectionPointImpl::Unadvise( *pdwCookie );
         Unadvise( *pdwCookie );
         pUnkSink->Release();             // Note :   register increments the refcount
      }                                   //          even if the function failed
   }
   Unlock();                              // Unlock the connection point list
   return hres;
}



//=========================================================================
// IConnectionPoint::Unadvise                                     INTERFACE
// --------------------------
//    Overriden ATL implementation of IConnectionPoint::Unadvise().
//    This function also removes the client's advise sink from the Global
//    Interface Table if required.
//=========================================================================
STDMETHODIMP AeComSubscription::Unadvise( DWORD dwCookie )
{
   Lock();                                      // Lock the connection point list

   HRESULT hresGIT = S_OK;
   hresGIT = core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( m_dwCookieGITEventSink );
                                                // Call the base class member
   HRESULT hres = IOPCEventSubscriptionConnectionPointImpl::Unadvise( dwCookie );

   Unlock();                                    // Unlock the connection point list

   if (FAILED( hres )) {
      return hres;
   }
   if (FAILED( hresGIT )) {
      return hresGIT;
   }
   return hres;
}
//DOM-IGNORE-END
#endif