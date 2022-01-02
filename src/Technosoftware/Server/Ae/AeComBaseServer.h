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

#ifndef __EVENTSERVER_H
#define __EVENTSERVER_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

class AeBaseServer;
class AeComSubscriptionManager;
class AeEvent;


//-----------------------------------------------------------------------
// CLASS
//-----------------------------------------------------------------------
class AeComBaseServer : public IOPCEventServer
{
         // This friend declaration allows the class AeComSubscriptionManager
         // to add itself to the list of this server using
         // AddSubscriptionToList() and RemoveSubscriptionFromList()
         // which is private and should not be accessed by other classes.
   friend class AeComSubscriptionManager;

// Construction / Destruction
public:
   AeComBaseServer();
   HRESULT  Create();
   ~AeComBaseServer();

// Operations
public:
   ///////////////////////////////////////////////////////////////////////////
   /////////////////////////////// IOPCEventServer ///////////////////////////
   ///////////////////////////////////////////////////////////////////////////

   STDMETHODIMP GetStatus(
                  /* [out] */                   OPCEVENTSERVERSTATUS
                                                            ** ppEventServerStatus
                  );

   STDMETHODIMP CreateEventSubscription(
                  /* [in] */                    BOOL           bActive,
                  /* [in] */                    DWORD          dwBufferTime,
                  /* [in] */                    DWORD          dwMaxSize,
                  /* [in] */                    OPCHANDLE      hClientSubscription,
                  /* [in] */                    REFIID         riid,
                  /* [iid_is][out] */           LPUNKNOWN   *  ppUnk,
                  /* [out] */                   DWORD       *  pdwRevisedBufferTime,
                  /* [out] */                   DWORD       *  pdwRevisedMaxSize
                  );                                              
                                                                  
   STDMETHODIMP QueryAvailableFilters(                            
                  /* [out] */                   DWORD       *  pdwFilterMask
                  );                                        
                                                            
   STDMETHODIMP QueryEventCategories(                       
                  /* [in] */                    DWORD          dwEventType,
                  /* [out] */                   DWORD       *  pdwCount,
                  /* [size_is][size_is][out] */ DWORD       ** ppdwEventCategories,
                  /* [size_is][size_is][out] */ LPWSTR      ** ppszEventCategoryDescs
                  );                                        
                                                            
   STDMETHODIMP QueryConditionNames(                        
                  /* [in] */                    DWORD          dwEventCategory,
                  /* [out] */                   DWORD       *  pdwCount,
                  /* [size_is][size_is][out] */ LPWSTR      ** ppszConditionNames
                  );                                        
                                                            
   STDMETHODIMP QuerySubConditionNames(                     
                  /* [in] */                    LPWSTR         szConditionName,
                  /* [out] */                   DWORD       *  pdwCount,
                  /* [size_is][size_is][out] */ LPWSTR      ** ppszSubConditionNames
                  );                                        
                                                            
   STDMETHODIMP QuerySourceConditions(                      
                  /* [in] */                    LPWSTR         szSource,
                  /* [out] */                   DWORD       *  pdwCount,
                  /* [size_is][size_is][out] */ LPWSTR      ** ppszConditionNames
                  );                                        
                                                            
   STDMETHODIMP QueryEventAttributes(                       
                  /* [in] */                    DWORD          dwEventCategory,
                  /* [out] */                   DWORD       *  pdwCount,
                  /* [size_is][size_is][out] */ DWORD       ** ppdwAttrIDs,
                  /* [size_is][size_is][out] */ LPWSTR      ** ppszAttrDescs,
                  /* [size_is][size_is][out] */ VARTYPE     ** ppvtAttrTypes
                  );                                              
                                                                  
   STDMETHODIMP TranslateToItemIDs(                               
                  /* [in] */                    LPWSTR         szSource,
                  /* [in] */                    DWORD          dwEventCategory,
                  /* [in] */                    LPWSTR         szConditionName,
                  /* [in] */                    LPWSTR         szSubconditionName,
                  /* [in] */                    DWORD          dwCount,
                  /* [size_is][in] */           DWORD       *  pdwAssocAttrIDs,
                  /* [size_is][size_is][out] */ LPWSTR      ** ppszAttrItemIDs,
                  /* [size_is][size_is][out] */ LPWSTR      ** ppszNodeNames,
                  /* [size_is][size_is][out] */ CLSID       ** ppCLSIDs
                  );

   STDMETHODIMP GetConditionState(
                  /* [in] */                    LPWSTR         szSource,
                  /* [in] */                    LPWSTR         szConditionName,
                  /* [in] */                    DWORD          dwNumEventAttrs,
                  /* [size_is][in] */           DWORD       *  pdwAttributeIDs,
                  /* [out] */                   OPCCONDITIONSTATE
                                                            ** ppConditionState
                  );

   STDMETHODIMP EnableConditionByArea(
                  /* [in] */                    DWORD          dwNumAreas,
                  /* [size_is][in] */           LPWSTR      *  pszAreas
                  );

   STDMETHODIMP EnableConditionBySource(
                  /* [in] */                    DWORD          dwNumSources,
                  /* [size_is][in] */           LPWSTR      *  pszSources
                  );

   STDMETHODIMP DisableConditionByArea(
                  /* [in] */                    DWORD          dwNumAreas,
                  /* [size_is][in] */           LPWSTR      *  pszAreas
                  );

   STDMETHODIMP DisableConditionBySource(
                  /* [in] */                    DWORD          dwNumSources,
                  /* [size_is][in] */           LPWSTR      *  pszSources
                  );

   STDMETHODIMP AckCondition(
                  /* [in] */                    DWORD          dwCount,
                  /* [string][in] */            LPWSTR         szAcknowledgerID,
                  /* [string][in] */            LPWSTR         szComment,
                  /* [size_is][in] */           LPWSTR      *  pszSource,
                  /* [size_is][in] */           LPWSTR      *  pszConditionName,
                  /* [size_is][in] */           FILETIME    *  pftActiveTime,
                  /* [size_is][in] */           DWORD       *  pdwCookie,
                  /* [size_is][size_is][out] */ HRESULT     ** ppErrors
                  );

   STDMETHODIMP CreateAreaBrowser(
                  /* [in] */                    REFIID         riid,
                  /* [iid_is][out] */           LPUNKNOWN   *  ppUnk
                  );

   ///////////////////////////////////////////////////////////////////////////

      // Handles new events for all subscriptions of this server.
   HRESULT  ProcessEvents( DWORD dwNumOfEvents, AeEvent** ppEvents );

      // Pure virtual function, must be implemented by derived classes
   virtual void FireShutdownRequest( LPCWSTR szReason ) = 0;

// Impmementation
protected:
      // The time this server instance was started.
   FILETIME m_ftStartTime;

      // The last time the server sent an event notification to the client.
   FILETIME m_ftLastUpdateTime;

private:
      // Adds an event subscription instance to the subscription list.
      // This method is called when a client creates a new subscription.
   HRESULT  AddSubscriptionToList( AeComSubscriptionManager* pSubscr );

      // Removes an event subscription instance from the subscription list.
      // This method is called when a client has released all interfaces
      // of the event subscription object.
   HRESULT  RemoveSubscriptionFromList( AeComSubscriptionManager* pSubscr );

      // The critical section to lock/unlock the list 
      // of Event Subscription instances.
   CComAutoCriticalSection m_csSubscrList;

      // The list of the subscriptions events created by the client.
   CSimpleArray<AeComSubscriptionManager*>   m_listSubscr;

   AeBaseServer*  m_pServerHandler;
};
//DOM-IGNORE-END

#endif // __EVENTSERVER_H
