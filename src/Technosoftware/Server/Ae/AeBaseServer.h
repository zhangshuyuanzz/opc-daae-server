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

#ifndef __EVENTSERVERCLASSHANDLER_H
#define __EVENTSERVERCLASSHANDLER_H


#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "AeEventArea.h"
#include "AeComBaseServer.h"
#include "AeCondition.h"


//-----------------------------------------------------------------------
// CLASS AeBaseServer
//-----------------------------------------------------------------------
class AeBaseServer
{
// Friends
         // This friend declaration allows the class AeComBaseServer
         // to add itself to the list of this handler using
         // AddServerToList() and RemoveServerFromList()
         // which is private and should not be accessed by the specific
         // part of the server.
   friend AeComBaseServer;
         // This friend class calls ExistEventAttribute(),
         // ExistEventCategory(), ExistSource() and ExistArea().
   friend AeComSubscriptionManager;
         // This friend function calls AckNotification()
   friend HRESULT AeCondition::Acknowledge( LPCWSTR szAcknowledgerID, LPCWSTR szComment,
                                      FILETIME ftActiveTime, AeEvent** ppEvent, BOOL useCurrentTime /* = TRUE */);
         // This friend function calls GetEventsForRefresh()
   friend unsigned __stdcall EventRefreshThreadHandler( void* pCreator );

// Construction
public:
   AeBaseServer();
   HRESULT  Create();

// Destruction
   ~AeBaseServer();

// Operations
public:

   ////////////////////////////////////////////////
   // Functions for initialization of the AE server
   ////////////////////////////////////////////////

   //
   // Functions to build the Event Space
   //
   HRESULT  AddSimpleEventCategory( DWORD dwCatID, LPCWSTR szDescr );
   HRESULT  AddTrackingEventCategory( DWORD dwCatID, LPCWSTR szDescr );
   HRESULT  AddConditionEventCategory( DWORD dwCatID, LPCWSTR szDescr );

   HRESULT  AddEventAttribute( DWORD dwCatID, DWORD dwAttrID, LPCWSTR szDescr, VARTYPE vt );

   HRESULT  AddSingleStateConditionDef( DWORD dwCatID, DWORD dwCondDefID,
                                       LPCWSTR szName, LPCWSTR szDef, DWORD dwSeverity, LPCWSTR szDescr,
                                       BOOL fAckRequired );


   HRESULT  AddMultiStateConditionDef( DWORD dwCatID, DWORD dwCondDefID, LPCWSTR szName );
   HRESULT  AddSubConditionDef( DWORD dwCondDefID, DWORD dwSubCondDefID, LPCWSTR szName, LPCWSTR szDef,
                              DWORD dwSeverity, LPCWSTR szDescr, BOOL fAckRequired );


   //
   // Functions to build the Process Area Space
   //
   HRESULT  AddArea( DWORD dwParentAreaID, DWORD dwAreaID, LPCWSTR szName );
   HRESULT  AddSource( DWORD dwAreaID, DWORD dwSrcID, LPCWSTR szName, BOOL fMultiSource = FALSE );
   HRESULT  AddExistingSource( DWORD dwAreaID, DWORD dwSrcID );
   HRESULT  AddCondition( DWORD dwSrcID, DWORD dwCondDefID, DWORD dwCondID );


   ////////////////////////////////////////////////////
   // Functions for event handling
   ////////////////////////////////////////////////////

   HRESULT  ProcessSimpleEvent( DWORD dwCatID, DWORD dwSrcID, LPCWSTR szMessage, DWORD dwSeverity, DWORD dwAttrCount = 0, LPVARIANT pvAttrValues = NULL, LPFILETIME pft = NULL );
   HRESULT  ProcessTrackingEvent( DWORD dwCatID, DWORD dwSrcID, LPCWSTR szMessage, DWORD dwSeverity, LPCWSTR szActorID, DWORD dwAttrCount = 0, LPVARIANT pvAttrValues = NULL, LPFILETIME pft = NULL  );
   HRESULT  ProcessConditionStateChanges( DWORD dwCount, AeConditionChangeStates* pCondStateChanges, BOOL useCurrentTime = TRUE, HRESULT* errors = NULL  );
   HRESULT  AckCondition( DWORD dwCondID, LPCWSTR szComment = NULL );

      // Fires a 'Shutdown Request' to all subscribed clients
   void FireShutdownRequest( LPCWSTR szReason = NULL );

// Implementation
protected:
   ////////////////////////////////////////////////////
   // Pure virtual functions which must be
   // implemented by the derived class.
   ////////////////////////////////////////////////////

   virtual HRESULT OnGetServerState(	/*[out]*/   OPCEVENTSERVERSTATE & serverState,
										/*[out]*/   LPWSTR				& vendor ) = 0;

   virtual HRESULT OnTranslateToItemIdentifier(
                                 DWORD dwCondID, DWORD dwSubCondDefID, DWORD dwAttrID,
                                 LPWSTR* pszItemID, LPWSTR* pszNodeName, CLSID* pCLSID ) = 0;

   virtual HRESULT OnAcknowledgeNotification(DWORD dwCondID, DWORD dwSubConDefID) = 0;

   virtual HRESULT OnConnectClient(  ) = 0;

   virtual HRESULT OnDisconnectClient(  ) = 0;


   ////////////////////////////////////////////////////
   // Functions called from AeComBaseServer
   ////////////////////////////////////////////////////

   HRESULT  QueryEventCategories(
   /* [in] */                    DWORD          dwEventType,
   /* [out] */                   DWORD       *  pdwCount,
   /* [size_is][size_is][out] */ DWORD       ** ppdwEventCategories,
   /* [size_is][size_is][out] */ LPWSTR      ** ppszEventCategoryDescs );

   HRESULT  QueryConditionNames(
   /* [in] */                    DWORD          dwEventCategory,
   /* [out] */                   DWORD       *  pdwCount,
   /* [size_is][size_is][out] */ LPWSTR      ** ppszConditionNames );

   HRESULT  QuerySubConditionNames(
   /* [in] */                    LPWSTR         szConditionName,
   /* [out] */                   DWORD       *  pdwCount,
   /* [size_is][size_is][out] */ LPWSTR      ** ppszSubConditionNames );

   HRESULT  QueryEventAttributes(
   /* [in] */                    DWORD          dwEventCategory,
   /* [out] */                   DWORD       *  pdwCount,
   /* [size_is][size_is][out] */ DWORD       ** ppdwAttrIDs,
   /* [size_is][size_is][out] */ LPWSTR      ** ppszAttrDescs,
   /* [size_is][size_is][out] */ VARTYPE     ** ppvtAttrTypes );

   HRESULT  QuerySourceConditions(
   /* [in] */                    LPWSTR         szSource,
   /* [out] */                   DWORD       *  pdwCount,
   /* [size_is][size_is][out] */ LPWSTR      ** ppszConditionNames );

   HRESULT  GetConditionState(
   /* [in] */                    LPWSTR         szSource,
   /* [in] */                    LPWSTR         szConditionName,
   /* [in] */                    DWORD          dwNumEventAttrs,
   /* [size_is][in] */           DWORD       *  pdwAttributeIDs,
   /* [out] */                   OPCCONDITIONSTATE
                                             ** ppConditionState );

   HRESULT  AckCondition(
   /* [in] */                    DWORD          dwCount,
   /* [string][in] */            LPWSTR         szAcknowledgerID,
   /* [string][in] */            LPWSTR         szComment,
   /* [size_is][in] */           LPWSTR      *  pszSource,
   /* [size_is][in] */           LPWSTR      *  pszConditionName,
   /* [size_is][in] */           FILETIME    *  pftActiveTime,
   /* [size_is][in] */           DWORD       *  pdwCookie,
   /* [size_is][size_is][out] */ HRESULT     ** ppErrors );

   HRESULT  EnableConditions(
   /* [in] */                    BOOL           fEnable,
   /* [in] */                    BOOL           fBySource,
   /* [in] */                    DWORD          dwCount,
   /* [size_is][in] */           LPWSTR      *  pszNames );

   HRESULT  TranslateToItemIDs(
   /* [in] */                    LPWSTR         szSource,
   /* [in] */                    DWORD          dwEventCategory,
   /* [in] */                    LPWSTR         szConditionName,
   /* [in] */                    LPWSTR         szSubconditionName,
   /* [in] */                    DWORD          dwCount,
   /* [size_is][in] */           DWORD       *  pdwAssocAttrIDs,
   /* [size_is][size_is][out] */ LPWSTR      ** ppszAttrItemIDs,
   /* [size_is][size_is][out] */ LPWSTR      ** ppszNodeNames,
   /* [size_is][size_is][out] */ CLSID       ** ppCLSIDs );


   ////////////////////////////////////////////////////
   // Functions called from EventRefreshThreadHandler()
   ////////////////////////////////////////////////////

   HRESULT  GetEventsForRefresh(
   /* [in] */                    AeComSubscriptionManager* pSubscr,
   /* [out] */                   CSimpleArray<AeEvent*>& arEventPtrs );


   ////////////////////////////////////////////////////
   // Functions called from AeComSubscriptionManager
   ////////////////////////////////////////////////////

   BOOL     ExistEventAttribute( DWORD dwEventCategory, DWORD dwAttrID );
   BOOL     ExistEventCategory( DWORD dwEventCategory );
   BOOL     ExistSource( LPCWSTR szName );
   BOOL     ExistArea( LPCWSTR szName );


   ////////////////////////////////////////////////////
   // Internal function and data members
   ////////////////////////////////////////////////////

   HRESULT FireEvent( AeEvent* pEvent );
   HRESULT FireEvents( AeEventArray& arEvents );

   HRESULT AddEventCategory( DWORD dwCatID, LPCWSTR szDescr, DWORD dwEventType );
   HRESULT ProcessEvent( DWORD dwCatID, DWORD dwSrcID, LPCWSTR szMessage, DWORD dwSeverity, LPCWSTR szActorID, DWORD dwAttrCount, LPVARIANT pvAttrValues, LPFILETIME pft );

   CSimpleMap<DWORD, AeCategory*>     m_mapCategories;
   CSimpleMap<DWORD, AeSource*>       m_mapSources;
   CSimpleMap<DWORD, AeConditionDefiniton*> m_mapConditionDefs;
   CSimpleMap<DWORD, AeCondition*>    m_mapConditions;

   AeSource* LookupSource( LPCWSTR szName );
   AeConditionDefiniton* LookupConditionDef( LPCWSTR szName );

   CComAutoCriticalSection m_csCatMap;          // lock/unlock m_mapCategories
   CComAutoCriticalSection m_csSrcMap;          // lock/unlock m_mapSources
   CComAutoCriticalSection m_csCondDefMap;      // lock/unlock m_mapConditionDefs
   CComAutoCriticalSection m_csCondMap;         // lock/unlock m_mapConditions

   EventArea              m_RootArea;

   ////////////////////////////////////////////////////
   // List of the connected clients
   ////////////////////////////////////////////////////
private:
      // Adds an event server instance to the event server list.
      // This method is called when a client connects.
   HRESULT AddServerToList( AeComBaseServer* pServer );

      // Removes a server instance from the event server list.
      // This method is called when a client has properly disconnected,
      // that is freed all the COM interfaces it owns.
   HRESULT RemoveServerFromList( AeComBaseServer* pServer );

      // The critical section to lock/unlock the list 
      // of event server instances.
   CComAutoCriticalSection m_csServers;

      // The list of the attached Event Server instances.
      // There is one list entry for each connected client.
   CSimpleArray<AeComBaseServer*>   m_arServers;
};


//-----------------------------------------------------------------------
// INLINEs of CLASS AeBaseServer
//-----------------------------------------------------------------------

inline HRESULT AeBaseServer::AddSimpleEventCategory( DWORD dwCatID, LPCWSTR szDescr )
                     { return AddEventCategory( dwCatID, szDescr, OPC_SIMPLE_EVENT ); }
inline HRESULT AeBaseServer::AddTrackingEventCategory( DWORD dwCatID, LPCWSTR szDescr )
                     { return AddEventCategory( dwCatID, szDescr, OPC_TRACKING_EVENT ); }
inline HRESULT AeBaseServer::AddConditionEventCategory( DWORD dwCatID, LPCWSTR szDescr )
                     { return AddEventCategory( dwCatID, szDescr, OPC_CONDITION_EVENT ); }

inline HRESULT AeBaseServer::ProcessSimpleEvent(
                     DWORD dwCatID, DWORD dwSrcID, LPCWSTR szMessage, DWORD dwSeverity,
                     DWORD dwAttrCount /*= 0 */, LPVARIANT pvAttrValues /*= NULL*/, LPFILETIME pft /*=NULL */ )
{
   return ProcessEvent( dwCatID, dwSrcID, szMessage, dwSeverity, NULL,
                        dwAttrCount, pvAttrValues, pft );
}

inline HRESULT AeBaseServer::ProcessTrackingEvent(
                     DWORD dwCatID, DWORD dwSrcID, LPCWSTR szMessage, DWORD dwSeverity, LPCWSTR szActorID,
                     DWORD dwAttrCount /*= 0 */, LPVARIANT pvAttrValues /*= NULL*/, LPFILETIME pft /*=NULL */ )
{
   _ASSERTE( szActorID != NULL );               // Must not be NULL
   return ProcessEvent( dwCatID, dwSrcID, szMessage, dwSeverity, szActorID,
                        dwAttrCount, pvAttrValues, pft );
}



//-----------------------------------------------------------------------
// TYPEDEF CONDCHANGESTATES and CLASS AeBaseServer
//-----------------------------------------------------------------------

typedef struct tagCONDCHANGESTATES {

   DWORD       m_dwCondID;
   DWORD       m_dwSubCondID;
   BOOL        m_fActiveState;
   WORD        m_wQuality;
   DWORD       m_dwAttrCount;
   LPVARIANT   m_pvAttrValues;
   LPCWSTR     m_szMessage;
   DWORD*      m_pdwSeverity;
   LPBOOL      m_pfAckRequired;
   LPFILETIME  m_pftTimeStamp;

} CONDCHANGESTATES;



class AeConditionChangeStates : protected tagCONDCHANGESTATES
{
// Construction
public:
   AeConditionChangeStates();                         // Default Constructor
                                                // Copy Constructor
   AeConditionChangeStates( const AeConditionChangeStates& cs );
   AeConditionChangeStates(                           // Construtor
            // mandatory
         DWORD       dwCondID,
         DWORD       dwSubCondID,
         BOOL        fActiveState,
         WORD        wQuality,
            // optional
         DWORD       dwAttrCount = 0,
         LPVARIANT   pvAttrValues = NULL,
         LPCWSTR     szMessage = NULL,
         DWORD*      pdwSeverity = NULL,
         LPBOOL      pfAckRequired = NULL,
         LPFILETIME  pftTimeStamp = NULL
      );

// Attributes
public:
   inline DWORD      &  CondID()          { return m_dwCondID; }
   inline DWORD      &  SubCondID()       { return m_dwSubCondID; }
   inline BOOL       &  ActiveState()     { return m_fActiveState; }
   inline WORD       &  Quality()         { return m_wQuality; }
   inline DWORD      &  AttrCount()       { return m_dwAttrCount; }
   inline LPVARIANT  &  AttrValuesPtr()   { return m_pvAttrValues; }
   inline LPCWSTR    &  Message()         { return m_szMessage; }
   inline DWORD*     &  SeverityPtr()     { return m_pdwSeverity; }
   inline LPBOOL     &  AckRequiredPtr()  { return m_pfAckRequired; }
   inline LPFILETIME &  TimeStampPtr()    { return m_pftTimeStamp; }
};

#endif // __EVENTSERVERCLASSHANDLER_H

