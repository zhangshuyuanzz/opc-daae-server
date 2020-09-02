/*
 * Copyright (c) 2011-2020 Technosoftware GmbH. All rights reserved
 * Web: https://technosoftware.com 
 * 
 * Purpose: 
 *
 * The Software is subject to the Technosoftware GmbH Source Code License Agreement, 
 * which can be found here:
 * https://technosoftware.com/documents/Source_License_Agreement.pdf
 */
                                                                          
#ifndef __APP_EVENTSERVER_H_
#define __APP_EVENTSERVER_H_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "AeBaseServer.h"

//-------------------------------------------------------------------------
// Event Category IDs
//-------------------------------------------------------------------------
      // Simple Event Category IDs
#define CATID_SYSMESSAGE               0x100
#define CATID_USERMESSAGE              0x110
      // Tracking Event Category IDs   
#define CATID_PCHANGE                  0x200
      // Condition Event Category IDs  
#define CATID_LEVEL                    0x300
                                       
                                       
//-------------------------------------------------------------------------
// Attribute IDs                       
//-------------------------------------------------------------------------
#define ATTRID_LEVEL_CV                0x400
#define ATTRID_PCHANGE_PREVVALUE       0x401
#define ATTRID_PCHANGE_NEWVALUE        0x402
                                       
                                       
//-------------------------------------------------------------------------
// Area IDs
//-------------------------------------------------------------------------
#define AREAID_RAMPS                   0x500
#define AREAID_LIMITS                  0x501
                                       

//-------------------------------------------------------------------------
// Source IDs                          
//-------------------------------------------------------------------------
#define SRCID_SYSTEM                   0x600

                                       
//-------------------------------------------------------------------------
// CLASS
//-------------------------------------------------------------------------
class AeServer : public AeBaseServer
{
// Construction
public:
   AeServer();
   HRESULT Create();

// Destruction
   ~AeServer();

// Implementation
protected:

   // Implementation of the pure virtual functions of the AeBaseServer

   HRESULT OnGetServerState(			/*[out]*/   OPCEVENTSERVERSTATE & serverState,
										/*[out]*/   LPWSTR				& vendor );

   HRESULT OnTranslateToItemIdentifier(
                           DWORD dwCondID, DWORD dwSubCondDefID, DWORD dwAttrID,
                           LPWSTR* pszItemID, LPWSTR* pszNodeName, CLSID* pCLSID );

   HRESULT OnAcknowledgeNotification( DWORD dwCondID, DWORD dwSubConDefID );

   HRESULT OnConnectClient();

   HRESULT OnDisconnectClient();


   // Data Members
   OPCEVENTSERVERSTATE  m_dwServerState;
};

#ifdef   _OPC_SRV_AE                            // Alarms & Events Server
// The Global Alarms&Evenst Server Handler
extern AeServer* gpEventServer;
#endif

#endif // __APP_EVENTSERVER_H_

