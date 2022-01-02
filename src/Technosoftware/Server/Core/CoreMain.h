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

#ifndef __COREMAIN_H_
#define __COREMAIN_H_

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include <string>

#ifdef   _OPC_SRV_DA                            // Data Access Server
#include "DaComBaseServer.h"
#include "DaBaseServer.h"
#endif
#ifdef   _OPC_SRV_AE                            // Alarms & Events Server
#include "AeBaseServer.h"
#endif
#include "OpcTextReader.h"

using std::string;

/**
 * @fn	HRESULT OnInitializeServer();
 *
 * @brief	Called when server is started (that is the first client connects)
 * 			Mainly should create instances of the DaBaseServer/AeBaseServer derivations but could
 * 			also read from registry, etc. Attention! this function is not called when the server
 * 			is started with parameters (like when it is installed)!
 *
 * @return	A hResult.
 */

HRESULT OnInitializeServer();

/**
 * @fn	HRESULT OnTerminateServer();
 *
 * @brief	Called when server is unloading (shutting down)
 * 			It's the counterpart of OnInitializeServer.
 *
 * @return	A hResult.
 */

HRESULT OnTerminateServer();

/**
 * @fn	HRESULT OnProcessParam( COpcString flags );
 *
 * @brief	Custom parameter processing when server started from command line or by COM (with
 * 			switches /Embedding or /Automation). The generic server part handles the
 * 			following arguments:
 * 				/Install 
 * 				/Deinstall 
 * 				/RegServer 
 * 				/UnregServer
 * 			
 *			For all other argument strings this function is called in the order they are in the command line. 
 *			Return E_FAIL to signal	that the server must not be run.
 * 			
 * 			OPC DA/AE Server SDK .NET and OPC DA/AE Server SDK DLL Server specific parameter handling: 
 * 			All Server specific parameters are added to the appropriate Registry entry if the global flag 'RegisterFlags' is set
 * 			(the fully qualified name of the DLL is used). A copy of the DLL-Driver specific
 * 			parameter string is stored in the global variable 'gpszDLLParams'.
 *
 * @param	flags	The flags.
 *
 * @return	A hResult.
 */

HRESULT OnProcessParam( COpcString flags );

#ifdef   _OPC_SRV_DA                            // Data Access Server

/**
 * @fn	DaBaseServer* OnGetDaBaseServer( LPCWSTR serverClassId );
 *
 * @brief	This function is called from the constructor of a COM server class (ex. DaComServer)
 * 			to get the instance of it's server class handler The interpretation of the parameter
 * 			has only sense when more COM Server classes are located in the same EXE, in that case
 * 			a different ID allows to attach to different handler instances (or classes)
 *
 * @param	serverClassId	Identifier for the server class.
 *
 * @return	null if it fails, else a DaBaseServer*.
 */

DaBaseServer* OnGetDaBaseServer( LPCWSTR serverClassId );
#endif

#ifdef   _OPC_SRV_AE                            // Alarms & Events Server
   // This function is called from the constructor of
   // a COM server class (ex. AeComServer) to get the instance
   // of it's event server class handler. A Server EXE has only one
   // event server class handler.
AeBaseServer* OnGetAeBaseServer();
#endif // _OPC_SRV_AE

// 
// Structure with Server Registry Definitions.
// 
typedef struct tagSERVERREGDEFS
{
      // Application
   LPOLESTR m_szAppID;
   LPOLESTR m_szAppDescr;

      // Data Access Server
   LPOLESTR m_szDA_CLSID;
   LPOLESTR m_szDA_IndepProgID;
   LPOLESTR m_szDA_CurrProgID;
   LPOLESTR m_szDA_IndepDescr;
   LPOLESTR m_szDA_CurrDescr;

   BOOL		m_fUseAE_Server;
      // Alarms & Events Server
   LPOLESTR m_szAE_CLSID;
   LPOLESTR m_szAE_IndepProgID;
   LPOLESTR m_szAE_CurrProgID;
   LPOLESTR m_szAE_IndepDescr;
   LPOLESTR m_szAE_CurrDescr;

} SERVERREGDEFS, *LPSERVERREGDEFS;

   // This function is called during installation
   // and returns server specific entries for the Registry.
LPSERVERREGDEFS OnGetServerRegistryDefs();
void LogEnable(const string loggerDefaultPath, const string arguments);

//DOM-IGNORE-END

#endif // __COREMAIN_H_

