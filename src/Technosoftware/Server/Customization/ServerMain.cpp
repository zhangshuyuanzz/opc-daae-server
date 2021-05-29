/*
 * Copyright (c) 2011-2021 Technosoftware GmbH. All rights reserved
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

//-----------------------------------------------------------------------------
// INLCUDE
//-----------------------------------------------------------------------------
#include "stdafx.h"                             // Generic server part headers
#include <windows.h>
#include <stdio.h>
#include "CoreMain.h"
// Application specific definitions
#ifdef NDEBUG
#include <stdio.h>
#endif
#include <Crtdbg.h>
#include "DaServer.h"
#ifdef _OPC_SRV_AE
#include "AeServer.h"
#endif
#include "OpcTextReader.h"
#include "IClassicBaseNodeManager.h"
#include "Logger.h"

using namespace IClassicBaseNodeManager;

//-----------------------------------------------------------------------------
// DEFINE
//-----------------------------------------------------------------------------
#define MSGBOX_TITLE          _T("OPC DA/AE Server SDK C++")

//=============================================================================
// The one and only DaServer/AeServer objects
//=============================================================================

// Global pointer to the Application Specific OPC Data Access Handler.
DaServer* gpDataServer = NULL;

#ifdef   _OPC_SRV_AE                            // Alarms & Events Server
// Global pointer to the Application Specific OPC Alarms&Events Handler.
AeServer* gpEventServer = NULL;
#endif

// The fully qualified name of the loaded DLL
static TCHAR gszFullyQualifiedDLLName[ _MAX_PATH ];

//-----------------------------------------------------------------------------
// Global Variables
//-----------------------------------------------------------------------------
DaBrowseMode	gdwBrowseMode = DaBrowseMode::Generic;					// Browse Mode
int         gdwUpdateCyclePeriod = 0;           // Device update rate in ms
bool 		gUseOnItemRequest = true;
bool		gUseOnRefreshItems = true;
bool 		gUseOnAddItem = false;
bool 		gUseOnRemoveItem = false;
LPWSTR	    vendor_name;


HINSTANCE   gDLLHandle = 0;                     // DLL Handle from LoadLibrary
TCHAR*      gpszDLLParams = NULL;               // Parameter string for the Application-DLL
BOOL        RegisterFlags = FALSE;

//-----------------------------------------------------------------------------
// LOCAL
//-----------------------------------------------------------------------------

//=============================================================================
// SERVER REGISTRY DEFINITIONS
//=============================================================================

// The following setting are overwritten by the assemply plugin defined values
static SERVERREGDEFS RegistryDefinitions =
{
	//    Application
	L"{8817FE7E-0C33-4ebb-BF7E-6A9C0C17FA02}",                  // Application identifier
	L"Technosoftware DA Sample Server V1.0",                        // Application description

	//    Data Access Server
	L"{1B8200C7-DFA0-473c-AED8-315037F3AAD7}",                  // CLSID of Data Access Server
	L"Technosoftware.DataSample",					            // Version independent ProgID
	L"Technosoftware.DataSample.10",								// ProgID of current server version
#ifdef _DEBUG
#define DESCR_POSTFIX L" (Debug)"
#else
#define DESCR_POSTFIX L""
#endif
	L"Technosoftware DA Sample Server" DESCR_POSTFIX,       // Friendly name of version independent server
	L"Technosoftware DA Sample Server V1.0" DESCR_POSTFIX,  // Friendly name of current server version

	FALSE,														// Don't use AE Server
	// Alarms & Events Server
	L"",                                                        // CLSID of Alarms & Events Server
	L"",     // This Sample Server provides                     // Version independent ProgID
	L"",     // no Alarms & Events functionality                // ProgID of current server version
	L"",                                                        // Friendly name of version independent server
	L""                                                         // Friendly name of current server version
};

//-----------------------------------------------------------------------------
// CODE
//-----------------------------------------------------------------------------

//=============================================================================
// IMPLEMENTATION INTERNAL FUNCTIONS
//=============================================================================

void LogEnable(const char* loggerDefaultPath)
{
	COpcString cAppDLLName;
	COpcString cFlags;
	BOOL     fErrMsgDisplayed = FALSE;

	int logLevel = OnGetLogLevel();
	const char * logPath = "";
	if (logLevel != LOG_LEVEL_DISABLE)
	{
		OnGetLogPath(logPath);

		if (logPath == NULL || strlen(logPath) == 0)
		{
			logPath = loggerDefaultPath;
		}

		// configure the output behaviour
		LogManager::getRef().setLoggerPath(LOGGER_MAIN_LOGGER_ID, logPath);
		LogManager::getRef().setLoggerLevel(LOGGER_MAIN_LOGGER_ID, logLevel);
		LogManager::getRef().setLoggerDisplay(LOGGER_MAIN_LOGGER_ID, false);
		LogManager::getRef().setLoggerFileLine(LOGGER_MAIN_LOGGER_ID, false);

		// start logger
		LogManager::getRef().start();

	}
	else
	{
		// stop logger
		LogManager::getRef().stop();
	}
}

void LogEnable(const string loggerDefaultPath, const string arguments)
{
	LogEnable(loggerDefaultPath.c_str());
}

//=============================================================================
// AddParamToRegistry
// ------------------
//    Adds a parameter to the Registry key
//    HKEY_CLASSES_ROOT\CLSID\{<clsid>}\LocalServer32
//=============================================================================
HRESULT AddParamToRegistry( LPCTSTR pszParam )
{
	// Get Registry Defs from DLL
	ClassicServerDefinition* pDefs = OnGetDaServerDefinition();

	_ASSERTE( pDefs != NULL );
	if (!pDefs)
		return E_FAIL;

	USES_CONVERSION;
	TCHAR szKey[70] = _T("CLSID\\");
	_tcscat_s( szKey, W2T(pDefs->ClsidServer) );
	_tcscat_s( szKey, _T("\\LocalServer32") );

	CRegKey key;
	LONG lRes = key.Open( HKEY_CLASSES_ROOT, szKey );
	if (lRes != ERROR_SUCCESS)
		return HRESULT_FROM_WIN32( lRes );

	TCHAR szLocalServer[512];
	DWORD dwc = sizeof (szLocalServer);
	lRes = key.QueryStringValue( _T(""), szLocalServer, &dwc );

	if (lRes != ERROR_SUCCESS)
		return HRESULT_FROM_WIN32( lRes );

	_tcscat_s( szLocalServer, _T(" ") );
	_tcscat_s( szLocalServer, pszParam );
	lRes = key.SetStringValue( _T(""), szLocalServer );
	return HRESULT_FROM_WIN32( lRes );
}


//=============================================================================
// IMPLEMENTATION GLOBAL INTERFACE FUNCTIONS
//=============================================================================


HRESULT OnProcessParam( COpcString flags ) 
{
	HRESULT  hres = S_OK;

	if (flags == _T("embedding") || flags == _T("automation"))
	{
		// Ignore the OLE switch parameters 'Embedding' and 'Automation'
		hres = S_OK;
	}
	else
	{
		if (RegisterFlags) {												// Update the Registry
			hres = AddParamToRegistry(flags);
		}
	}
	return hres;
}

//=============================================================================
// OnGetServerRegistryDefs
// -----------------------
//    Returns the address of the structure which specifies all
//    definitions required to register the server in the Registry.
//
//    This function also loads the application DLL.
//
//    Used Application DLL and command-line parameters :
//       Uses the name of the DLL from the command-line if specified.
//       Sets the global flag 'RegisterFlags' if the command-line
//       option /RegServer  or /Install is specified.
//
//    Also the global members gDLLHandle and gszFullyQualifiedDLLName
//    are initialized.
//
//    Note :   Terminate the application and do not return from this
//             function if something goes wrong.
//=============================================================================
LPSERVERREGDEFS OnGetServerRegistryDefs()
{
	HRESULT  hres = S_OK;
	BOOL     fErrMsgDisplayed = FALSE;
	BOOL     fAbort = FALSE;
	COpcString cFlags;

	if (SUCCEEDED( hres )) {
		OnStartupSignal( GetCommandLine() );

		// parse command line arguments.
		COpcText cText;
		COpcTextReader cReader(GetCommandLine());

		// skip until first command flag - if any.
		cText.SetType(COpcText::Delimited);
		cText.SetDelims(L"-/");

		if (cReader.GetNext(cText))
		{
			// read first command flag.
			cText.SetType(COpcText::NonWhitespace);
			cText.SetEofDelim();

			while (cReader.GetNext(cText))
			{
				COpcString cFlags = ((COpcString&)cText).ToLower();

	            // register module as local server.
		        if (cFlags== _T("regserver")) RegisterFlags = TRUE;

			}
		}
	}

	// Error handling
	if (FAILED( hres )) {
		fAbort = TRUE;
		if (!fErrMsgDisplayed) {                  // Display error message if not yet done
			TCHAR buf[512] = _T("\0");

			LPTSTR   pDescr = NULL;
			// Get the description of the error from the system.
			if (FormatMessage(
				FORMAT_MESSAGE_ALLOCATE_BUFFER |
				FORMAT_MESSAGE_FROM_SYSTEM |
				FORMAT_MESSAGE_IGNORE_INSERTS,
				NULL,                         // No message source
				hres,                         // Error Code
				// Language Identifier
				LANGIDFROMLCID( LOCALE_SYSTEM_DEFAULT ),
				(LPTSTR)&pDescr,              // Output string buffer
				0,                            // no static message buffer
				NULL )) {                     // no message insertion.

					// Description found
					_tcscat_s( buf, pDescr );           // Add description to the message to display
					LocalFree( pDescr );
			}
			MessageBox( NULL, buf, MSGBOX_TITLE, MB_OK | MB_ICONSTOP ); 
		}
	}

	if (fAbort) {
		exit(hres);								// Terminate the application if something goes wrong
	}

	WCHAR    wDelimiter;

	// Get Server Parameters from DLL
	hres = OnGetDaServerParameters(&gdwUpdateCyclePeriod, &wDelimiter, &gdwBrowseMode);

	// Error handling
	if (FAILED( hres )) {
		LOGFMTE("OnGetDAServerParameters() failed with hres = 0x%x.", hres);
		exit(hres);								// Terminate the application if something goes wrong
	}

	LOGFMTT("OnGetDAServerParameters() finished with hres = 0x%x.", hres);

	// Get Server Optimization Parameters from DLL
	hres = OnGetDaOptimizationParameters( &gUseOnItemRequest, &gUseOnRefreshItems, &gUseOnAddItem, &gUseOnRemoveItem );

	// Error handling
	if (FAILED( hres )) {
		LOGFMTE("OnGetDAOptimizationParameters() failed with hres = 0x%x.", hres);
		exit(hres);								// Terminate the application if something goes wrong
	}

	LOGFMTT("OnGetDAOptimizationParameters() finished with hres = 0x%x.", hres);

	DaBranch::SetDelimiter(wDelimiter);
	// Get DA Server Registry Defs from DLL
	ClassicServerDefinition* pRegDefsDLL = OnGetDaServerDefinition();
	if (pRegDefsDLL == NULL) {
		LOGFMTE("OnGetDaServerDefinition() failed with returning NULL.");
		return NULL;
	}                                            // Initialize the new Registry Definition Structure
	// Application
	RegistryDefinitions.m_szAppID            = pRegDefsDLL->ClsidApp;
	RegistryDefinitions.m_szAppDescr         = pRegDefsDLL->NameCurrServer;

	// Data Access Server
	RegistryDefinitions.m_szDA_CLSID         = pRegDefsDLL->ClsidServer;
	RegistryDefinitions.m_szDA_IndepProgID   = pRegDefsDLL->PrgidServer;
	RegistryDefinitions.m_szDA_CurrProgID    = pRegDefsDLL->PrgidCurrServer;
	RegistryDefinitions.m_szDA_IndepDescr    = pRegDefsDLL->NameServer;
	RegistryDefinitions.m_szDA_CurrDescr     = pRegDefsDLL->NameCurrServer;

	RegistryDefinitions.m_fUseAE_Server		 = FALSE;

	LOGFMTI(  "OnGetDaServerDefinition() returned DA server definitions:" );
	OPCWSTOAS( RegistryDefinitions.m_szDA_CLSID );
		LOGFMTI( "   CLSID       '%s'", OPCastr );
    }
    OPCWSTOAS( RegistryDefinitions.m_szDA_IndepProgID );  
		LOGFMTI( "   IndepProgID '%s'", OPCastr );
    }
	OPCWSTOAS( RegistryDefinitions.m_szDA_CurrProgID );
		LOGFMTI( "   CurrProgID  '%s'", OPCastr );
    }
	OPCWSTOAS( RegistryDefinitions.m_szDA_IndepDescr );
		LOGFMTI( "   IndepDescr  '%s'", OPCastr );
    }
	OPCWSTOAS( RegistryDefinitions.m_szDA_CurrDescr );
		LOGFMTI( "   CurrDescr   '%s'", OPCastr );
    }

	vendor_name = pRegDefsDLL->CompanyName;

	OPCWSTOAS( vendor_name );
	LOGFMTT( "   VendorName  '%s'", OPCastr );
    }

#ifdef _OPC_SRV_AE
	// Get AE Server Registry Defs from DLL

	pRegDefsDLL = OnGetAeServerDefinition();
	if (pRegDefsDLL == NULL || !pRegDefsDLL->ClsidServer) {
		LOGFMTT("OnGetAeServerDefinition() returned NULL.");
		return &RegistryDefinitions;
	}                                            // Initialize the new Registry Definition Structure

	RegistryDefinitions.m_fUseAE_Server		 = TRUE;

	// Alarms&Events Server
	RegistryDefinitions.m_szAE_CLSID         = pRegDefsDLL->ClsidServer;
	RegistryDefinitions.m_szAE_IndepProgID   = pRegDefsDLL->PrgidServer;
	RegistryDefinitions.m_szAE_CurrProgID    = pRegDefsDLL->PrgidCurrServer;
	RegistryDefinitions.m_szAE_IndepDescr    = pRegDefsDLL->NameServer;
	RegistryDefinitions.m_szAE_CurrDescr     = pRegDefsDLL->NameCurrServer;

	LOGFMTI("OnGetAeServerDefinition() returned AE server definitions:");
	OPCWSTOAS(RegistryDefinitions.m_szAE_CLSID);
		LOGFMTI("   CLSID       '%s'", OPCastr);
	}
	OPCWSTOAS(RegistryDefinitions.m_szAE_IndepProgID);
		LOGFMTI("   IndepProgID '%s'", OPCastr);
	}
	OPCWSTOAS(RegistryDefinitions.m_szAE_CurrProgID);
		LOGFMTI("   CurrProgID  '%s'", OPCastr);
	}
	OPCWSTOAS(RegistryDefinitions.m_szAE_IndepDescr);
		LOGFMTI("   IndepDescr  '%s'", OPCastr);
	}
	OPCWSTOAS(RegistryDefinitions.m_szAE_CurrDescr);
		LOGFMTI("   CurrDescr   '%s'", OPCastr);
	}

#endif
	return &RegistryDefinitions;
}



//=============================================================================
// OnInitializeServer
// ------------------
//    This function is called when the server EXE is loaded in memory
//    and creates an instance of the Data Access Server Handler instance.
//    This function is called from the generic part function
//    CoreGenericMain::InitializeServer() in module CoreMain.cpp.
//=============================================================================
HRESULT OnInitializeServer() 
{
	//_CrtSetDbgFlag( _CRTDBG_LEAK_CHECK_DF | _CrtSetDbgFlag( _CRTDBG_REPORT_FLAG ) );
	// Init the global Data Access Server Instance
	gpDataServer = new DaServer();
	if (gpDataServer == NULL) {
		return E_OUTOFMEMORY;
	}
#ifdef   _OPC_SRV_AE                            // Alarms & Events Server
	// Init the global Alarms&Events Server Instance
	if (RegistryDefinitions.m_fUseAE_Server == TRUE)
	{
		gpEventServer = new AeServer();
		if (gpEventServer == NULL) {
			return E_OUTOFMEMORY;
		}
	}
#endif
	HRESULT hres;
#ifdef   _OPC_SRV_AE                            // Alarms & Events Server
	if (gpEventServer != NULL)
	{
		hres = gpEventServer->Create();
	}
#endif
	hres = gpDataServer->Create(
		gdwUpdateCyclePeriod,					// Device Update rate in ms
		gpszDLLParams,							// Parameter string for the Application-DLL
		MSGBOX_TITLE );							// Server Instance Name

	if (FAILED( hres )) {
#ifdef   _OPC_SRV_AE                            // Alarms & Events Server
		if (gpEventServer != NULL)
		{
			delete gpEventServer;
			gpEventServer = NULL;
		}
#endif
		delete gpDataServer;
		gpDataServer = NULL;
	}
	return hres;
}



//=============================================================================
// OnTerminateServer
// -----------------
//    This function is called when the server EXE in unloaded from memory.
//    Destroys all what the OnInitializeServer() function has created.
//    This function also frees the application DLL.
//=============================================================================
HRESULT OnTerminateServer() 
{
	HRESULT hres = S_OK;

	OnShutdownSignal();

#ifdef   _OPC_SRV_AE                            // Alarms & Events Server
	if (gpEventServer) {
		delete gpEventServer;
		gpEventServer = NULL;
	}
#endif
	if (gpDataServer) {
		delete gpDataServer;
		gpDataServer = NULL;
	}

	if (gDLLHandle) {
		FreeLibrary( gDLLHandle );
		gDLLHandle = 0;
	}
	if (gpszDLLParams) {
		free( gpszDLLParams );
		gpszDLLParams = NULL;
	}
	return hres;
}



//=============================================================================
// OnGetDaBaseServer
// ------------------
//    Returns a pointer to the Data Access Server Handler instance for
//    the given Server Class ID.
// Parameters:
//    IN
//       szServerClassID         OPC:     It's always a NULL-Pointer
//                                        because OPC supports only
//                                        one COM-Server per EXE. 
//                               CALL-R:  Text string with the identifier
//                                        of the requested Data Access
//                                        Server Handler. The identifier
//                                        can be defined with function
//                                        DaBaseServer::Create(). 
//=============================================================================
DaBaseServer* OnGetDaBaseServer( LPCWSTR szServerClassID )
{
	return gpDataServer;
}

#ifdef   _OPC_SRV_AE                            // Alarms & Events Server

//=========================================================================
// OnGetAeBaseServer
// -----------------------
//    Returns a pointer to the Event Server Handler instance.
//=========================================================================
AeBaseServer* OnGetAeBaseServer()
{
	return gpEventServer;
}
#endif

