/*
 * Copyright (c) 2020 Technosoftware GmbH. All rights reserved
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

 //DOM-IGNORE-BEGIN

 //-----------------------------------------------------------------------
 // INCLUDEs
 //-----------------------------------------------------------------------
#include "stdafx.h"
#include <stdio.h>
#include <comcat.h>

#include "resource.h"
#include "DaComServer.h"
#ifdef _OPC_SRV_AE
#include "AeComServer.h"
#endif
#include "CoreMain.h"
#include "OpcTextReader.h"
#include "CoreGenericMain.h"
#include "Logger.h"

#include <iostream>
#include <string>

//#ifdef _DEBUG
//#define _CTRDBG_MAP_ALLOC
//#include <stdlib.h>
//#include <crtdbg.h>
//#endif

using std::string;

//-----------------------------------------------------------------------
// GLOBALs
//-----------------------------------------------------------------------
// Global memory management object pointer must be created and
// must be released. Defined in CoreGenericMain.h.
IMalloc*    pIMalloc = NULL;
BOOL        fCOMInitialized = FALSE;
CoreGenericMain  _Module;
BOOL		fUseAEServer = FALSE;

// Placeholder for the Server CLSIDs.
CLSID CLSID_OPCServer = { 0x0, 0x0, 0x0, { 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0 } };
CLSID CLSID_OPCEventServer = { 0x0, 0x0, 0x0, { 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0 } };

extern LPWSTR    gVendorName;

#ifdef _NT_SERVICE
#if (_ATL_VER < 0x0700)                         // SetStringValue/QueryStringValue exists since ATL version 7.0
#define CREGKEY_SetStringValue(n,v)       SetValue(v,n)
#define CREGKEY_QueryStringValue(n,v,c)   QueryValue(v,n,c)
#else
#define CREGKEY_SetStringValue(n,v)       SetStringValue(n,v)
#define CREGKEY_QueryStringValue(n,v,c)   QueryStringValue(n,v,c)
#endif
#endif


//-----------------------------------------------------------------------
// INTERNALs
//-----------------------------------------------------------------------
// Define substitution string replacement for
// the Server Registry definitions.

static struct _ATL_REGMAP_ENTRY  pCustSubstDef[] = {
    { L"EXENAME",           NULL },     // Module name without path
    { L"APPID",             NULL },     // Application ID
    { L"APPDESCR",          NULL },     // Application description 

    { L"DA_SERVER_CLSID",   NULL },     // Class ID of Data Access server
    { L"DA_SERVER_ID",      NULL },     // Version independent ProgID
    { L"DA_SERVER_ID_V",    NULL },     // ProgID of current server
    { L"DA_SERVER_NAME",    NULL },     // Friendly name of version independent server
    { L"DA_SERVER_NAME_V",  NULL },     // Friendly name of current server version

    { L"AE_SERVER_CLSID",   NULL },     // Class ID of Alarms & Events server
    { L"AE_SERVER_ID",      NULL },     // Version independent ProgID
    { L"AE_SERVER_ID_V",    NULL },     // ProgID of current server
    { L"AE_SERVER_NAME",    NULL },     // Friendly name of version independent server
    { L"AE_SERVER_NAME_V",  NULL },     // Friendly name of current server version
    { L"VENDOR",            NULL },     // Vendor name
    { NULL,						NULL }
};  // End of table


//-----------------------------------------------------------------------
// MAPs
//-----------------------------------------------------------------------
// Objects to be registered
BEGIN_OBJECT_MAP(ObjectMap)
#ifdef _OPC_SRV_DA          // Data Access Server
    OBJECT_ENTRY(CLSID_OPCServer, DaComServer)
#endif
#ifdef _OPC_SRV_AE          // Alarms & Events Server
    OBJECT_ENTRY(CLSID_OPCEventServer, AeComServer)
#endif
END_OBJECT_MAP()


//-----------------------------------------------------------------------
// FORWARD DECLARATIONs
//-----------------------------------------------------------------------
static LPCTSTR FindOneOf(LPCTSTR p1, LPCTSTR p2);
static HRESULT CreateComCatOPC(void);
static HRESULT UnRegisterServer(const LPSERVERREGDEFS pRegDefs);
static HRESULT RegisterServer(const LPSERVERREGDEFS pRegDefs, BOOL fAsService);
static void CleanupReplacementMapForRegistryUpdate(_ATL_REGMAP_ENTRY aMapEntries[]);
static void SetupReplacementMapForRegistryUpdate(_ATL_REGMAP_ENTRY aMapEntries[],
    const LPSERVERREGDEFS pRegDefs);
#ifdef _NT_SERVICE
static BOOL IsInstalledAsService(LPCTSTR lpszServiceName);
static HRESULT UninstallService(LPCTSTR lpszServiceName);
static HRESULT InstallAsService(LPCTSTR lpszServiceName, LPCTSTR lpszServiceDescription);
static HRESULT IsRegisteredAsService(LPCTSTR lpszAppID, BOOL & fAsService);
static HRESULT AdjustRegistryForServerType(BOOL fRegisterAsService,
    LPCTSTR lpszAppID, LPCTSTR lpszServiceName);
static LPTSTR CreateSafeServiceName(LPCTSTR szNameBase);
#endif // _NT_SERVICE

//-----------------------------------------------------------------------
// INLINE DECLARATIONS
//-----------------------------------------------------------------------

inline bool IsWindowsVersionOrGreater(WORD wMajorVersion, WORD wMinorVersion, WORD wServicePackMajor)
{
    OSVERSIONINFOEXW osvi = { sizeof(osvi), 0, 0, 0, 0, { 0 }, 0, 0 };
    DWORDLONG        const dwlConditionMask = VerSetConditionMask(
        VerSetConditionMask(
            VerSetConditionMask(
                0, VER_MAJORVERSION, VER_GREATER_EQUAL),
            VER_MINORVERSION, VER_GREATER_EQUAL),
        VER_SERVICEPACKMAJOR, VER_GREATER_EQUAL);

    osvi.dwMajorVersion = wMajorVersion;
    osvi.dwMinorVersion = wMinorVersion;
    osvi.wServicePackMajor = wServicePackMajor;

    return VerifyVersionInfoW(&osvi, VER_MAJORVERSION | VER_MINORVERSION | VER_SERVICEPACKMAJOR, dwlConditionMask) != FALSE;
}

inline bool IsWindowsXpOrGreater()
{
    return IsWindowsVersionOrGreater(HIBYTE(_WIN32_WINNT_WINXP), LOBYTE(_WIN32_WINNT_WINXP), 0);
}

string getPathName(const string& s) {

    char sep = '/';

#ifdef _WIN32
    sep = '\\';
#endif

    size_t i = s.rfind(sep, s.length());
    if (i != string::npos) {
        return(s.substr(0, i));
    }

    return("");
}

//=======================================================================
// MAIN PROGRAM
//=======================================================================
extern "C" int WINAPI _tWinMain(HINSTANCE hInstance,
    HINSTANCE /*hPrevInstance*/,
    LPTSTR lpCmdLine,
    int /*nShowCmd*/)
{
//#ifdef _DEBUG
//	int flag = _CrtSetDbgFlag(_CRTDBG_REPORT_FLAG);
//	flag |= _CRTDBG_LEAK_CHECK_DF;
//	_CrtSetDbgFlag(flag);
//#endif

    HRESULT hres = S_OK;
    LPCTSTR szSrvName = _T("OPC DA/AE Server");

#ifdef _ATL_MIN_CRT
    lpCmdLine = GetCommandLine();             // This is line necessary for _ATL_MIN_CRT
#endif

    TCHAR szPath[_MAX_PATH];                 // Get the executable file path
    if (!::GetModuleFileName(NULL, szPath, _MAX_PATH))
        return HRESULT_FROM_WIN32(GetLastError());

#ifdef _OPC_NET
    USES_CONVERSION;
    string pathandfile = W2A(szPath);
#else
    // Retrieve the full path for the current module.
    if (GetModuleFileName(NULL, szPath, sizeof szPath) == 0)
        return -1;
    string pathandfile = szPath;
#endif

    string path = getPathName(pathandfile);

    string file = pathandfile.substr(path.length() + 1);


#if _OPC_NET
    const LPWSTR command = GetCommandLine();
    const string commandLine = CW2A(command);
#else
    const string commandLine = GetCommandLine();
#endif

    int position = commandLine.find(file);
    position += file.length();
    string arguments = commandLine.substr(position);

    LogEnable(path, arguments);

    // Get server name from resource
    szSrvName = _Module.m_VersionInfo.GetValue(_T("FileDescription"));
    _ASSERTE(szSrvName);

    // Display a start message
    LOGFMTI("Starting '%s'", T2CA(szSrvName));

    // Get the Server specific Registry Definitions from the generic part.
    LPSERVERREGDEFS pRegDefs;

    pRegDefs = OnGetServerRegistryDefs();
    _ASSERTE(pRegDefs);

#ifdef _NT_SERVICE
#ifdef _OPC_NET
    LPTSTR szSafeServiceName = CreateSafeServiceName(OLE2T(pRegDefs->m_szAppDescr));
    _ASSERTE(szSafeServiceName);
#else
    USES_CONVERSION;
    LPTSTR szSafeServiceName = CreateSafeServiceName(W2A(pRegDefs->m_szAppDescr));
    _ASSERTE(szSafeServiceName);
#endif
    LOGFMTI("Service Name is '%s'", T2CA(szSafeServiceName));
#endif
    // Initialize the server specific CLSIDs with
    // the values from the generic part.
#ifdef _OPC_SRV_DA
    hres = CLSIDFromString(pRegDefs->m_szDA_CLSID, &CLSID_OPCServer);
    _ASSERTE(SUCCEEDED(hres));
#endif
#ifdef _OPC_SRV_AE
    hres = CLSIDFromString(pRegDefs->m_szAE_CLSID, &CLSID_OPCEventServer);
    if (SUCCEEDED(hres))
    {
        fUseAEServer = true;
    }
    else
    {
        fUseAEServer = false;
        hres = S_OK;
    }
#endif

    // Initialize Global Members
    hres = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    _ASSERTE(SUCCEEDED(hres));
    if (SUCCEEDED(hres)) {
        fCOMInitialized = TRUE;
    }

    hres = CoGetMalloc(MEMCTX_TASK, &pIMalloc);// COM Memory Manager
    _ASSERTE(SUCCEEDED(hres));

    _Module.Init(ObjectMap, hInstance);

    // Create the Component Categories
    hres = CreateComCatOPC();
    _ASSERTE(SUCCEEDED(hres));

    bool gfRegister = FALSE;
    bool fRun = TRUE;

    LOGFMTI("Parse command line arguments '%s'", arguments.c_str());
    if (arguments.length() > 0)
    {
        // parse command line arguments.
        COpcText cText;
        COpcTextReader cReader(arguments.c_str());

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
                if (cFlags == _T("regserver") || cFlags == _T("install"))
                {
                    hres = RegisterServer(pRegDefs, FALSE);
                    fRun = FALSE;
                }
                else if (cFlags == _T("unregserver") || cFlags == _T("deinstall"))
                {
                    hres = UnRegisterServer(pRegDefs);
                    fRun = FALSE;
                }
                else if (cFlags == _T("service"))
                {
                    hres = RegisterServer(pRegDefs, TRUE);
                    fRun = FALSE;
                }
                else {									// Server specific command
                    if (FAILED(OnProcessParam(cFlags))) {
                        fRun = FALSE;					// Invalid argument
                    }
                }
            }
        }
    }

//#ifdef DEBUG
//	if (_CrtDumpMemoryLeaks()) {
//		int breaks = 1;
//	}
//#endif

    // Start the Server
    if (fRun) {
#ifdef _NT_SERVICE
        BOOL fRegisteredAsService;
        if (SUCCEEDED(hres)) {
            hres = IsRegisteredAsService(OLE2CT(pRegDefs->m_szAppID), fRegisteredAsService);
        }
#endif

        if (SUCCEEDED(hres)) {
            hres = _Module.InitializeServer();
        }
        if (SUCCEEDED(hres)) {
            LOGFMTI("Start the server.");
#ifdef _NT_SERVICE
            _Module.Start(fRegisteredAsService, szSafeServiceName);
#else
            _Module.Start(FALSE, NULL);
#endif
#ifdef _OPC_NET
            ::OnTerminateServer();						// Call global OnTerminateServer
#endif
        }
    }
    else
    {
        ::OnTerminateServer();						// Call global OnTerminateServer
    }

#ifdef _NT_SERVICE
    delete[] szSafeServiceName;
#endif

	// Display a stop message
    LOGFMTI("Terminating server");

    //stop logger
    LogManager::getRef().stop();
	
	_Module.m_VersionInfo.Clear();

#ifdef DEBUG
	if (_CrtDumpMemoryLeaks()) {
		int breaks = 1;
	}
#endif

    return hres;
}



//-----------------------------------------------------------------------
// CoreGenericMain Implementation
//-----------------------------------------------------------------------

//=======================================================================
// Constructor
//     The module class is defined in CoreGenericMain.h.
//=======================================================================
CoreGenericMain::CoreGenericMain(void)
{
    m_pGIT = NULL;

    m_fInitialized = FALSE;                // Set server EXE as not initialized
    m_dwThreadID = NULL;
    m_lpszServiceName = NULL;
    m_fStartedAsService = FALSE;
    // Set up the initial service status
    m_hServiceStatus = NULL;
    m_ServiceStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS | SERVICE_INTERACTIVE_PROCESS;
    m_ServiceStatus.dwCurrentState = SERVICE_STOPPED;
    m_ServiceStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP;
    m_ServiceStatus.dwWin32ExitCode = 0;
    m_ServiceStatus.dwServiceSpecificExitCode = 0;
    m_ServiceStatus.dwCheckPoint = 0;
    m_ServiceStatus.dwWaitHint = 0;

    HRESULT hres = m_VersionInfo.Create();
    _ASSERTE(SUCCEEDED(hres));
}



//=======================================================================
// Destructor
//=======================================================================
CoreGenericMain::~CoreGenericMain(void)
{
    if (m_fInitialized) {                        // Terminate server EXE if not done yet      
        ::OnTerminateServer();                    // Call global OnTerminateServer
        m_fInitialized = FALSE;                   // implemented in APP_SERVER.cpp
    }

    if (m_pGIT) {                                // Release the Global Interface Table (GIT)
        m_pGIT->Release();
    }

    if (pIMalloc) {
        pIMalloc->Release();                      // Release the COM Memory Manager
    }
    if (fCOMInitialized) {
        CoUninitialize();                         // Close the COM library
    }
}


//=======================================================================
LONG CoreGenericMain::Unlock()
{
    LONG l = CComModule::Unlock();
    if (l == 0 && !m_fStartedAsService) {
        if (CoSuspendClassObjects() == S_OK) {
            PostThreadMessage(m_dwThreadID, WM_QUIT, 0, 0);
        }
    }
    return l;
}



//=======================================================================
HRESULT CoreGenericMain::InitializeServer(void)
{
    LOGFMTI("Initializing the server.");
    HRESULT hres = S_OK;
    if (m_fInitialized == FALSE) {               // Not yet initialized

        // If the system platform is Windows 2000 we use a Global Interface Table (GIT)
        // to marshal the callback interface pointers
        // (allowing other apartments access to the callback interface pointers).

        if (IsWindowsXpOrGreater())
        {
            // Create Global Interface Table (GIT) for marshaling
            hres = CoCreateInstance(CLSID_StdGlobalInterfaceTable, NULL, CLSCTX_INPROC_SERVER,
                IID_IGlobalInterfaceTable, (LPVOID *)&m_pGIT);
        }
        else
        {
            hres = CO_E_NOT_SUPPORTED;
        }
        if (SUCCEEDED(hres)) {
            hres = ::OnInitializeServer();         // Call global OnInitializeServer
        }
        // implemented in app_server.cpp
        if (SUCCEEDED(hres)) {
            m_fInitialized = TRUE;
        }
    }
    return hres;
}



//=======================================================================
void CoreGenericMain::Start(BOOL fAsService, LPTSTR lpszServiceName)
{
    _ASSERTE(m_fInitialized);                  // Must be initialized

    m_lpszServiceName = lpszServiceName;
    m_fStartedAsService = fAsService;

    SERVICE_TABLE_ENTRY st[] = { { lpszServiceName, _ServiceMain },
    { NULL, NULL }
    };

    if (fAsService && !::StartServiceCtrlDispatcher(st)) {
        m_fStartedAsService = fAsService = FALSE;
    }

    if (fAsService == FALSE)
        Run();
}



//=======================================================================
void CoreGenericMain::Run(void)
{
    m_dwThreadID = GetCurrentThreadId();

    HRESULT hres = _Module.RegisterClassObjects(CLSCTX_LOCAL_SERVER, REGCLS_MULTIPLEUSE);
    _ASSERTE(SUCCEEDED(hres));

    if (m_fStartedAsService) {
        LogEvent(_T("Service started"));
        SetServiceStatus(SERVICE_RUNNING);
    }

    MSG msg;
    while (GetMessage(&msg, 0, 0, 0)) {
        DispatchMessage(&msg);
    }

    _Module.RevokeClassObjects();
}


//=======================================================================
// Static entry point
void WINAPI CoreGenericMain::_ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv)
{
    _Module.ServiceMain(dwArgc, lpszArgv);
}

inline void CoreGenericMain::ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv)
{
    // Register the control request handler
    m_ServiceStatus.dwCurrentState = SERVICE_START_PENDING;
    m_hServiceStatus = RegisterServiceCtrlHandler(m_lpszServiceName, _ServiceHandler);
    if (m_hServiceStatus == NULL) {
        LogEvent(_T("Handler not installed"));
        return;
    }
    SetServiceStatus(SERVICE_START_PENDING);

    m_ServiceStatus.dwWin32ExitCode = S_OK;
    m_ServiceStatus.dwCheckPoint = 0;
    m_ServiceStatus.dwWaitHint = 0;

    Run();                                       // When the Run function returns, the
    // service has stopped.
    SetServiceStatus(SERVICE_STOPPED);
    LogEvent(_T("Service stopped"));
}



//=======================================================================
// Static entry point
void WINAPI CoreGenericMain::_ServiceHandler(DWORD dwOpcode)
{
    _Module.ServiceHandler(dwOpcode);
}

inline void CoreGenericMain::ServiceHandler(DWORD dwOpcode)
{
    switch (dwOpcode) {

    case SERVICE_CONTROL_STOP:
        SetServiceStatus(SERVICE_STOP_PENDING);
        PostThreadMessage(m_dwThreadID, WM_QUIT, 0, 0);
        break;
    case SERVICE_CONTROL_PAUSE:
        break;
    case SERVICE_CONTROL_CONTINUE:
        break;
    case SERVICE_CONTROL_INTERROGATE:
        break;
    case SERVICE_CONTROL_SHUTDOWN:
        break;
    default:
        LogEvent(_T("Bad service request"));
        break;
    }
}


//=======================================================================
void CoreGenericMain::SetServiceStatus(DWORD dwState)
{
    m_ServiceStatus.dwCurrentState = dwState;
    ::SetServiceStatus(m_hServiceStatus, &m_ServiceStatus);
}



//=======================================================================
void CoreGenericMain::LogEvent(LPCTSTR pszFormat, ...)
{
    TCHAR    chMsg[256];
    HANDLE   hEventSource;
    LPTSTR   lpszStrings[1];
    va_list  pArg;

    va_start(pArg, pszFormat);
    _vstprintf_s(chMsg, 256, pszFormat, pArg);
    va_end(pArg);

    lpszStrings[0] = chMsg;

    if (m_fStartedAsService) {
        // Get a handle to use with ReportEvent().
        hEventSource = RegisterEventSource(NULL, m_lpszServiceName);
        if (hEventSource != NULL) {
            // Write to event log.
            ReportEvent(hEventSource, EVENTLOG_INFORMATION_TYPE, 0, 0, NULL, 1, 0, (LPCTSTR *)&lpszStrings[0], NULL);
            DeregisterEventSource(hEventSource);
        }
    }
    else {
        // As we are not running as a service, just write the error to the console.
        _putts(chMsg);
    }
    USES_CONVERSION;
    LOGFMTI(T2A(chMsg));
}



//-----------------------------------------------------------------------
// VersionInformation Implementation
//-----------------------------------------------------------------------

//=======================================================================
VersionInformation::VersionInformation(void)
{
    m_pInfo = NULL;
    m_wMajor = 0;
    m_wMinor = 0;
    m_wBuild = 0;
    // Set default Language ID and Code Page
    _tcsncpy(m_szTranslation, _T("000004E4\\"), sizeof(m_szTranslation) / sizeof(TCHAR));
}



//=======================================================================
VersionInformation::~VersionInformation()
{
    if (m_pInfo)
        delete[] m_pInfo;
}



//=======================================================================
// Create
// ------
//    Reads the Version Info from the resource into a buffer. Also the
//    attributes and the translation string is initialized with the
//    values from the resource file.
//=======================================================================
HRESULT VersionInformation::Create(HMODULE hModule /* = NULL */)
{
    DWORD dwDummyHandle = 0;            // Will always be set to zero.
    DWORD dwLen = 0;
    TCHAR szModuleName[_MAX_PATH];

    // Get the Name of the Server.
    if (GetModuleFileName(hModule, szModuleName, _MAX_PATH) == 0) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    dwLen = GetFileVersionInfoSize(szModuleName, &dwDummyHandle);
    if (!dwLen)
        return HRESULT_FROM_WIN32(GetLastError());

    if (m_pInfo)                        // It's not the firts Create() call
        delete[] m_pInfo;               // Release existing information buffer

    m_pInfo = new BYTE[dwLen];          // Allocate buffer for version info.
    if (!m_pInfo)
        return E_OUTOFMEMORY;

    if (!GetFileVersionInfo(szModuleName, 0, dwLen, m_pInfo))
        return HRESULT_FROM_WIN32(GetLastError());

    WORD* pwVal = NULL;                 // Get translation info
    UINT  uLenVal = 0;
    if (VerQueryValue(m_pInfo, _T("\\VarFileInfo\\Translation"), (LPVOID *)&pwVal, &uLenVal) && (uLenVal >= sizeof(DWORD))) {
        _stprintf(m_szTranslation, _T("%04x%04x\\"), pwVal[0], pwVal[1]);
    }
    // Initialize class attributes
    VS_FIXEDFILEINFO* pFI;
    if (!VerQueryValue(m_pInfo, _T("\\"), (LPVOID *)&pFI, &uLenVal))
        return HRESULT_FROM_WIN32(GetLastError());

    m_wMajor = HIWORD(pFI->dwFileVersionMS);
    m_wMinor = LOWORD(pFI->dwFileVersionMS);
    m_wBuild = LOWORD(pFI->dwFileVersionLS);

    return S_OK;
}

void VersionInformation::Clear()
{
	if (m_pInfo)                        // It's not the firts Create() call
		delete[] m_pInfo;               // Release existing information buffer
}


//=======================================================================
// GetValue
// --------
//    Returns a pointer to a value from the previous loaded version info
//    buffer. If there is no value for the specified key then the
//    function returns NULL.
//    E.g. a key name is "CompanyName".
//=======================================================================
LPCTSTR VersionInformation::GetValue(LPCTSTR szKeyName)
{
    if (m_pInfo) {
        LPCTSTR  pVal;
        UINT     uLenVal;
        TCHAR    szQuery[512] = _T("\\StringFileInfo\\");

        _tcscat(szQuery, m_szTranslation);
        _tcscat(szQuery, szKeyName);

        if (VerQueryValue(m_pInfo, szQuery, (LPVOID *)&pVal, &uLenVal)) {
            return pVal;
        }
    }
    return NULL;
}



//-----------------------------------------------------------------------
// Utility Functions
//
//    Although some of these functions are big they are declared
//    inline since they are only used once.
//-----------------------------------------------------------------------

//=======================================================================
// Creates the Component Categories for OPC Toolkit based Servers.
//=======================================================================
inline static HRESULT CreateComCatOPC(void)
{
    HRESULT hres = S_OK;
    LOGFMTT("Create the Component Categories");

    ICatRegister*  pICat;

    // Get the Component Category Manager
    hres = CoCreateInstance(
        CLSID_StdComponentCategoriesMgr, NULL,
        CLSCTX_INPROC_SERVER, IID_ICatRegister, (LPVOID *)&pICat);

    if (SUCCEEDED(hres)) {
        CATEGORYINFO   CatInfo[4];

        // Set the Category Infos
        CatInfo[0].catid = CATID_OPCDAServer10;
        CatInfo[0].lcid = 0x0409;
        wcscpy(CatInfo[0].szDescription, OPC_CATEGORY_DESCRIPTION_DA10);

        CatInfo[1].catid = CATID_OPCDAServer20;
        CatInfo[1].lcid = 0x0409;
        wcscpy(CatInfo[1].szDescription, OPC_CATEGORY_DESCRIPTION_DA20);

        CatInfo[2].catid = CATID_OPCDAServer30;
        CatInfo[2].lcid = 0x0409;
        wcscpy(CatInfo[2].szDescription, OPC_CATEGORY_DESCRIPTION_DA30);

        CatInfo[3].catid = IID_OPCEventServerCATID;
        CatInfo[3].lcid = 0x0409;
        wcscpy(CatInfo[3].szDescription, OPC_EVENTSERVER_CAT_DESC);

        // Register the Categories
        hres = pICat->RegisterCategories(4, CatInfo);
        LOGFMTT("RegisterCategories finished: 0x%X", hres);
        pICat->Release();

        if (FAILED(hres)) {
            LOGFMTE("CreateComCatOPC() failed with error: Cannot register the Categories: 0x%X", hres);
            return hres;
        }
    }
    else {
        LOGFMTE("CreateComCatOPC() failed with error: Cannot get the Component Category Manager");
        return hres;
    }

    LOGFMTT("Create the Component Categories finished: 0x%X", hres);
    return hres;
}



//=======================================================================
// Server Registration
//
// CLSID_OPCServer and CLSID_OPCEventServer must be initialized.
//=======================================================================
static HRESULT RegisterServer(const LPSERVERREGDEFS pRegDefs, BOOL fAsService)
{
    LOGFMTI("Register the Server as %s", fAsService ? "Service" : "Local Server");

    HRESULT hres;
#ifdef _NT_SERVICE
    USES_CONVERSION;
    LPTSTR lpszServiceName = CreateSafeServiceName(OLE2T(pRegDefs->m_szAppDescr));
    if (!lpszServiceName) {
        hres = E_OUTOFMEMORY;
    }
    else {                                    // Remove any previous service since it
        // may point to the incorrect file.
        hres = UninstallService(lpszServiceName);
    }
    if (SUCCEEDED(hres)) {
#endif // _NT_SERVICE

        SetupReplacementMapForRegistryUpdate(pCustSubstDef, pRegDefs);
        // Register the server as defined in the resource
#ifdef _OPC_SRV_DA
        hres = _Module.UpdateRegistryFromResource(IDR_DATA_ACCESS, TRUE, pCustSubstDef);
        if (SUCCEEDED(hres)) {
#endif
#ifdef _OPC_SRV_AE
            if (fUseAEServer)
            {
                hres = _Module.UpdateRegistryFromResource(IDR_ALARMS_EVENTS, TRUE, pCustSubstDef);
            }
#endif
#ifdef _OPC_SRV_DA
        }
#endif

        CleanupReplacementMapForRegistryUpdate(pCustSubstDef);

#ifdef _NT_SERVICE
    }

    if (SUCCEEDED(hres)) {                     // Adjust the AppID for Local Server or Service.
        hres = AdjustRegistryForServerType(fAsService,
            OLE2CT(pRegDefs->m_szAppID),
            lpszServiceName);
    }

    if (SUCCEEDED(hres) && fAsService) {       // Create Service
        hres = InstallAsService(lpszServiceName, OLE2CT(pRegDefs->m_szAppDescr));
    }
#endif // _NT_SERVICE

    if (SUCCEEDED(hres)) {                     // Registers all objects in the ObjectMap and
        hres = _Module.RegisterServer(            // also the TypeLib as defined by the module.
#ifdef _OPC_SRV_DA
            TRUE        // Only Data Access Server has a TypeLib
#endif
        );
    }

    // Register the implemented Categories
    if (SUCCEEDED(hres)) {

        LOGFMTI("Register the Server as implementing the Categories");
        ICatRegister*  pICat;

        // Get the Component Category Manager
        HRESULT hres = CoCreateInstance(
            CLSID_StdComponentCategoriesMgr, NULL,
            CLSCTX_INPROC_SERVER, IID_ICatRegister, (LPVOID *)&pICat);
        if (SUCCEEDED(hres)) {
            CATID catid[3];

#ifdef _OPC_SRV_DA
            catid[0] = CATID_OPCDAServer10;
            catid[1] = CATID_OPCDAServer20;
            catid[2] = CATID_OPCDAServer30;
            hres = pICat->RegisterClassImplCategories(CLSID_OPCServer, 3, catid);
            if (SUCCEEDED(hres)) {
#endif
#ifdef _OPC_SRV_AE
                catid[0] = IID_OPCEventServerCATID;
                hres = pICat->RegisterClassImplCategories(CLSID_OPCEventServer, 1, catid);
#endif
#ifdef _OPC_SRV_DA
            }
#endif
            pICat->Release();

            if (FAILED(hres)) {
                LOGFMTE("Cannot register the Server as implementing the Categories");
            }
        }
        else {
            LOGFMTI("Cannot get the Component Category Manager");
        }
    }

    if (FAILED(hres)) {
        char szMsg[80];
        sprintf_s(szMsg, 80, "Error 0x%lX. Server Registration failed.", hres);
        LOGFMTI(szMsg);
        USES_CONVERSION;
        MessageBox(NULL, A2T(szMsg), _T("OPC DA/AE Server SDK"), MB_OK);
    }

#ifdef _NT_SERVICE
    if (lpszServiceName) {
        delete[] lpszServiceName;
    }
#endif

    return hres;
}



//=======================================================================
// Server Un-Registration
//
// CLSID_OPCServer and CLSID_OPCEventServer must be initialized.
//=======================================================================
inline static HRESULT UnRegisterServer(const LPSERVERREGDEFS pRegDefs)
{
    LOGFMTI("Unregister the Server");
    // UnRegister the implemented Categories

    HRESULT  hres = S_OK;
#ifdef _NT_SERVICE
    LPTSTR   lpszServiceName = NULL;
#endif // _NT_SERVICE

    // UnRegister the implemented Categories
    LOGFMTI("Unregister the Server as implementing the Categories");
    ICatRegister*  pICat;

    // Get the Component Category Manager
    hres = CoCreateInstance(
        CLSID_StdComponentCategoriesMgr, NULL,
        CLSCTX_INPROC_SERVER, IID_ICatRegister, (LPVOID *)&pICat);
    if (SUCCEEDED(hres)) {
        CATID catid[3];

#ifdef _OPC_SRV_DA
        catid[0] = CATID_OPCDAServer10;
        catid[1] = CATID_OPCDAServer20;
        catid[2] = CATID_OPCDAServer30;
        hres = pICat->UnRegisterClassImplCategories(CLSID_OPCServer, 3, catid);
        if (SUCCEEDED(hres)) {
#endif
#ifdef _OPC_SRV_AE
            catid[0] = IID_OPCEventServerCATID;
            hres = pICat->UnRegisterClassImplCategories(CLSID_OPCEventServer, 1, catid);
#endif
#ifdef _OPC_SRV_DA
        }
#endif
        pICat->Release();

        if (FAILED(hres)) {
            LOGFMTE("Cannot unregister the Server as implementing the Categories");
        }
    }
    else {
        LOGFMTE("Cannot get the Component Category Manager");
    }

#ifdef _NT_SERVICE
    if (SUCCEEDED(hres)) {
        USES_CONVERSION;
        lpszServiceName = CreateSafeServiceName(OLE2T(pRegDefs->m_szAppDescr));
        if (!lpszServiceName) {
            hres = E_OUTOFMEMORY;
        }
        else {
            // Uninstall from Service Manager
            hres = UninstallService(lpszServiceName);
            if (SUCCEEDED(hres)) {
                // Removes the values from the AppID if
                hres = AdjustRegistryForServerType( // registered as Service.
                    FALSE,
                    OLE2CT(pRegDefs->m_szAppID),
                    lpszServiceName);
            }
        }
    }
#endif // _NT_SERVICE

    if (SUCCEEDED(hres)) {

        SetupReplacementMapForRegistryUpdate(pCustSubstDef, pRegDefs);
        // Unregister the server as defined in the resource
#ifdef _OPC_SRV_DA
        hres = _Module.UpdateRegistryFromResource(IDR_DATA_ACCESS, FALSE, pCustSubstDef);
        if (SUCCEEDED(hres)) {
#endif
#ifdef _OPC_SRV_AE
            hres = _Module.UpdateRegistryFromResource(IDR_ALARMS_EVENTS, FALSE, pCustSubstDef);
#endif
#ifdef _OPC_SRV_DA
        }
#endif

        CleanupReplacementMapForRegistryUpdate(pCustSubstDef);
    }

    if (SUCCEEDED(hres)) {
        hres = _Module.UnregisterServer();        // Unregisters all objects defined in the ObjectMap.
    }

    if (FAILED(hres)) {
        char szMsg[80];
        sprintf_s(szMsg, 80, "Error 0x%lX. Server Un-Registration failed.", hres);
        LOGFMTI(szMsg);
        USES_CONVERSION;
        MessageBox(NULL, A2T(szMsg), _T("OPC Server SDK Classic"), MB_OK);
    }

#ifdef _NT_SERVICE
    if (lpszServiceName) {
        delete[] lpszServiceName;
    }
#endif

    return hres;
}



//=======================================================================
// Defines the substitution string replacement for the custom Interface
// registry definitions.
//=======================================================================
static void SetupReplacementMapForRegistryUpdate(_ATL_REGMAP_ENTRY aMapEntries[],
    const LPSERVERREGDEFS pRegDefs)
{
    USES_CONVERSION;
    int         i;
    TCHAR       szPath[_MAX_PATH];

    GetModuleFileName(_Module.m_hInst, szPath, _MAX_PATH);          // Get module name and truncate path
    for (i = (int)_tcslen(szPath); (TCHAR)(szPath[i - 1]) != (TCHAR)'\\'; --i);

    aMapEntries[0].szData = _wcsdup(T2W(&szPath[i]));             // Module name without path
    aMapEntries[1].szData = _wcsdup(pRegDefs->m_szAppID);         // Application ID
    aMapEntries[2].szData = _wcsdup(pRegDefs->m_szAppDescr);      // Application description

    aMapEntries[3].szData = _wcsdup(pRegDefs->m_szDA_CLSID);      // Class ID of Data Access server
    aMapEntries[4].szData = _wcsdup(pRegDefs->m_szDA_IndepProgID);// Version independent ProgID
    aMapEntries[5].szData = _wcsdup(pRegDefs->m_szDA_CurrProgID); // ProgID of current server
    aMapEntries[6].szData = _wcsdup(pRegDefs->m_szDA_IndepDescr); // Friendly name of version independent server
    aMapEntries[7].szData = _wcsdup(pRegDefs->m_szDA_CurrDescr);  // Friendly name of current server version

    aMapEntries[8].szData = _wcsdup(pRegDefs->m_szAE_CLSID);      // Class ID of Data Access server
    aMapEntries[9].szData = _wcsdup(pRegDefs->m_szAE_IndepProgID);// Version independent ProgID
    aMapEntries[10].szData = _wcsdup(pRegDefs->m_szAE_CurrProgID); // ProgID of current server
    aMapEntries[11].szData = _wcsdup(pRegDefs->m_szAE_IndepDescr); // Friendly name of version independent server
    aMapEntries[12].szData = _wcsdup(pRegDefs->m_szAE_CurrDescr);  // Friendly name of current server version

    // Get the vendor name from the resource
    aMapEntries[13].szData = _wcsdup(gVendorName);

}



//=======================================================================
// Deletes all defined substitution strings from the replacement map.
//=======================================================================
static void CleanupReplacementMapForRegistryUpdate(_ATL_REGMAP_ENTRY aMapEntries[])
{
    _ATL_REGMAP_ENTRY*   pEntry = aMapEntries;

    while (pEntry->szKey) {
        if (pEntry->szData) {
            free((LPVOID)pEntry->szData);
            pEntry->szData = NULL;
        }
        pEntry++;
    }
}



//-----------------------------------------------------------------------
// Find one the defined characters
//-----------------------------------------------------------------------
static LPCTSTR FindOneOf(LPCTSTR p1, LPCTSTR p2)
{
    while (*p1) {
        LPCTSTR p = p2;
        while (*p != NULL) {
            if (*p1 == *p++)
                return p1 + 1;
        }
        p1++;
    }
    return NULL;
}



#ifdef _NT_SERVICE
//-----------------------------------------------------------------------
// Service Utility Functions
//-----------------------------------------------------------------------

//=======================================================================
// Adjusts the AppID for Local Server or Service
//=======================================================================
static HRESULT AdjustRegistryForServerType(BOOL fRegisterAsService,
    LPCTSTR lpszAppID, LPCTSTR lpszServiceName)
{
    _ASSERT(lpszAppID);
    _ASSERT(lpszServiceName);
    CRegKey  keyAppID, key;
    LONG     lRes;

    // Open the key HKEY_CLASSES_ROOT\AppID\{MyServerAppID}
    lRes = keyAppID.Open(HKEY_CLASSES_ROOT, _T("AppID"), KEY_WRITE);
    if (lRes == ERROR_SUCCESS) {

        lRes = key.Open(keyAppID, lpszAppID, KEY_WRITE);
        if (lRes == ERROR_SUCCESS) {
            // key HKEY_CLASSES_ROOT\AppID\{MyServerAppID} is open
            // Remove the values.
            key.DeleteValue(_T("LocalService")); // Cause no error if values does not exist.
            key.DeleteValue(_T("ServiceParameters"));

            if (fRegisterAsService) {              // Register the Server as Service.
                // Add the values.
                lRes = key.CREGKEY_SetStringValue(_T("LocalService"), lpszServiceName);

                if (lRes == ERROR_SUCCESS) {
                    lRes = key.CREGKEY_SetStringValue(_T("ServiceParameters"), _T("-Service"));
                }
            }
        }
    }
    return HRESULT_FROM_WIN32(lRes);
}



inline static HRESULT IsRegisteredAsService(LPCTSTR lpszAppID, BOOL & fAsService)
{
    _ASSERT(lpszAppID);
    CRegKey  keyAppID, key;
    LONG     lRes;

    fAsService = FALSE;

    // Open the key HKEY_CLASSES_ROOT\AppID\{MyServerAppID}
    lRes = keyAppID.Open(HKEY_CLASSES_ROOT, _T("AppID"), KEY_READ);
    if (lRes == ERROR_SUCCESS) {

        lRes = key.Open(keyAppID, lpszAppID, KEY_READ);
        if (lRes == ERROR_SUCCESS) {
            // key HKEY_CLASSES_ROOT\AppID\{MyServerAppID} is open

            TCHAR szValue[_MAX_PATH];
            DWORD dwLen = _MAX_PATH;

            lRes = key.CREGKEY_QueryStringValue(_T("LocalService"), szValue, &dwLen);
            if (lRes == ERROR_SUCCESS) {
                LOGFMTI("The server will be started as service.");
                fAsService = TRUE;
            }
            return S_OK;
        }
    }
    return HRESULT_FROM_WIN32(lRes);
}



static BOOL IsInstalledAsService(LPCTSTR lpszServiceName)
{
    BOOL fResult = FALSE;

    SC_HANDLE hSCM = ::OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);
    if (hSCM != NULL) {
        SC_HANDLE hService = ::OpenService(hSCM, lpszServiceName, SERVICE_QUERY_CONFIG);
        if (hService != NULL) {
            fResult = TRUE;
            ::CloseServiceHandle(hService);
        }
        ::CloseServiceHandle(hSCM);
    }
    return fResult;
}



inline static HRESULT UninstallService(LPCTSTR lpszServiceName)
{
    if (!IsInstalledAsService(lpszServiceName)) {
        return S_OK;                              // Not installed
    }

    SC_HANDLE hSCM = ::OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);
    if (hSCM == NULL) {
        MessageBox(NULL, _T("Couldn't open service manager"), lpszServiceName, MB_OK);
        return HRESULT_FROM_WIN32(GetLastError());
    }

    SC_HANDLE hService = ::OpenService(hSCM, lpszServiceName, SERVICE_STOP | DELETE);
    if (hService == NULL) {
        ::CloseServiceHandle(hSCM);
        MessageBox(NULL, _T("Couldn't open service"), lpszServiceName, MB_OK);
        return HRESULT_FROM_WIN32(GetLastError());
    }

    SERVICE_STATUS status;
    ::ControlService(hService, SERVICE_CONTROL_STOP, &status);


    BOOL bDelete = ::DeleteService(hService);
    ::CloseServiceHandle(hService);
    ::CloseServiceHandle(hSCM);

    if (bDelete)
        return S_OK;

    MessageBox(NULL, _T("Service could not be deleted"), lpszServiceName, MB_OK);
    return HRESULT_FROM_WIN32(GetLastError());
}



inline static HRESULT InstallAsService(LPCTSTR lpszServiceName, LPCTSTR lpszServiceDescription)
{
    if (IsInstalledAsService(lpszServiceName)) {
        return S_OK;                              // Already installed
    }

    TCHAR szFilePath[_MAX_PATH];                 // Get the executable file path
    if (!::GetModuleFileName(NULL, szFilePath, _MAX_PATH))
        return HRESULT_FROM_WIN32(GetLastError());

    SC_HANDLE hSCM = ::OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);
    if (hSCM == NULL) {
        MessageBox(NULL, _T("Couldn't open service manager"), lpszServiceName, MB_OK);
        return HRESULT_FROM_WIN32(GetLastError());
    }

    SC_HANDLE hService = ::CreateService(
        hSCM, lpszServiceName, lpszServiceDescription,
        SERVICE_ALL_ACCESS, SERVICE_WIN32_OWN_PROCESS | SERVICE_INTERACTIVE_PROCESS,
        SERVICE_DEMAND_START, SERVICE_ERROR_NORMAL,
        szFilePath, NULL, NULL, _T("RPCSS\0"), NULL, NULL);

    if (hService == NULL) {
        ::CloseServiceHandle(hSCM);
        MessageBox(NULL, _T("Couldn't create service"), lpszServiceName, MB_OK);
        return HRESULT_FROM_WIN32(GetLastError());
    }

    SERVICE_DESCRIPTION ServiceDescr;
    ServiceDescr.lpDescription = (LPTSTR)lpszServiceDescription;
    if (!ChangeServiceConfig2(hService, SERVICE_CONFIG_DESCRIPTION, &ServiceDescr)) {
        MessageBox(NULL, _T("Cannot set the description of the service"), lpszServiceName, MB_OK);
    }

    ::CloseServiceHandle(hService);
    ::CloseServiceHandle(hSCM);
    return S_OK;
}



//=======================================================================
// Creates a name which can be used as Service Name.
// The created name does not include any white-space, forward-slash (/)
// or back-slash (\) characters.
// The caller is responsible for freeing the memory allocated for the
// string by calling the delete [] operator.
//=======================================================================
static LPTSTR CreateSafeServiceName(LPCTSTR szNameBase)
{
    LPTSTR   lpszDest, p;

    lpszDest = p = new TCHAR[_tcslen(szNameBase) + 1];
    if (!p) return NULL;

    while (*szNameBase) {
        if ((_istspace(*szNameBase) ||
            ((*szNameBase) == TCHAR('\\')) ||
            ((*szNameBase) == TCHAR('/'))) == 0) {

            *p = *szNameBase;
            p = _tcsinc(p);
        }
        szNameBase = _tcsinc(szNameBase);
    }
    *p = TCHAR('\0');
    return lpszDest;
}

#endif // _NT_SERVICE

//DOM-IGNORE-END
