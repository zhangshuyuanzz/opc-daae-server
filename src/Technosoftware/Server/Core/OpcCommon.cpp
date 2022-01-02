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

//DOM-IGNORE-BEGIN

//-------------------------------------------------------------------------
// INLCUDE
//-------------------------------------------------------------------------
#include "stdafx.h"
#include "UtilityFuncs.h"
#include "OpcCommon.h"
#include "Logger.h"

//-------------------------------------------------------------------------
// DEFINES
//-------------------------------------------------------------------------.
/** @brief	The proxy name for OPC DA. */
#define  PROXY_NAME_OPC_DA    _T("opcproxy")
/** @brief	The proxy name for OPC AE. */
#define  PROXY_NAME_OPC_AE    _T("opc_aeps")

//-------------------------------------------------------------------------
// CODE
//-------------------------------------------------------------------------

/**
 * @fn	OpcCommon::OpcCommon()
 *
 * @brief	Constructor
 */

OpcCommon::OpcCommon()
{
   m_dwLCID = LOCALE_SYSTEM_DEFAULT;
}

/**
 * @fn	HRESULT OpcCommon::Create()
 *
 * @brief	=========================================================================
 * 			 Initializer
 * 			 -----------
 * 			    Must be called after construction.
 * 			=========================================================================.
 *
 * @return	A hResult.
 */

HRESULT OpcCommon::Create()
{
   return ClientName.SetString( L"Unspecified" );
}

/**
 * @fn	OpcCommon::~OpcCommon()
 *
 * @brief	=========================================================================
 * 			 Destructor
 * 			=========================================================================.
 */

OpcCommon::~OpcCommon()
{
	// do nothing.
}



///////////////////////////////////////////////////////////////////////////
///////////////////////////////// IOPCCommon //////////////////////////////
///////////////////////////////////////////////////////////////////////////

/**
 * @fn	STDMETHODIMP OpcCommon::SetLocaleID( LCID dwLcid )
 *
 * @brief	=========================================================================
 * 			 IOPCCommon::SetLocaleID                                        INTERFACE
 * 			=========================================================================.
 *
 * @param	dwLcid	The lcid.
 *
 * @return	A STDMETHODIMP.
 */

STDMETHODIMP OpcCommon::SetLocaleID(
   /* [in] */                    LCID        dwLcid )
{
   LOGFMTI( "IOPCCommon::SetLocaleID, %lX", dwLcid );
   DWORD    dwCount, i;
   LCID*    pdwLcid;

                                             // Check if the requested LCID is supported
   HRESULT  hres = QueryAvailableLocaleIDs( &dwCount, &pdwLcid );
   if (SUCCEEDED( hres )) {
      for (i=0; i<dwCount; i++) {
         if (pdwLcid[i] == dwLcid) {
            m_csMembers.Lock();
            m_dwLCID = dwLcid;
            m_csMembers.Unlock();
            break;                           // The requested LCID is supported
         }
      }
      pIMalloc->Free( pdwLcid );

      hres = (i==dwCount) ? E_INVALIDARG     // The requested LCID is not supported
                          : S_OK;            // The requested LCID is supported
   }
   return hres;
}

/**
 * @fn	STDMETHODIMP OpcCommon::GetLocaleID( LCID * pdwLcid )
 *
 * @brief	=========================================================================
 * 			 IOPCCommon::GetLocaleID                                        INTERFACE
 * 			=========================================================================.
 *
 * @param [in,out]	pdwLcid	If non-null, the pdw lcid.
 *
 * @return	The locale identifier.
 */

STDMETHODIMP OpcCommon::GetLocaleID(
   /* [out] */                   LCID     *  pdwLcid )
{
   LOGFMTI( "IOPCCommon::GetLocaleID" );
   m_csMembers.Lock();
   *pdwLcid = m_dwLCID;
   m_csMembers.Unlock();
   return S_OK;
}

/**
 * @fn	STDMETHODIMP OpcCommon::QueryAvailableLocaleIDs( DWORD * pdwCount, LCID ** pdwLcid )
 *
 * @brief	=========================================================================
 * 			 IOPCCommon::QueryAvailableLocaleIDs                            INTERFACE
 * 			=========================================================================.
 *
 * @param [in,out]	pdwCount	If non-null, number of pdws.
 * @param [in,out]	pdwLcid 	If non-null, the pdw lcid.
 *
 * @return	The available locale i ds.
 */

STDMETHODIMP OpcCommon::QueryAvailableLocaleIDs(
   /* [out] */                   DWORD    *  pdwCount,
   /* [size_is][size_is][out] */ LCID     ** pdwLcid )
{
   LOGFMTI( "IOPCCommon::QueryAvailableLocaleIDs" );
   *pdwCount = 4;

   *pdwLcid = (LCID *)pIMalloc->Alloc( *pdwCount * sizeof (LCID) );
   if (*pdwLcid == NULL) {
      *pdwCount = 0;
      return E_OUTOFMEMORY;
   }

   (*pdwLcid)[0] = LOCALE_SYSTEM_DEFAULT;
   (*pdwLcid)[1] = LOCALE_USER_DEFAULT;
   (*pdwLcid)[2] = LOCALE_NEUTRAL;
   (*pdwLcid)[3] =  //  English American (English Default)
                    MAKELCID( MAKELANGID( LANG_ENGLISH, SUBLANG_ENGLISH_US ), SORT_DEFAULT );
   return S_OK;
}

/**
 * @fn	STDMETHODIMP OpcCommon::GetErrorString( HRESULT dwError, LPWSTR * ppString )
 *
 * @brief	=========================================================================
 * 			 IOPCCommon::GetErrorString                                     INTERFACE
 * 			=========================================================================.
 *
 * @param	dwError				The error.
 * @param [in,out]	ppString	If non-null, the string.
 *
 * @return	The error string.
 */

STDMETHODIMP OpcCommon::GetErrorString(
   /* [in] */                    HRESULT     dwError,
   /* [string][out] */           LPWSTR   *  ppString )
{
   LOGFMTI( "IOPCCommon::GetErrorString for error 0x%lX", dwError );

   m_csMembers.Lock();
   LCID dwCurrenLCID = m_dwLCID;
   m_csMembers.Unlock();
   HRESULT hr = GetErrorString( dwError, dwCurrenLCID, ppString );
   if (SUCCEEDED( hr )) {
      USES_CONVERSION;
      LOGFMTI( "   Returns: %s", W2A(*ppString) );
   }
   return hr;
}

/**
 * @fn	STDMETHODIMP OpcCommon::SetClientName( LPCWSTR szName )
 *
 * @brief	=========================================================================
 * 			 IOPCCommon::SetClientName                                      INTERFACE
 * 			=========================================================================.
 *
 * @param	szName	The name.
 *
 * @return	A STDMETHODIMP.
 */

STDMETHODIMP OpcCommon::SetClientName(
   /* [string][in] */            LPCWSTR     szName )
{
	USES_CONVERSION;
	LOGFMTI( "IOPCCommon::SetClientName to %s", W2A(szName) );

   m_csMembers.Lock();
   HRESULT hres = ClientName.SetString( szName );
   m_csMembers.Unlock();
   return hres;
}

////////////////////////////////////////////////////////////////////////////



//-------------------------------------------------------------------------
// OPERATION
//-------------------------------------------------------------------------

/**
 * @fn	HRESULT OpcCommon::GetErrorString( HRESULT dwError, LCID dwLcid, LPWSTR * ppString )
 *
 * @brief	=========================================================================
 * 			 GetErrorString
 * 			 --------------
 * 			    Gets an error message for an OPC or system error code using the specified locale
 * 			    idenitifier.
 * 			=========================================================================.
 *
 * @param	dwError				The error.
 * @param	dwLcid				The lcid.
 * @param [in,out]	ppString	If non-null, the string.
 *
 * @return	The error string.
 */

HRESULT OpcCommon::GetErrorString(
                  /* [in] */                    HRESULT        dwError,
                  /* [in] */                    LCID           dwLcid,
                  /* [string][out] */           LPWSTR      *  ppString
                  )
{
   *ppString = NULL;

   HRESULT  hres;
   if  (HRESULT_FACILITY( dwError ) == FACILITY_ITF) {
                                                // Specified error code is an OPC or CALL-R error.
                                                // Note : It's for all english languages the same text.
      hres = GetOPCErrorString( dwError, dwLcid, ppString );
   }
   else {
                                                // Specified error is not an OPC error.
                                                // Look for error message in the system message table.
      hres = GetErrorStringFromModule( NULL, dwError, dwLcid, ppString );
   }
   return hres;
}

/**
 * @fn	HRESULT OpcCommon::GetClientName( LPWSTR * pszName )
 *
 * @brief	=========================================================================
 * 			 GetClientName
 * 			 -------------
 * 			    Returns a text string copy of the client name. Memory is allocated from the
 * 			    Global COM Memory Manager.
 * 			=========================================================================.
 *
 * @param [in,out]	pszName	If non-null, the name.
 *
 * @return	The client name.
 */

HRESULT OpcCommon::GetClientName(
                  /* [string][out] */           LPWSTR      *  pszName
                  )
{
   m_csMembers.Lock();
   *pszName = ClientName.CopyCOM();
   m_csMembers.Unlock();
   return *pszName ? S_OK : E_OUTOFMEMORY;
}



//-------------------------------------------------------------------------
// IMPLEMENTATION
//-------------------------------------------------------------------------

/**
 * @fn	HRESULT OpcCommon::GetOPCErrorString( HRESULT dwError, LCID dwLcid, LPWSTR * ppString )
 *
 * @brief	=========================================================================
 * 			 GetOPCErrorString                                               INTERNAL
 * 			 -----------------
 * 			    Gets an error message for an OPC error code using the specified locale
 * 			    idenitifier.
 * 			
 * 			    Note : Only the neutral and english locales are supported.
 * 			
 * 			 parameter:
 * 			    dwError                    error ocde dwLcid                     the locale for
 * 			    the returned string ppString                   where the error string will be
 * 			    saved
 * 			
 * 			 return:
 * 			    S_OK                       Error message found E_INVALIDARG               It is
 * 			    not an OPC specific error code E_OUTOFMEMORY              Not enough memory
 * 			    DISP_E_UNKNOWNLCID;        Requested locale is not supported
 * 			=========================================================================.
 *
 * @param	dwError				The error.
 * @param	dwLcid				The lcid.
 * @param [in,out]	ppString	If non-null, the string.
 *
 * @return	The opc error string.
 */

HRESULT OpcCommon::GetOPCErrorString(
   /* [in] */                    HRESULT     dwError,
   /* [in] */                    LCID        dwLcid,
   /* [string][out] */           LPWSTR   *  ppString )
{
   _ASSERTE( ppString );
   *ppString = NULL;

                                                // Primary language identifier 
   USHORT usPrimaryLanguage = PRIMARYLANGID( LANGIDFROMLCID( dwLcid ) );

   if ((usPrimaryLanguage != LANG_ENGLISH) &&
       (usPrimaryLanguage != LANG_NEUTRAL)) {
      return DISP_E_UNKNOWNLCID;                // Requested locale is not supported.
   }

   //
   // Get the name of the Proxy where the text is read from
   //
   LPCTSTR pszProxyName = NULL;
   CComQIPtr<IOPCServer, &IID_IOPCServer> pDAServer(this);
   if (pDAServer) {
      pszProxyName = PROXY_NAME_OPC_DA;
   }
   else {
      CComQIPtr<IOPCEventServer, &IID_IOPCEventServer> pAEServer(this);
      if (pAEServer) {
         pszProxyName = PROXY_NAME_OPC_AE;
      }  
   }
   if (!pszProxyName) {
      return E_INVALIDARG;                      // Neither a DA Server nor a AE Server
   }

   //
   // Read the text from the Proxy
   //
   BOOL fLoaded = FALSE;
   HMODULE hModule = GetModuleHandle( pszProxyName );
   if (hModule == NULL) {
      hModule = LoadLibrary( pszProxyName );
      if (hModule == NULL) {                    // The Proxy for the appropriate Server Interface must exist
         return E_FAIL;                         // Don't return detailed error code available with
      }                                         // GetLastError() because not defined as function result value.
      else {
         fLoaded = TRUE;
      }
   }
   HRESULT hr = GetErrorStringFromModule( hModule, dwError, dwLcid, ppString );
   if (fLoaded) {
      FreeLibrary( hModule );                   // Release the loaded library
   }
   return hr;
}

/**
 * @fn	HRESULT OpcCommon::GetErrorStringFromModule( HMODULE hModule, HRESULT dwError, LCID dwLcid, LPWSTR * ppString )
 *
 * @brief	=========================================================================
 * 			 GetErrorStringFromModule                                        INTERNAL
 * 			 ------------------------
 * 			    Gets an error message from the message table of the specified module or from the
 * 			    system.
 * 			
 * 			 parameter:
 * 			    hModule                    searches in the module with this handle
 * 			                               or in the system if NULL
 * 			    dwError                    error ocde dwLcid                     the locale for
 * 			    the returned string ppString                   where the error string will be
 * 			    saved
 * 			
 * 			 return:
 * 			    S_OK                       Error message found E_INVALIDARG               Error
 * 			    message not found E_OUTOFMEMORY              Not enough memory DISP_E_UNKNOWNLCID;
 * 			    Requested locale is not supported
 * 			=========================================================================.
 *
 * @param	hModule				The module.
 * @param	dwError				The error.
 * @param	dwLcid				The lcid.
 * @param [in,out]	ppString	If non-null, the string.
 *
 * @return	The error string from module.
 */

HRESULT OpcCommon::GetErrorStringFromModule(
   /* [in] */                    HMODULE     hModule,
   /* [in] */                    HRESULT     dwError,
   /* [in] */                    LCID        dwLcid,
   /* [string][out] */           LPWSTR   *  ppString )
{
   HRESULT  hres;
   DWORD    dwRTC;
   LPTSTR   lpMsgBuf = NULL;
   DWORD    dwFlags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_IGNORE_INSERTS;

   if (hModule)   dwFlags |= FORMAT_MESSAGE_FROM_HMODULE;
   else           dwFlags |= FORMAT_MESSAGE_FROM_SYSTEM;

   dwRTC = FormatMessage(
                     dwFlags,
                     hModule,                   // From Module or System
                     dwError,                   // Error Code
                     LANGIDFROMLCID( dwLcid ),  // Language Identifier
                     (LPTSTR)&lpMsgBuf,         // Output string buffer
                     0,                         // no static message buffer
                     NULL );                    // no message insertion.

   if (dwRTC == 0) {                            // Error occured.
      if (GetLastError() == ERROR_RESOURCE_LANG_NOT_FOUND) {
         hres = DISP_E_UNKNOWNLCID;             // There is no message defined for the specified locale.
      }
      else {
         return E_INVALIDARG;                   // There is no message defined for the specified error.
      }
   }
   else {                                       // Message found.

      TCHAR* pc = &lpMsgBuf[ dwRTC-1 ];         // Truncate \r\n
      while (pc && (*pc == (TCHAR)'\r' || (*pc == (TCHAR)'\n'))) {
         *pc = (TCHAR)'\0';
         pc--;
         dwRTC--;
      }
   
      USES_CONVERSION;
      *ppString = WSTRClone( T2COLE( lpMsgBuf ), pIMalloc );
      hres = (*ppString) ? S_OK : E_OUTOFMEMORY;
   }
   if (lpMsgBuf) {                              // Free the temporary buffer.
      LocalFree( lpMsgBuf );
   }
   return hres;
}

//DOM-IGNORE-END
