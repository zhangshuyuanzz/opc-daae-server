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

#ifndef __CoreGenericMain_H_
#define __CoreGenericMain_H_

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000


//-----------------------------------------------------------------------
// CLASS VersionInformation
//-----------------------------------------------------------------------
//
// With this internal class you have access to the
// version information stored in the resource.
//
class VersionInformation
{
public:
// Constructions
   VersionInformation( void );
   ~VersionInformation();

// Attributes
   WORD     m_wMajor;         // Major Product Version
   WORD     m_wMinor;         // Minor Product Version
   WORD     m_wBuild;         // Build Number of Product

// Operations
   HRESULT  Create( HMODULE hModule = nullptr );
   LPCTSTR  GetValue( LPCTSTR szKeyName );
   void     Clear();

// Implementation
protected:
   BYTE* m_pInfo;
   TCHAR m_szTranslation[10];
};


//-----------------------------------------------------------------------
// CLASS CoreGenericMain
//-----------------------------------------------------------------------
//
// You may derive a class from CComModule and use it if you want to
// override something, but do not change the name of _Module.
//
class CoreGenericMain : public CComModule
{
public:
// Constructions
   CoreGenericMain( void );
   ~CoreGenericMain( void );

// Attributes
   VersionInformation            m_VersionInfo;
   LPGLOBALINTERFACETABLE  m_pGIT;     // Global Interface Table for
                                       // marshaling of callback interface pointers.

// Operations
   HRESULT  InitializeServer( void );
   void     Start( BOOL fAsService, LPTSTR lpszServiceName );

// Overridables
   LONG     Unlock();

// Implementation
protected:
   BOOL                    m_fInitialized;
   DWORD                   m_dwThreadID;
   SERVICE_STATUS_HANDLE   m_hServiceStatus;
   SERVICE_STATUS          m_ServiceStatus;
   LPTSTR                  m_lpszServiceName;
   BOOL                    m_fStartedAsService;

   void     Run( void );
   void     ServiceMain( DWORD dwArgc, LPTSTR* lpszArgv );
   void     ServiceHandler( DWORD dwOpcode );
   void     SetServiceStatus( DWORD dwState );
   void     LogEvent( LPCTSTR pszFormat, ... );

private:
   static void WINAPI _ServiceMain( DWORD dwArgc, LPTSTR* lpszArgv );
   static void WINAPI _ServiceHandler( DWORD dwOpcode );
};


//-----------------------------------------------------------------------
// EXTERNALs
//-----------------------------------------------------------------------

extern CoreGenericMain core_generic_main;
extern IMalloc*   pIMalloc;

EXTERN_C CLSID CLSID_OPCServer;

//DOM-IGNORE-END

#endif // __CoreGenericMain_H_

