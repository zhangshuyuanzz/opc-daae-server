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

#ifndef __OpcCommon_H
#define __OpcCommon_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "WideString.h"                         // for WideString


/////////////////////////////////////////////////////////////////////////////
// Class declaration
/////////////////////////////////////////////////////////////////////////////
class OpcCommon : public IOPCCommon
{
// Construction
public:
   OpcCommon();
   HRESULT  Create();

//  Destruction
   ~OpcCommon();

// Operations
public:
   ///////////////////////////////////////////////////////////////////////////
   ///////////////////////////////// IOPCCommon //////////////////////////////
   ///////////////////////////////////////////////////////////////////////////

   STDMETHODIMP SetLocaleID(
                  /* [in] */                    LCID           dwLcid
                  );

   STDMETHODIMP GetLocaleID(
                  /* [out] */                   LCID        *  pdwLcid
                  );

   STDMETHODIMP QueryAvailableLocaleIDs(
                  /* [out] */                   DWORD       *  pdwCount,
                  /* [size_is][size_is][out] */ LCID        ** pdwLcid
                  );

   STDMETHODIMP GetErrorString(
                  /* [in] */                    HRESULT        dwError,
                  /* [string][out] */           LPWSTR      *  ppString
                  );

   STDMETHODIMP SetClientName(
                  /* [string][in] */            LPCWSTR        szName
                  );

   ///////////////////////////////////////////////////////////////////////////

   HRESULT  GetErrorString(
                  /* [in] */                    HRESULT        dwError,
                  /* [in] */                    LCID           dwLcid,
                  /* [string][out] */           LPWSTR      *  ppString
                  );

   HRESULT  GetClientName(
                  /* [string][out] */           LPWSTR     *   pszName
                  );


public:
    WideString ClientName;                      // The name of the client.

// Impmementation
protected:
   LCID        m_dwLCID;               // The default locale ID for the server/client session.

   CComAutoCriticalSection             // Protect data member access.
               m_csMembers;            // Functions are called from different threads.

   HRESULT GetOPCErrorString( HRESULT dwError, LCID dwLcid, LPWSTR* ppString );
   HRESULT GetErrorStringFromModule( HMODULE hModule, HRESULT dwError, LCID dwLcid, LPWSTR* ppString );
};
//DOM-IGNORE-END

#endif // __OpcCommon_H

