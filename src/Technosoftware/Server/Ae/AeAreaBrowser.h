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

#ifndef __EventAreaBrowser_H
#define __EventAreaBrowser_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

class EventArea;

EXTERN_C CLSID CLSID_IOPCEventAreaBrowser;


//-----------------------------------------------------------------------
// CLASS
//-----------------------------------------------------------------------
class ATL_NO_VTABLE EventAreaBrowser :
   public CComObjectRootEx<CComMultiThreadModel>,
   public CComCoClass<EventAreaBrowser, &CLSID_IOPCEventAreaBrowser>,
   public IOPCEventAreaBrowser
{
public:
   ///////////////////////////////////////////////////////////////////////////
   // ATL macros
   ///////////////////////////////////////////////////////////////////////////
                                                // Eentry points for COM
   BEGIN_COM_MAP(EventAreaBrowser)
      COM_INTERFACE_ENTRY(IOPCEventAreaBrowser)
   END_COM_MAP()

   DECLARE_NO_REGISTRY()                        // Avoid default ATL registration
   DECLARE_NOT_AGGREGATABLE(EventAreaBrowser)

   ///////////////////////////////////////////////////////////////////////////

// Construction / Destruction
public:
   EventAreaBrowser();
   HRESULT  Create( EventArea* pRootArea );
   ~EventAreaBrowser();

// Operations
public:
   ///////////////////////////////////////////////////////////////////////////
   //////////////////////////// IOPCEventAreaBrowser /////////////////////////
   ///////////////////////////////////////////////////////////////////////////

   STDMETHODIMP ChangeBrowsePosition(
                  /* [in] */                    OPCAEBROWSEDIRECTION
                                                                  dwBrowseDirection,
                  /* [string][in] */            LPCWSTR           szString
                  );
        
   STDMETHODIMP BrowseOPCAreas(
                  /* [in] */                    OPCAEBROWSETYPE   dwBrowseFilterType,
                  /* [string][in] */            LPCWSTR           szFilterCriteria,
                  /* [out] */                   LPENUMSTRING   *  ppIEnumString
                  );
        
   STDMETHODIMP GetQualifiedAreaName(
                  /* [in] */                    LPCWSTR           szAreaName,
                  /* [string][out] */           LPWSTR         *  pszQualifiedAreaName
                  );
        
   STDMETHODIMP GetQualifiedSourceName(
                  /* [in] */                    LPCWSTR           szSourceName,
                  /* [string][out] */           LPWSTR         *  pszQualifiedSourceName
                  );

   ///////////////////////////////////////////////////////////////////////////

// Impmementation
protected:
   EventArea* m_pCurrentArea;
};
//DOM-IGNORE-END

#endif // __EventAreaBrowser_H
