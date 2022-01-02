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

#ifndef __OPCEVENTSERVER_H
#define __OPCEVENTSERVER_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "OpcCommon.h"
#include "AeComBaseServer.h"
#include "ShutdownImpl.h"

EXTERN_C CLSID CLSID_IOPCEventServer;


//-----------------------------------------------------------------------
// CLASS
//-----------------------------------------------------------------------
class ATL_NO_VTABLE AeComServer :
   public CComObjectRootEx<CComMultiThreadModel>,
   public CComCoClass<AeComServer, &CLSID_IOPCEventServer>,
   public OpcCommon,
   public AeComBaseServer,
   public IConnectionPointContainerImpl<AeComServer>,
   public IOPCShutdownConnectionPointImpl<AeComServer>
{
public:
   ///////////////////////////////////////////////////////////////////////////
   // ATL macros
   ///////////////////////////////////////////////////////////////////////////
                                                // Entry points for COM
   BEGIN_COM_MAP(AeComServer)
      COM_INTERFACE_ENTRY(IOPCCommon)
      COM_INTERFACE_ENTRY(IOPCEventServer)
      COM_INTERFACE_ENTRY(IConnectionPointContainer)
      COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
   END_COM_MAP()

   BEGIN_CONNECTION_POINT_MAP(AeComServer)
      CONNECTION_POINT_ENTRY(IID_IOPCShutdown)
   END_CONNECTION_POINT_MAP()

   DECLARE_NO_REGISTRY()                        // Avoid default ATL registration
   DECLARE_GET_CONTROLLING_UNKNOWN()            // Required by ATL for virtual func GetControllingUnknown
   DECLARE_PROTECT_FINAL_CONSTRUCT()            // Required for Aggregation 

   ///////////////////////////////////////////////////////////////////////////

// Construction / Destruction
public:
   AeComServer() {}
   ~AeComServer() {}

// Attributes
public:
   CComPtr<IUnknown> m_pUnkMarshaler;

// Operations
public:
   inline void FireShutdownRequest( LPCWSTR szReason )
   {
      IOPCShutdownConnectionPointImpl<AeComServer>::FireShutdownRequest( szReason );
   }

// Overrides
public:
   HRESULT  FinalConstruct()
   {
      HRESULT hres = CoCreateFreeThreadedMarshaler(
                                 GetControllingUnknown(), 
                                 &m_pUnkMarshaler.p);
      if (SUCCEEDED( hres )) {
         hres = OpcCommon::Create();
      }
      if (SUCCEEDED( hres )) {
         hres = AeComBaseServer::Create();
      }
      return hres;
   }

   HRESULT  FinalRelease()
   {
      m_pUnkMarshaler.Release();
      return S_OK;
   }

// Impmementation
protected:
};
//DOM-IGNORE-END

#endif // __OPCEVENTSERVER_H
