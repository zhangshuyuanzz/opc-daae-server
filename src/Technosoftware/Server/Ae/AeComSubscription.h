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

#ifndef __OPCEVENTSUBSCRIPTION_H
#define __OPCEVENTSUBSCRIPTION_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "AeComSubscriptionManager.h"

EXTERN_C CLSID CLSID_IOPCEventSubscription;

class AeComSubscription;


//-----------------------------------------------------------------------
// TYPEDEF
//-----------------------------------------------------------------------
typedef IConnectionPointImpl
                     <  AeComSubscription,
                        &IID_IOPCEventSink,
                        CComUnkArray<1>         // Note: There must be a space after CComUnkArray<1> !!!
                     > IOPCEventSubscriptionConnectionPointImpl;


//-----------------------------------------------------------------------
// CLASS
//-----------------------------------------------------------------------
class ATL_NO_VTABLE AeComSubscription :
   public CComObjectRootEx<CComMultiThreadModel>,
   public CComCoClass<AeComSubscription, &CLSID_IOPCEventSubscription>,
   public AeComSubscriptionManager,
   public IConnectionPointContainerImpl<AeComSubscription>,
   public IOPCEventSubscriptionConnectionPointImpl
{
public:
   ///////////////////////////////////////////////////////////////////////////
   // ATL macros
   ///////////////////////////////////////////////////////////////////////////
                                                // entry points for COM
   BEGIN_COM_MAP(AeComSubscription)
      COM_INTERFACE_ENTRY(IOPCEventSubscriptionMgt)
      COM_INTERFACE_ENTRY(IConnectionPointContainer)
   END_COM_MAP()

   BEGIN_CONNECTION_POINT_MAP(AeComSubscription)
      CONNECTION_POINT_ENTRY(IID_IOPCEventSink)
   END_CONNECTION_POINT_MAP()

   DECLARE_NO_REGISTRY()                        // avoid default ATL registration
   DECLARE_NOT_AGGREGATABLE(AeComSubscription)

   ///////////////////////////////////////////////////////////////////////////

// Construction / Destruction
public:
   AeComSubscription()  { m_dwCookieGITEventSink = 0; }
   ~AeComSubscription() {}

// Operations
public:
                                                // Implemented in module AeComSubscriptionManager.cpp
   HRESULT FireOnEvent( DWORD dwNumOfEvents, ONEVENTSTRUCT* pEvent,
                        BOOL fRefresh = FALSE, BOOL fLastRefresh = FALSE );

// Overridables      
public:
                                                // IConnectionPoint funtions
   STDMETHODIMP Advise( IUnknown* pUnkSink, DWORD* pdwCookie );
   STDMETHODIMP Unadvise( DWORD dwCookie );
                                                // Cleanup AeComSubscriptionManager instance
   void  FinalRelease() { CleanupSubscription(); }

// Impmementation
protected:
   // IOPCEventSink interface access
   // ---
   HRESULT GetEventSinkInterface( IOPCEventSink** ppSink );

      // Cookie used by Global Interface Table related functions to access
      // the IOPCEventSink interface from other apartments.
      // This member is used only if the Global Interface Table is used for
      // marshaling of callback interfaces. Therefore the flag
      // _Module.m_fMarshalCallbacks is TRUE.
   private:
   DWORD m_dwCookieGITEventSink;
   // ---
};
//DOM-IGNORE-END

#endif // __OPCEVENTSUBSCRIPTION_H
