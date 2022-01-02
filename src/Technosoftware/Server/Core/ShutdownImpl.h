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

#ifndef __SHUTDOWNIMPL_H
#define __SHUTDOWNIMPL_H

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
#include "Logger.h"


///////////////////////////////////////////////////////////////////////////
// Class Template IOPCShutdownConnectionPointImpl
///////////////////////////////////////////////////////////////////////////

template <class T>
class ATL_NO_VTABLE IOPCShutdownConnectionPointImpl :
      public IConnectionPointImpl
                  <  T,
                     &IID_IOPCShutdown,
                     CComUnkArray<1>            // Note: There must be a space after CComUnkArray<1> !!!>
                  >
{
public:
   typedef IConnectionPointImpl
                  <  T,
                     &IID_IOPCShutdown,
                     CComUnkArray<1>            // Note: There must be a space after CComUnkArray<1> !!!
                  > _BaseClassCP;

// Construction / Destruction
public:
   IOPCShutdownConnectionPointImpl();
   ~IOPCShutdownConnectionPointImpl();

// Operations
   void FireShutdownRequest( LPCWSTR szReason );

// Overridables
      // Overloaded IConnectionPoint functions
	STDMETHOD(Advise)(IUnknown* pUnkSink, DWORD* pdwCookie);
	STDMETHOD(Unadvise)(DWORD dwCookie);

// Implementation
protected:
   HRESULT GetShutdownInterface( IOPCShutdown** ppShutdown );

private:
      // Cookie used by Global Interface Table related functions to access
      // the IOPCShutdown interface from other apartments.
      // This member is used only if the Global Interface Table is used for
      // marshaling of callback interfaces (the flag
      // _Module.m_fMarshalCallbacks is TRUE).
   DWORD m_dwCookieGITShutdown;
};



///////////////////////////////////////////////////////////////////////////
///////////////////////////// Implementation //////////////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// Constructor
//=========================================================================
template <class T>
IOPCShutdownConnectionPointImpl<T>::IOPCShutdownConnectionPointImpl()
{
   m_dwCookieGITShutdown = 0;
}



//=========================================================================
// Destructor
//=========================================================================
template <class T>
IOPCShutdownConnectionPointImpl<T>::~IOPCShutdownConnectionPointImpl()
{
      if (m_dwCookieGITShutdown) {
         core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( m_dwCookieGITShutdown );
      }
}



///////////////////////////////////////////////////////////////////////////
///////////////////////////// IConnectionPoint ////////////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IConnectionPoint::Advise                                       INTERFACE
// ------------------------
//    Overriden ATL implementation of IConnectionPoint::Advise.
//    This function also registers the client's shutdown sink in the Global
//    Interface Table if required.
//=========================================================================
template <class T>
STDMETHODIMP IOPCShutdownConnectionPointImpl<T>::Advise(IUnknown* pUnkSink, DWORD* pdwCookie)
{
   {
      IID iid;
      GetConnectionInterface( &iid );
      if (IsEqualIID( iid, IID_IOPCShutdown )) {
		  LOGFMTI( "IConnectionPoint::Advise, IID_IOPCShutdown" );
      }
      else {
		  LOGFMTI( "IConnectionPoint::Advise, Unknown IID" );
      }
   }

   T* pT = static_cast<T*>(this);
   pT->Lock();                                  // Lock the connection point list

                                                // Call the base class member
   HRESULT hres = _BaseClassCP::Advise( pUnkSink, pdwCookie );
   if (SUCCEEDED( hres )) {
                                                // Register the callback interface in the global interface table
      hres = core_generic_main.m_pGIT->RegisterInterfaceInGlobal( pUnkSink, IID_IOPCShutdown, &m_dwCookieGITShutdown );
      if (FAILED( hres )) {
         m_dwCookieGITShutdown = 0;
         _BaseClassCP::Unadvise( *pdwCookie );
         pUnkSink->Release();                   // Note :   register increments the refcount
      }                                         //          even if the function failed     
   }
   pT->Unlock();                                // Unlock the connection point list
   return hres;
}



//=========================================================================
// IConnectionPoint::Unadvise                                     INTERFACE
// --------------------------
//    Overriden ATL implementation of IConnectionPoint::Unadvise.
//    This function also removes the client's shutdown sink from the Global
//    Interface Table if required.
//=========================================================================
template <class T>
STDMETHODIMP IOPCShutdownConnectionPointImpl<T>::Unadvise(DWORD dwCookie)
{
	LOGFMTI( "IConnectionPoint::Unadvise, IID_IOPCShutdown" );

   T* pT = static_cast<T*>(this);
   pT->Lock();                                  // Lock the connection point list

   HRESULT hresGIT = S_OK;
   hresGIT = core_generic_main.m_pGIT->RevokeInterfaceFromGlobal( m_dwCookieGITShutdown );
   m_dwCookieGITShutdown = 0;

   HRESULT hres = _BaseClassCP::Unadvise( dwCookie );
   
   pT->Unlock();                                // Unlock the connection point list

   if (FAILED( hres )) {
      return hres;
   }
   if (FAILED( hresGIT )) {
      return hresGIT;
   }
   return hres;
}



///////////////////////////////////////////////////////////////////////////
///////////////////// IOPCShutdownConnectionPointImpl /////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// GetShutdownInterface                                            INTERNAL
// --------------------
//    Gets a pointer to the IOPCShutdown interface from the connection
//    point list or from the Global Interface Table.
//        
//    The GetCallbackInterface() method calls IUnknown::AddRef() on the
//    pointer obtained in the ppShutdown parameter. It is the caller's
//    responsibility to call Release() on this pointer.
//
// Return Code:
//    S_OK                    All succeeded
//    CONNECT_E_NOCONNECTION  There is no connection established
//                            (IConnectionPoint::Advise not called)
//    E_xxx                   Other error occured
//=========================================================================
template <class T>
HRESULT IOPCShutdownConnectionPointImpl<T>::GetShutdownInterface( IOPCShutdown** ppShutdown )
{
   HRESULT hres = CONNECT_E_NOCONNECTION;       // There is no registered callback function

   *ppShutdown = NULL;
   T* pT = static_cast<T*>(this);
   pT->Lock();                                  // Lock the connection point list

   IUnknown** pp = m_vec.begin();               // There can be only one registered shutdown sink.
   if (*pp) {
      hres = core_generic_main.m_pGIT->GetInterfaceFromGlobal( m_dwCookieGITShutdown, IID_IOPCShutdown, (LPVOID*)ppShutdown );
   }

   pT->Unlock();                                // Unlock the connection point list
   return hres;
}



//=========================================================================
// FireShutdownRequest                                               PUBLIC
// -------------------
//    Fires a 'Shutdown Request' to the subscribed client
//=========================================================================
template <class T>
void IOPCShutdownConnectionPointImpl<T>::FireShutdownRequest( LPCWSTR szReason )
{
   IOPCShutdown* pShutdown;
   if (SUCCEEDED( GetShutdownInterface( &pShutdown ) )) {
      pShutdown->ShutdownRequest( szReason ? szReason : L"" );
      pShutdown->Release();                     // All is done with this interface
   }
}

#endif // __SHUTDOWNIMPL_H
