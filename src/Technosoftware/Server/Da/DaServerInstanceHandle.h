/*
 * Copyright (c) 2020 Technosoftware GmbH. All rights reserved
 * Web: https://technosoftware.com 
 * 
 * The source code in this file is covered under a dual-license scenario:
 *   - Owner of a purchased license: RPL 1.5
 *   - GPL V3: everybody else
 *
 * RPL license terms accompanied with this source code.
 * See https://technosoftware.com/license/RPLv15License.txt
 *
 * GNU General Public License as published by the Free Software Foundation;
 * version 3 of the License are accompanied with this source code.
 * See https://technosoftware.com/license/GPLv3License.txt
 *
 * This source code is distributed in the hope that it will be useful, but
 * WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY
 * or FITNESS FOR A PARTICULAR PURPOSE.
 */

#ifndef __SERVERINSTANCEHANDLE_H__
#define __SERVERINSTANCEHANDLE_H__

//DOM-IGNORE-BEGIN

   // Class to implement Server Instance Handle as smart pointer
class DaServerInstanceHandle : public IUnknown
{
   // IUnknown implementation
public:
   STDMETHODIMP         QueryInterface( REFIID iid, void** ppvObject )
   {
      *ppvObject = NULL;
      return E_NOTIMPL;
   }

   STDMETHODIMP_(ULONG) AddRef( void )
   {
      return InterlockedIncrement( &m_lRefCount );
   }

   STDMETHODIMP_(ULONG) Release( void )
   {
      if (InterlockedDecrement( &m_lRefCount ) == 0) {
         delete this;
         return 0;
      }
      return m_lRefCount;
   }

   // Constructor
public:
   DaServerInstanceHandle() : m_lRefCount(0) {}

private:
   long  m_lRefCount;
};

typedef CComPtr<DaServerInstanceHandle>  SRVINSTHANDLE;
//DOM-IGNORE-END

#endif //__SERVERINSTANCEHANDLE_H__
