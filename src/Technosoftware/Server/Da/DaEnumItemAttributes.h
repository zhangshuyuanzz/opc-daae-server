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

#ifndef _ENUMOPCITEMATTRIBUTES_H
#define _ENUMOPCITEMATTRIBUTES_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

class DaEnumItemAttributes : public IEnumOPCItemAttributes
{
public:
   DaEnumItemAttributes( LPUNKNOWN, ULONG, OPCITEMATTRIBUTES*, IMalloc*
                           );
   ~DaEnumItemAttributes( void);

                           // the IUnknown Methods
   STDMETHODIMP         QueryInterface( REFIID iid, LPVOID* ppInterface);
   STDMETHODIMP_(ULONG) AddRef( void);
   STDMETHODIMP_(ULONG) Release( void);

   STDMETHODIMP Next( 
       /* [in] */ ULONG celt,
       /* [size_is][size_is][out] */ OPCITEMATTRIBUTES __RPC_FAR *__RPC_FAR *ppItemAttr,
       /* [out] */ ULONG __RPC_FAR *pceltFetched);
   
   STDMETHODIMP Skip( 
       /* [in] */ ULONG celt);
   
   STDMETHODIMP Reset( void);
   
   STDMETHODIMP Clone( 
       /* [out] */ IEnumOPCItemAttributes __RPC_FAR *__RPC_FAR *ppEnumItemAttributes);
        
                           // Member Variables
private:

   ULONG                m_cRef;        // Object reference count
   LPUNKNOWN            m_pUnkRef;     // IUnknown for ref counting
   ULONG                m_iCur;        // Current element
   ULONG                m_cUnk;        // Number of unknowns in use
   OPCITEMATTRIBUTES *  m_prgItemAttr; // Source of OPCITEMATTRIBUTES
   IMalloc           *  m_pmem;        // memory allocator to use
};
//DOM-IGNORE-END

#endif // _ENUMOPCITEMATTRIBUTES_H
