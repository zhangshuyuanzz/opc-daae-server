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

#ifndef ENUMCLASS_H
#define ENUMCLASS_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

typedef  CComObject< CComEnum <  IEnumString, &IID_IEnumString,
                                 BSTR, _Copy<LPOLESTR>
                              >  > COpcComEnumString;         

typedef  CComObject< CComEnum <  IEnumUnknown, &IID_IEnumUnknown,
                                 IUnknown *, _CopyInterface< IUnknown >
                              >  > COpcComEnumUnknown;


typedef  CComObject < CComEnum<  IEnumVARIANT, &IID_IEnumVARIANT,
                                 VARIANT, _Copy< VARIANT >
                              > >  COpcComEnumVARIANT;

//-----------------------------------------------------------------------
// OPC specific implementation of an IEnumUnknown Interface for the Server
//-----------------------------------------------------------------------
class IOpcEnumVariant : public IEnumVARIANT
{
  public:
   IOpcEnumVariant( 
         short       EnumType, 
         ULONG       nElements, 
         void        *pElements, 
         LPUNKNOWN   pUnkRef, 
         IMalloc     *pmem
      );
   IOpcEnumVariant(
         short       EnumType,
         ULONG       nElements,     // number of Variants in pVariants
         VARIANT     *pElements,    // ptr to the array to enumerate 
         LPUNKNOWN   pUnkRef,       // use for reference counting (the 'parent')
         IMalloc     *pmem  
      );

   ~IOpcEnumVariant( void);

                              // the IUnknown Methods
   STDMETHODIMP         QueryInterface( REFIID iid, LPVOID* ppInterface);
   STDMETHODIMP_(ULONG) AddRef( void);
   STDMETHODIMP_(ULONG) Release( void);

                              // the IEnumUnknown Methods
   STDMETHODIMP         Next (
      ULONG    Requested,
      VARIANT  *pElements,
      ULONG    *pFetched
      );
        
   STDMETHODIMP         Skip (
      ULONG nElements
      );
        
   STDMETHODIMP         Reset(
      void
      );
        
   STDMETHODIMP         Clone( 
      IEnumVARIANT **ppEnum
      );

                           // Member Variables
  private:

    void ConvertToVariant(
         short       EnumType,
         ULONG       Elem,       // number of LPUNKNOWNs in prgUnk
         void        *pElem      // array to enumerate ( type is LPWSTR or IUnknown* )
      );

   short          m_EnumType;    // Enum BSTR or Enum IUnknown (IDispatch)
   ULONG          m_nCurElem;    // next/current position w/in m_rgContents
   ULONG          m_nElements;   // total count of variants
   VARIANT     *  m_pElements;   // contents of collection

   ULONG          m_cRef;        // Object reference count
   LPUNKNOWN      m_pUnkRef;     // LPUNKNOWN for ref counting
   IMalloc     *  m_pmem;        // memory allocator to use

};

//DOM-IGNORE-END



#endif
