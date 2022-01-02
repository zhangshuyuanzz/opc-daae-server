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

#include "stdafx.h"
#include "enumclass.h"

///////////////////////////////////////////////////////////////////////////////////////// 
///////////////////////////////////////////////////////////////////////////////////////// 
// IEnumVariant /////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////// 
///////////////////////////////////////////////////////////////////////////////////////// 


//--------------------------------------------------------------------
// IOpcEnumVariant   Constructor
//--------------------------------------------------------------------
IOpcEnumVariant::IOpcEnumVariant(
         short       EnumType,
         ULONG       nElements,     // number of LPUNKNOWNs in prgUnk
         void        *pElements,    // ptr to the array to enumerate 
                                    // (we will make a local copy).
         LPUNKNOWN   pUnkRef,       // use for reference counting (the 'parent')
         IMalloc     *pmem
         )
{

   m_cRef         = 0;
   m_pUnkRef      = pUnkRef;
   m_pmem         = pmem;

   m_nCurElem     = 0;
   m_nElements    = nElements;
   m_EnumType     = EnumType;
   
   if (nElements == 0) {
      m_pElements = NULL;
   } else {
      m_pElements = new VARIANT[ (UINT)nElements ] ;
   }

   if( m_pElements ) {
      ConvertToVariant( EnumType, nElements, (LPUNKNOWN *)pElements );
   }


   if ( EnumType == VT_UNKNOWN ) {
            // have to AddRef the objects so that they 
            // cannot be deleted while using them
         ULONG i;
      i = 0;
      while ( i < nElements ) {
         V_UNKNOWN( &m_pElements[i] )->AddRef();
         i ++;
      }
   }

}


//--------------------------------------------------------------------
// IOpcEnumVariant   Constructor ( for cloning )
//--------------------------------------------------------------------
IOpcEnumVariant::IOpcEnumVariant(
         short       EnumType,
         ULONG       nElements,     // number of Variants in pElements
         VARIANT     *pElements,    // ptr to the array to enumerate 
                                    // (we will make a local copy).
         LPUNKNOWN   pUnkRef,       // use for reference counting (the 'parent')
         IMalloc     *pmem
         )
{
      ULONG    i;

   m_cRef         = 0;
   m_pUnkRef      = pUnkRef;
   m_pmem         = pmem;

   m_nCurElem     = 0;
   m_nElements    = nElements;
   m_EnumType     = EnumType;
   
   m_pElements = new VARIANT[ (UINT)nElements ] ;

   if ( m_pElements != NULL ) {
      i = 0;
      while ( i < nElements ) {
         VariantInit( &(m_pElements[i]) );
         VariantCopy( &(m_pElements[i]), &(pElements[i]) );
         i ++;
      }
   }

}


//--------------------------------------------------------------------
// IOpcEnumVariant   Destructor
//--------------------------------------------------------------------
IOpcEnumVariant::~IOpcEnumVariant(void)
{
         ULONG   i;

   if (m_pElements != NULL)   {
         // free all the Variant array elements
      for (i=0; i < m_nElements; i++)   {
            // clear the variant (also releases a IUnknown interface! )
         VariantClear( &( m_pElements[i] ) );   
      }
         // free the variant array
      delete m_pElements;
   }

   return;
}


//--------------------------------------------------------------------
// IOpcEnumVariant::ConvertToVariant
//
// ConvertToVariant   
//    converts the IUnknown* or LPWSTR to a VARIANT array
//--------------------------------------------------------------------
void IOpcEnumVariant::ConvertToVariant(
         short       EnumType,
         ULONG       NElem,     // number of LPUNKNOWNs in prgUnk
         void        *pElem     // array to enumerate ( type is LPWSTR or IUnknown* )
   )
{
      ULONG    i;

   i = 0;
   while ( i < NElem )  {
      VariantInit( &m_pElements[i] );
      if ( EnumType == VT_UNKNOWN ) {
            // create VARIANT IUnknown
         V_VT( &m_pElements[i] )       = VT_UNKNOWN; 
         V_UNKNOWN( &m_pElements[i] )  = (LPUNKNOWN)( ( (LPUNKNOWN *) pElem)[i]);

      } else if ( EnumType == VT_BSTR ) {
            // create VARIANT BSTR
         V_VT( &m_pElements[i] )    = VT_BSTR;
         V_BSTR( &m_pElements[i] )  = SysAllocString( ( LPWSTR ) ((LPWSTR *)pElem)[i] );
      }
      i ++;
   }

}



//--------------------------------------------------------------------
//  IOpcEnumVariant::Next
//
//  Returns the next elements in the enumeration.
//  Return Value:
//    S_OK if successful, 
//    S_FALSE if less than requested elemens are available
//--------------------------------------------------------------------
STDMETHODIMP IOpcEnumVariant::Next( 
         ULONG       Requested,       // max number of LPUNKNOWNs to return 
         VARIANT     *pElements,      // the VARIANT ARRAY containing the IUnknown- and BSTR- VARIANTs
         ULONG         *pFetched )      // where to store how many we actually enumerated.
{
      ULONG       nReturn;
      ULONG       maxcount; 
      HRESULT     res;

      //!!!??? pFetched not set!
         
                  // If this enumerator is empty - return FALSE
   if (NULL==m_pElements) {
      for (maxcount=0; maxcount<Requested; maxcount++) {
         VariantInit( &(pElements[maxcount]) );
      }
      *pFetched = 0L;
      return S_FALSE;
   }

     
   maxcount = Requested;

                  // If user passed null for count of items returned
                  // Then he is only allowed to ask for 1 item
   if (NULL == pFetched)   {
         // !!!???!!!  what is this?
      if (1L != Requested) {
         return E_POINTER;
      }
   } else {
      *pFetched = 0L; // default
   }

   nReturn = 0;
                  // If we are at end of list return FALSE
   if (m_nCurElem >= m_nElements) {
      return S_FALSE;
   }

                  // Return as many as we have left in list up to request count
   while ( (m_nCurElem < m_nElements) && (Requested > 0) ) {
      VariantInit( &(pElements[nReturn]) );
      res = VariantCopy( &(pElements[nReturn]) , &(m_pElements[m_nCurElem]) );

      m_nCurElem++;                        // And move on to the next one
      nReturn++;
      Requested--;
   }

   if ( NULL != pFetched ) {
      *pFetched = nReturn;
   }

   if (nReturn == maxcount) {
      return S_OK;
   } else {
      return S_FALSE;
   }
}



//--------------------------------------------------------------------
// IOpcEnumVariant::Skip
//
// Skips the next n elements in the enumeration.
//
// Parameters:
//  cSkip         ULONG number of elements to skip.
//
// Return Value:
//  HRESULT     S_OK if successful, 
//            S_FALSE if we could not skip the requested number.
//--------------------------------------------------------------------
STDMETHODIMP IOpcEnumVariant::Skip(ULONG cSkip)
{
   if ( ( ( m_nCurElem + cSkip ) >= m_nElements) || ( NULL == m_pElements ) )
      return S_FALSE;

   m_nCurElem += cSkip;
   return S_OK;
}


//--------------------------------------------------------------------
// IOpcEnumVariant::Reset
//
// Resets the current element index in the enumeration to zero.
//--------------------------------------------------------------------
STDMETHODIMP IOpcEnumVariant::Reset(void)
{
   m_nCurElem = 0;
   return S_OK;
}


//--------------------------------------------------------------------
// IOpcEnumVariant::Clone
//
// Returns another IEnumVARIANT with the same state as ourselves.
//
// Parameters:
//  ppEnum        IEnumVARIANT * in which to return the new object.
//--------------------------------------------------------------------
STDMETHODIMP IOpcEnumVariant::Clone(IEnumVARIANT **ppEnum)
{
      IOpcEnumVariant   *pNew;

   *ppEnum=NULL;

               //Create the clone
   pNew= new IOpcEnumVariant( m_EnumType, m_nElements, m_pElements, m_pUnkRef, m_pmem
                              );

   if (NULL==pNew)     return E_OUTOFMEMORY;

   pNew->AddRef();

            // Set the 'state' of the clone to match the state if this
   pNew->m_nElements = m_nElements;

   *ppEnum = pNew;

   return S_OK;
}



//--------------------------------------------------------------------
// IOpcEnumVariant::QueryInterface
//--------------------------------------------------------------------
STDMETHODIMP IOpcEnumVariant::QueryInterface(
                  REFIID riid,
                  LPVOID *ppv)
{
         *ppv=NULL;

                // Enumerators are separate objects with their own
                // QueryInterface behavior.
   if( IID_IUnknown==riid || IID_IEnumVARIANT==riid ) {
      *ppv=(LPVOID)this;
   }
   if( *ppv ) {
      ((LPUNKNOWN)*ppv)->AddRef();
      return S_OK;
   }

   return E_NOINTERFACE;
}


//--------------------------------------------------------------------
// IOpcEnumVariant::AddRef
//--------------------------------------------------------------------
STDMETHODIMP_(ULONG) IOpcEnumVariant::AddRef(void)
{
   // Addref this object and also the 'parent' if any
   ++m_cRef;
   if( m_pUnkRef != NULL )   m_pUnkRef->AddRef();
   return m_cRef;
}



//--------------------------------------------------------------------
// IOpcEnumVariant::Release
//--------------------------------------------------------------------
STDMETHODIMP_(ULONG) IOpcEnumVariant::Release(void)
{
   // Release this object and also the 'parent' if any
   if( m_pUnkRef != NULL )   m_pUnkRef->Release();
   if( 0L != --m_cRef )      return m_cRef;
   delete this;
   return 0;
}

//DOM-IGNORE-END



