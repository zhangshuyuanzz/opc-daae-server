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
#include "DaEnumItemAttributes.h"
#include "UtilityFuncs.h"
#include "Logger.h"

//--------------------------------------------------------------------
// DaEnumItemAttributes   Constructor
//--------------------------------------------------------------------
DaEnumItemAttributes::DaEnumItemAttributes(
         LPUNKNOWN  pUnkRef,  // use for reference counting (the 'parent')
         ULONG      cUnk,     // number of LPUNKNOWNs in prgUnk
         OPCITEMATTRIBUTES *prgItemAttr,   // ptr to the array to enumerate 
                              // (we will make a local copy).
         IMalloc   *pmem
         )
{
   m_cRef    = 0;
   m_pUnkRef = pUnkRef;
   m_iCur    = 0;
   m_cUnk    = cUnk;
   m_pmem    = pmem;

   if ( cUnk == 0 ) {
      m_prgItemAttr = NULL;
   } else {
      m_prgItemAttr = new OPCITEMATTRIBUTES[ cUnk ];
   }
   if ( m_prgItemAttr != NULL ) {

      OPCITEMATTRIBUTES* pDst;
      OPCITEMATTRIBUTES* pSrc;

      for (ULONG i=0; i<cUnk; i++) {   // copy all attribute members
         pDst = &m_prgItemAttr[i];
         pSrc = &prgItemAttr[i];

         memcpy( pDst, pSrc, sizeof (OPCITEMATTRIBUTES) );

         if (pSrc->szAccessPath) {
            pDst->szAccessPath = WSTRClone( pSrc->szAccessPath, NULL );
         }

         if (pSrc->szItemID) {
            pDst->szItemID = WSTRClone( pSrc->szItemID, NULL );
         }
                                       // Copy EU info
         VariantInit( &pDst->vEUInfo );
         VariantCopy( &pDst->vEUInfo, &pSrc->vEUInfo );
      }
   }

}




//--------------------------------------------------------------------
// IOPCEnumUnknown   Destructor
//--------------------------------------------------------------------
DaEnumItemAttributes::~DaEnumItemAttributes(void)
{
   
   if (m_prgItemAttr == NULL) {
      return;
   }

   OPCITEMATTRIBUTES* pAttr;

   for (ULONG i=0; i<m_cUnk; i++) {    // release all attribute members
      pAttr = &m_prgItemAttr[i];

      if (pAttr->szAccessPath) {
         WSTRFree( pAttr->szAccessPath, NULL );
      }
      if (pAttr->szItemID) {
         WSTRFree( pAttr->szItemID, NULL );
      };
      if (pAttr->pBlob) {
         delete pAttr->pBlob;
      }
      VariantClear( &pAttr->vEUInfo );
   }

   delete m_prgItemAttr;
}






//--------------------------------------------------------------------
//  IOPCEnumUnknown::QueryInterface
//--------------------------------------------------------------------
STDMETHODIMP DaEnumItemAttributes::QueryInterface(
                  REFIID riid,
                  LPVOID *ppv)
{
         *ppv=NULL;

                // Enumerators are separate objects with their own
                // QueryInterface behavior.
   if( IID_IUnknown==riid || IID_IEnumOPCItemAttributes==riid ) {
      *ppv=(LPVOID)this;
   }
   if( *ppv ) {
      ((LPUNKNOWN)*ppv)->AddRef();
      return S_OK;
   }

   return E_NOINTERFACE;
}


//--------------------------------------------------------------------
//  IOPCEnumUnknown::AddRef
//--------------------------------------------------------------------
STDMETHODIMP_(ULONG) DaEnumItemAttributes::AddRef(void)
{
   // Addref this object and also the 'parent' if any
   //
   ++m_cRef;
   if(m_pUnkRef != NULL) 
      m_pUnkRef->AddRef();
   return m_cRef;
}



//--------------------------------------------------------------------
//  IOPCEnumUnknown::Release
//--------------------------------------------------------------------
STDMETHODIMP_(ULONG) DaEnumItemAttributes::Release(void)
{
   // Release this object and also the 'parent' if any
   //
   if(m_pUnkRef != NULL) 
      m_pUnkRef->Release();

   if (0L!=--m_cRef)
      return m_cRef;

   delete this;
   return 0;
}






//--------------------------------------------------------------------
//  IOPCEnumUnknown::Next
//
//  Returns the next element in the enumeration.
//  Return Value:
//  HRESULT       S_OK if successful, S_FALSE otherwise,
//--------------------------------------------------------------------
STDMETHODIMP DaEnumItemAttributes::Next( 
       /* [in] */ ULONG celt,
       /* [size_is][size_is][out] */ OPCITEMATTRIBUTES __RPC_FAR *__RPC_FAR *ppItemAttr,
       /* [out] */ ULONG __RPC_FAR *pceltFetched
)
//STDMETHODIMP DaEnumItemAttributes::Next(
//         ULONG      cRequested,  // max number of LPUNKNOWNs to return 
//         IUnknown **ppUnk,       // where to store the pointer
//         ULONG     *pActual )    // where to store how many we actually enumerated.
{
      ULONG                  cReturn;
      ULONG                  maxcount;
      ULONG                LeftToRead;
      OPCITEMATTRIBUTES    *IASource, *IADest;

   LOGFMTI( "IEnumOPCItemAttributes::Next %lu item(s)", celt ); 

   if ( ppItemAttr == NULL ) {
      LOGFMTE( "Next() failed with invalid argument(s): ppItemAttr is NULL" );
      return E_INVALIDARG;
   }

   *ppItemAttr = NULL;

   if ( pceltFetched == NULL ) {
      LOGFMTE( "Next() failed with invalid argument(s): pceltFetched is NULL" );
      return E_INVALIDARG;
   }

   *pceltFetched = 0L;      // default

   if ( celt == 0 ) {
      LOGFMTI( "%lu item(s) fetched", *pceltFetched ); 
      return S_OK;
   }

   maxcount = celt;
   cReturn = 0L;


                  // If this enumerator is empty - return FALSE
   if (NULL==m_prgItemAttr) {
      LOGFMTE( "Next() failed with eror: The enumerator is empty" );
      return S_FALSE;
   }

   //               // If user passed null for count of items returned
   //               // Then he is only allowed to ask for 1 item
   //if (NULL == pceltFetched)   {
   //   if (1L != celt)    return E_POINTER;
   //}

   //               // If we are at end of list return FALSE
   //if (m_iCur >= m_cUnk) {
   //   return S_FALSE;
   //}

   
   LeftToRead = m_cUnk - m_iCur;

   if ( LeftToRead == 0 ) {
      LOGFMTE( "Next() failed with eror: There are no more items to read" );
      return S_FALSE;
   }

   if ( LeftToRead < celt ) {
         // requested more than available
      celt = LeftToRead;
   }

   *ppItemAttr = (OPCITEMATTRIBUTES *) pIMalloc->Alloc( celt * sizeof(OPCITEMATTRIBUTES) );

                  // Return as many as we have left in list up to request count
   while ( (m_iCur < m_cUnk) && (celt > 0) ) {

      IASource = &(m_prgItemAttr[m_iCur]);
      IADest = &((*ppItemAttr)[cReturn]);

         // copy the structure
      memcpy( IADest, IASource, sizeof(OPCITEMATTRIBUTES) );
            //IADest->pBlob = NULL;
            //IADest->szItemID = NULL;
         // override fields needing pIMalloc 
      IADest->szItemID = WSTRClone( IASource->szItemID , pIMalloc );
      if ( IASource->szAccessPath != NULL ) {
         IADest->szAccessPath = WSTRClone( IASource->szAccessPath , pIMalloc );
      } 
      if ( IASource->pBlob != NULL ) {
         IADest->pBlob = (BYTE *)pIMalloc->Alloc( IASource->dwBlobSize );
         memcpy( IADest->pBlob, IASource->pBlob, IASource->dwBlobSize );
            // note that IASource->dwBlobSize == IADest->dwBlobSize
      }
      VariantInit( &(IADest->vEUInfo) );
      VariantCopy( &(IADest->vEUInfo), &(IASource->vEUInfo) );

      m_iCur++;                        // And move on to the next one
      cReturn++;
      celt--;
   }

   if (NULL != pceltFetched) {    
      *pceltFetched = cReturn;
   }
   LOGFMTI( "%lu item(s) fetched", *pceltFetched ); 

   if (cReturn == maxcount)   
      return S_OK;
   else                       
      return S_FALSE;
}



//--------------------------------------------------------------------
// Skips the next n elements in the enumeration.
//
// Parameters:
//  cSkip         ULONG number of elements to skip.
//
// Return Value:
//  HRESULT     S_OK if successful, 
//            S_FALSE if we could not skip the requested number.
//--------------------------------------------------------------------
STDMETHODIMP DaEnumItemAttributes::Skip( 
       /* [in] */ ULONG celt
)
{
   LOGFMTI( "NOT IMPLEMENTED: IEnumOPCItemAttributes::Skip %lu item(s)", celt ); 
   return E_NOTIMPL;
}



//--------------------------------------------------------------------
// Resets the current element index in the enumeration to zero.
//--------------------------------------------------------------------
STDMETHODIMP DaEnumItemAttributes::Reset( void )
{
   LOGFMTI( "IEnumOPCItemAttributes::Reset" ); 
   m_iCur=0;
   return S_OK;
}



//--------------------------------------------------------------------
// Returns another IEnumUnknown with the same state as ourselves.
//
// Parameters:
//  ppEnum        LPENUMUNKNOWN * in which to return the new object.
//--------------------------------------------------------------------
STDMETHODIMP DaEnumItemAttributes::Clone( 
       /* [out] */ IEnumOPCItemAttributes __RPC_FAR *__RPC_FAR *ppEnumItemAttributes
)
{
   LOGFMTI( "NOT IMPLEMENTED: IEnumOPCItemAttributes::Clone" ); 
   return E_NOTIMPL;
}

//DOM-IGNORE-END




