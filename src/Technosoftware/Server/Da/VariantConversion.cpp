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

//-----------------------------------------------------------------------
// INLCUDE
//-----------------------------------------------------------------------
#include "stdafx.h"
#include <comdef.h>                             // for _bstr_t
#include "variantconversion.h"

//=========================================================================
// VariantFromVariantArrayElement                                  INTERNAL
// ------------------------------
//    Creates a Variant from the element of an Array-Variant.
//    The data pointer must be locked outside of this function.
//=========================================================================
static HRESULT VariantFromVariantArrayElement(
            /* [out] */    LPVARIANT pvDest,
            /* [in]  */    VARTYPE vtRequested,
            /* [in]  */    const void HUGEP* pData,
            /* [in]  */    DWORD dwNdx,
            /* [in]  */    VARTYPE vtElement )
{
   HRESULT  hr = S_OK;
   VARIANT  vTmpByRef;

   VariantInit( &vTmpByRef );

   V_VT( &vTmpByRef ) = vtElement | VT_BYREF;
   switch (vtElement) {
      case VT_I1:       V_I1REF( &vTmpByRef )      = &((CHAR*)pData)[dwNdx];           break;
      case VT_UI1:      V_UI1REF( &vTmpByRef )     = &((BYTE*)pData)[dwNdx];           break;
      case VT_I2:       V_I2REF( &vTmpByRef )      = &((SHORT*)pData)[dwNdx];          break;
      case VT_UI2:      V_UI2REF( &vTmpByRef )     = &((USHORT*)pData)[dwNdx];         break;
      case VT_I4:       V_I4REF( &vTmpByRef )      = &((LONG*)pData)[dwNdx];           break;
      case VT_UI4:      V_UI4REF( &vTmpByRef )     = &((ULONG*)pData)[dwNdx];          break;
      case VT_R4:       V_R4REF( &vTmpByRef )      = &((FLOAT*)pData)[dwNdx];          break;
      case VT_R8:       V_R8REF( &vTmpByRef )      = &((DOUBLE*)pData)[dwNdx];         break;
      case VT_CY:       V_CYREF( &vTmpByRef )      = &((CY*)pData)[dwNdx];             break;
      case VT_DATE:     V_DATEREF( &vTmpByRef )    = &((DATE*)pData)[dwNdx];           break;
      case VT_BSTR:     V_BSTRREF( &vTmpByRef )    = &((BSTR*)pData)[dwNdx];           break;
      case VT_BOOL:     V_BOOLREF( &vTmpByRef )    = &((VARIANT_BOOL*)pData)[dwNdx];   break;
      case VT_VARIANT:  V_VARIANTREF( &vTmpByRef ) = &((VARIANT*)pData)[dwNdx];        break;
      default:          hr = OPC_E_BADTYPE;                                            break;
   }

   if (SUCCEEDED( hr )) {
      hr = VariantChangeType( pvDest, &vTmpByRef, 0, vtRequested );
      if ((V_VT( pvDest ) == VT_EMPTY) && SUCCEEDED( hr )) {
         hr = OPC_E_BADTYPE;                    // It may happen that VariantChangeType returns SUCCEEDED,
      }                                         // but the conversion failed.
   }

   VariantClear( &vTmpByRef );
   return hr;

} // VariantFromVariantArrayElement



//=========================================================================
// PutVariantArrayElementFromVariant                               INTERNAL
// ---------------------------------
//    Stores the value of a Variant as data element at given location
//    in the specified Array-Variant.
//=========================================================================
static HRESULT PutVariantArrayElementFromVariant(
            /* [out] */    LPVARIANT pvDest,
            /* [in]  */    long* plElementNdx,
            /* [in]  */    const LPVARIANT pvElement )
{
   HRESULT  hr = S_OK;
   switch (V_VT( pvElement )) {
      case VT_I1:       hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx, &V_I1( pvElement ) );    break;
      case VT_UI1:      hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx, &V_UI1( pvElement ) );   break;
      case VT_I2:       hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx, &V_I2( pvElement ) );    break;
      case VT_UI2:      hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx, &V_UI2( pvElement ) );   break;
      case VT_I4:       hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx, &V_I4( pvElement ) );    break;
      case VT_UI4:      hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx, &V_UI4( pvElement ) );   break;
      case VT_R4:       hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx, &V_R4( pvElement ) );    break;
      case VT_R8:       hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx, &V_R8( pvElement ) );    break;
      case VT_CY:       hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx, &V_CY( pvElement ) );    break;
      case VT_DATE:     hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx, &V_DATE( pvElement ) );  break;
      case VT_BSTR:     hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx,  V_BSTR( pvElement ) );  break;
      case VT_BOOL:     hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx, &V_BOOL( pvElement ) );  break;
      case VT_VARIANT:  hr = SafeArrayPutElement( V_ARRAY( pvDest ), plElementNdx, pvElement );             break;
      default:          hr = OPC_E_BADTYPE;  break;
   }
   return hr;

} // PutVariantArrayElementFromVariant



//=========================================================================
// VariantFromVariant
// ------------------
//    For a description see the comments in the module header.
//=========================================================================
HRESULT VariantFromVariant(
            /* [out] */    LPVARIANT pvDest,
            /* [in]  */    VARTYPE vtRequested,
            /* [in]  */    const LPVARIANT pvSrc )
{
   HRESULT hr = S_OK;

   if (V_VT( pvSrc ) == vtRequested) {
                                                // The canonical and requested data types are identical.
      hr = VariantCopy( pvDest, pvSrc );        // Get current value.
   }
   else if (vtRequested & VT_ARRAY) {

      if (V_ISARRAY( pvSrc )) {
         //
         // Convert from an Array to an Array with different basic types
         //

         VARTYPE     vtRequestedArrayType = vtRequested & VT_TYPEMASK;
         void HUGEP* pData = NULL;
         SAFEARRAY*  psa = V_ARRAY( pvSrc );

         hr = SafeArrayAccessData( psa, &pData );
         if (SUCCEEDED( hr )) {

            //  Create the destination Array-Variant
            V_VT( pvDest )    = vtRequested;
            V_ARRAY( pvDest ) = SafeArrayCreate( vtRequestedArrayType, 1, psa->rgsabound );
            if (!V_ARRAY( pvDest )) {
               hr = E_OUTOFMEMORY;
            }
            else {
               VARIANT  vTmpDst;
               DWORD    dwLBound = psa->rgsabound[0].lLbound;
               DWORD    dwUBound = dwLBound + psa->rgsabound[0].cElements;

               // Convert all elements of the Array-Variant
               for (DWORD i=dwLBound; i < dwUBound; i++) {
                  VariantInit( &vTmpDst );

                  hr = VariantFromVariantArrayElement(   &vTmpDst, vtRequestedArrayType,
                                                         pData, i,
                                                         V_VT( pvSrc ) & VT_TYPEMASK );   // Variant Type of the Source Element

                  if (SUCCEEDED( hr )) {
                     hr = PutVariantArrayElementFromVariant( pvDest,
                                                             (long*)&i, &vTmpDst );
                  }
                  VariantClear( &vTmpDst );
                  if (FAILED( hr )) break;
               }
            }
            HRESULT hrTmp = SafeArrayUnaccessData( psa );
            if (SUCCEEDED( hr )) hr = hrTmp;
         }
      }
      else {
         //
         // Convert from a Simple Type an Array with one element of a requested type
         //

         VARTYPE  vtRequestedArrayType = vtRequested & VT_TYPEMASK;

         //  Create the destination Array-Variant
         SAFEARRAYBOUND rgs;
         rgs.cElements  = 1;
         rgs.lLbound    = 0;

         V_VT( pvDest )    = vtRequested;
         V_ARRAY( pvDest ) = SafeArrayCreate( vtRequestedArrayType, 1, &rgs );
         if (!V_ARRAY( pvDest )) {
            hr = E_OUTOFMEMORY;
         }
         else {
            VARIANT  vTmpDst;
            VariantInit( &vTmpDst );

            hr = VariantFromVariant( &vTmpDst, vtRequestedArrayType,
                                     pvSrc );

            if (SUCCEEDED( hr )) {
               long i = 0;
               hr = PutVariantArrayElementFromVariant( pvDest,
                                                       &i, &vTmpDst );
            }
            VariantClear( &vTmpDst );
         }
      }
   }
   else if ((vtRequested & VT_BSTR) && V_ISARRAY( pvSrc )) {
      //
      // Convert an Array-Variant to a BSTR-Variant
      //
      void HUGEP* pData = NULL;
      SAFEARRAY*  psa = V_ARRAY( pvSrc );

      hr = SafeArrayAccessData( psa, &pData );
      if (SUCCEEDED( hr )) {
            _bstr_t  bstr;
            VARIANT  vTmp;

            DWORD dwLBound = psa->rgsabound[0].lLbound;
            DWORD dwUBound = dwLBound + psa->rgsabound[0].cElements;

            for (DWORD i=dwLBound; i < dwUBound; i++) {
               VariantInit( &vTmp );

               hr = VariantFromVariantArrayElement(   &vTmp, VT_BSTR,
                                                      pData, i,
                                                      V_VT( pvSrc ) & VT_TYPEMASK );      // Variant Type of the Source Element
               if (SUCCEEDED( hr )) {
                  try {
                     bstr += V_BSTR( &vTmp );
                     if (i < dwUBound -1) {     // If not last element value
                        bstr += "; ";           // then insert a delimiter character
                     }
                  }
                  catch (_com_error &e) {
                     hr = e.Error();
                  }
               }
               VariantClear( &vTmp );
               if (FAILED( hr )) break;
            }
            if (SUCCEEDED( hr )) {
               try {
                  V_VT( pvDest )    = VT_BSTR;
                  V_BSTR( pvDest )  = bstr.copy();
               }
               catch (_com_error &e) {
                  hr = e.Error();
               }
            }

            HRESULT hrTmp = SafeArrayUnaccessData( psa );
            if (SUCCEEDED( hr )) hr = hrTmp;
      }
   }
   else {
      hr = VariantChangeType( pvDest, pvSrc, 0, vtRequested );
      if ((V_VT( pvDest ) == VT_EMPTY) && SUCCEEDED( hr )) {
         hr = OPC_E_BADTYPE;                    // It may happen that VariantChangeType returns SUCCEEDED,
      }                                         // but the conversion failed.
   }

   if (FAILED( hr )) {                          // Cannot get new value
      VariantClear( pvDest );                   // Initialize result value

      switch (hr) {                             // Convert DISP_E_xxx error codes to OPC error codes
         // deactivated to avoid a problem in the CTT 2.00.7.1120
         //case DISP_E_TYPEMISMATCH:         
         case DISP_E_BADVARTYPE:    hr = OPC_E_BADTYPE;  break;
         case DISP_E_OVERFLOW:      hr = OPC_E_RANGE;    break;
      }
   }
   return hr;

} // VariantFromVariant

//DOM-IGNORE-END
