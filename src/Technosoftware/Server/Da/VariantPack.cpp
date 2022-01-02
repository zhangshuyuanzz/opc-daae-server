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
#include "VariantPack.h"

//-----------------------------------------------------------------------
// Returns the number of bytes needed to marshall a variant.
// This can be used to compute the total size required for the stream.
// If -1 is returned some error occured
//-----------------------------------------------------------------------
long SizePackedVariant( VARIANT *vp )
{
      HRESULT     hres;
      SAFEARRAY*  pArr;
      long        LBound, UBound;
      long        len, i;
      VARTYPE     vtBase;

   len = sizeof (VARIANT);
   vtBase = V_VT( vp );
   
   if ((vtBase & VT_BYREF) != 0) {
      return -1;                                            // Don't accept by ref 

   } else if ((vtBase & VT_ARRAY) != 0) {
                                                            // This is a safe array !
      len += sizeof (SAFEARRAY);

      vtBase = vtBase - VT_ARRAY;
   
      pArr = V_ARRAY( vp );

      if (SafeArrayGetDim( pArr ) != 1) {
         return -1;                                         // Only one dimensional arrays supported by OPC V1.0A.
      }

      hres = SafeArrayGetLBound( pArr, 1, &LBound );        // Dimension set to 1, Check Return Code
      if (FAILED( hres )) {                                 
         return -1;
      }
      hres = SafeArrayGetUBound( pArr, 1, &UBound );        // Dimension set to 1, Check Return Code
      if (FAILED( hres )) {                                 
         return -1;
      }
      hres = SafeArrayLock( pArr );                         // Lock for direct array manipulations
      if (FAILED( hres )) {                                 
         return -1;
      }

      if (vtBase == VT_BSTR) {                              // It's an array of BSTRs

         BSTR     bstr;
                                                            
         for (i=LBound; i<=UBound; i++) {                   // For size use <=

            bstr = ((BSTR *)pArr->pvData)[i - LBound];      // Direct array memory access
            if (bstr != NULL) {
               len += SysStringByteLen( bstr );
            }
            len += sizeof (DWORD) + sizeof (WCHAR);         // Add size for len and 0.
         }
      }
      else if (vtBase == VT_VARIANT) {
                                                            // It's an array of VARIANTs
         long     nlen;
         VARIANT* pVariantElem;

         for (i=LBound; i<=UBound; i++) {                   // For size use <=
                                                            // Direct array memory access
            pVariantElem = &((VARIANT *)pArr->pvData)[i - LBound];
            if (pVariantElem == NULL) {
               SafeArrayUnlock( pArr );
               return -1;                                   // Check Return Code
            }
            nlen = SizePackedVariant( pVariantElem );
            if (nlen == -1) {
               SafeArrayUnlock( pArr );
               return -1;                                   // Error
            }
            len += nlen;
         }
      }
      else {                                                // Fixed size elements
                                                            // +1 for Array size
         len += ((UBound - LBound)+1) * SafeArrayGetElemsize( pArr );
      }

      hres = SafeArrayUnlock( pArr );
      if (FAILED( hres )) {
         return -1;
      }
   }
   else {
                                                            // Simple type
      if (vtBase == VT_BSTR) {
         if (V_BSTR( vp ) != NULL) {
            len += SysStringByteLen( V_BSTR( vp ) );
         }
         len += sizeof (DWORD) + sizeof (WCHAR);            // Add size for len and 0.
      }
   }
   return len;
}





//-----------------------------------------------------------------------
// Copy the VARIANT structure and append the data
// for BSTR and Array types.
// If -1 is returned some error occured
//-----------------------------------------------------------------------
long CopyPackVariant( char *dest, VARIANT *vp )
{
      HRESULT     hres;
      SAFEARRAY*  pArr;
      long        LBound, UBound;
      long        i;
      DWORD       len, nlen;   
      VARTYPE     vtBase;


   len = 0;
   memcpy( dest, vp, sizeof (VARIANT) );                    // Copy the VARIANT structure
      
   dest += sizeof (VARIANT);                                // Point behind struct
   len  += sizeof (VARIANT);

   vtBase = V_VT( vp );
   
   if ((vtBase & VT_BYREF) != 0) {
      return -1;                                            // Don't accept by ref 
   }
   else if ((vtBase & VT_ARRAY) != 0) {
                                                            // This is a safe array!
      vtBase = vtBase - VT_ARRAY;

      pArr = V_ARRAY( vp );

      if (SafeArrayGetDim( pArr ) != 1) {
         return -1;                                         // Only one dimensional arrays supported by OPC V1.0A.
      }
      hres = SafeArrayGetLBound( pArr, 1, &LBound );        // Dimension always 1
      if (FAILED( hres )) {                                 
         return -1;
      }
      hres = SafeArrayGetUBound( pArr, 1, &UBound );        // Dimension always 1
      if (FAILED( hres )) {
         return -1;
      }
      hres = SafeArrayLock( pArr );                         // Lock for direct array manipulations
      if (FAILED( hres )) {                                 
         return -1;
      }

      memcpy( dest, pArr, sizeof (SAFEARRAY) );             // Copy SAFEARRAY data structure
      ((SAFEARRAY *)dest)->pvData = NULL;
      dest += sizeof (SAFEARRAY);
      len  += sizeof (SAFEARRAY);

                                                            // Copy data of array
      if (vtBase == VT_BSTR) {
                                                            // It's an array of BSTRs
         BSTR     bstr;

         for (i=LBound; i<=UBound; i++) {                   // For size use <=

            bstr = ((BSTR *)pArr->pvData)[i - LBound];      // Direct array memory access
            if (bstr == NULL) {
               nlen = 0;
            }
            else {
               nlen = SysStringByteLen( bstr ) ;            // Include 00 at the end 
            }
               
            *((DWORD *)dest) = nlen;                        // Copy size of string
            dest += sizeof (DWORD);
            len  += sizeof (DWORD);
            if (nlen > 0) {
               memcpy( dest, bstr, nlen + sizeof (WCHAR) ); // Copy string with 00
            }
            else {
               memset( dest, 0, sizeof (WCHAR) );           // Set only 00
            }

            dest += nlen + sizeof (WCHAR);
            len  += nlen + sizeof (WCHAR);
         }
      }
      else if (vtBase == VT_VARIANT) {
                                                            // It's an array of VARIANTs
         VARIANT* pVariantElem;

         for (i=LBound; i<=UBound; i++) {                   // For size use <=

            pVariantElem = &((VARIANT *)pArr->pvData)[i - LBound];
            if (pVariantElem == NULL) {
               SafeArrayUnlock( pArr );
               return -1;                                   // Check Return Code
            }
            nlen = CopyPackVariant( dest, pVariantElem );   // Recursive call
            if (nlen == -1) {
               SafeArrayUnlock( pArr );
               return -1;                                   // Error
            }
            dest += nlen;
            len  += nlen;
         }
      }
      else {                                                // Fixed size elements
                                                            // +1 for Array size
         nlen = ((UBound - LBound)+1) * SafeArrayGetElemsize( pArr );
         memcpy( dest, pArr->pvData, nlen );
         dest += nlen;
         len  += nlen;
      }

      hres = SafeArrayUnlock( pArr );
      if (FAILED( hres )) {
         return -1;
      }
   }
   else {
                                                            // Simple type
      if (vtBase == VT_BSTR) {

         if (V_BSTR(  vp ) == NULL) {
            nlen = 0;
         }
         else {
            nlen = SysStringByteLen( V_BSTR( vp ) );        // Count bytes 
         }
         
         *((DWORD *)dest) = nlen;                           // Copy size of string
         dest += sizeof (DWORD);
         len  += sizeof (DWORD);

         if (nlen > 0) {
                                                            // Copy string with 00
            memcpy( dest, V_BSTR( vp ), nlen + sizeof (WCHAR) );
         }
         else {
            memset( dest, 0, sizeof (WCHAR) );              // Set only 00
         }
         dest += nlen + sizeof (WCHAR);
         len  += nlen + sizeof (WCHAR);
      }
   }
   return len;
}

 
//DOM-IGNORE-END
