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

#ifndef __UtilityFuncs_H_
#define __UtilityFuncs_H_

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000


// VARIANT to TimeBias
HRESULT ConvertVariantToTimeBias( 
            VARIANT sourceTimeBias, 
            long **destTimeBias 
       );

   // Blob: VARIANT to BYTE*
HRESULT ConvertBlobVariantToBytes( 
            VARIANT  *pSourceBlob,
            DWORD    *pDestBlobSize,
            BYTE     **ppDestBlob,
            IMalloc  *pmem
         );

   // Blob: BYTE* to VARIANT
HRESULT ConvertBlobBytesToVariant(
            DWORD    SourceBlobSize,
            BYTE     *pSourceBlob, 
            VARIANT  **ppDestBlob, 
            IMalloc  *pmem
       );


#define  ComAllOPCtring(s) WSTRClone( s, pIMalloc )
#define  ComFreeString(s)  WSTRFree( s, pIMalloc )

WCHAR * WSTRClone( const WCHAR *oldstr, IMalloc *pmem = NULL );
void WSTRFree( WCHAR * c, IMalloc *pmem = NULL );


HRESULT VariantInitArrayOf( VARIANT    *pVar, 
                            long       size, 
                            VARTYPE    vt );


HRESULT VariantUninitArrayOf( VARIANT *pVar );

HRESULT CopyOPCITEMDEF( OPCITEMDEF *dest, OPCITEMDEF *src );
HRESULT FreeOPCITEMDEF( OPCITEMDEF *idef );

HRESULT FileTimeToDATE( FILETIME *srcFT, DATE &destDATE );

//DOM-IGNORE-END

#endif //__UtilityFuncs_H_

