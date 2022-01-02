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
#include "UtilityFuncs.h"

//=======================================================================
// as the variant representation of a Blob is not known
// here are the conversion functions
// if something doesn't work with your Blobs
// change here!
//=======================================================================
HRESULT ConvertBlobVariantToBytes( 
            VARIANT  *pSourceBlob,
            DWORD    *pDestBlobSize,
            BYTE     **ppDestBlob,
            IMalloc  *pmem
         )
{
      long           theElemSize;
      long           theNrOfElem;
      long           theBlobSize;

      // now transform a Variant to a BYTE *
      // here we assume a blob is an array of something (probably VT_UI1)
   if ((pSourceBlob == NULL) || (!V_ISARRAY( pSourceBlob ))) {
         // cannot recognize the Blob ... take it as empty
      *ppDestBlob = NULL; 
      *pDestBlobSize = 0;
   } 
   else {         // the blob is not empty
      theElemSize = SafeArrayGetElemsize( V_ARRAY( pSourceBlob ) );
      SafeArrayGetUBound( V_ARRAY( pSourceBlob ), 1, &theNrOfElem );
      theBlobSize = theNrOfElem * theElemSize;
      if( pmem == NULL ) {
         *ppDestBlob = new BYTE[ theBlobSize ];
      } else {
         *ppDestBlob = (BYTE *) pmem->Alloc( theBlobSize );
      }
      memcpy( *ppDestBlob, V_ARRAY( pSourceBlob )->pvData, theBlobSize );
      *pDestBlobSize = theBlobSize;
   }

   return S_OK;
}

//=======================================================================
//=======================================================================
HRESULT ConvertBlobBytesToVariant(
            DWORD    SourceBlobSize,
            BYTE     *pSourceBlob, 
            VARIANT  **ppDestBlob, 
            IMalloc  *pmem
       )
{
      VARIANT           *pBlob;
      SAFEARRAYBOUND    boundsize;
      HRESULT           res;

   pBlob = *ppDestBlob;

   if ( pmem == NULL ) {
      pBlob = new VARIANT[1];
   } else {
      pBlob = (VARIANT *)pmem->Alloc( sizeof( VARIANT ) );
   }

   if ( pBlob == NULL ) {
      return E_OUTOFMEMORY;
   }
   VariantInit( pBlob );
   V_VT( pBlob ) = VT_ARRAY | VT_UI1;
   boundsize.cElements = SourceBlobSize;
   boundsize.lLbound = 0;
   V_ARRAY( pBlob ) = SafeArrayCreate( VT_UI1, 1, &boundsize );
   if (V_ARRAY( pBlob ) == NULL ) {
      res = E_OUTOFMEMORY;
      goto BlobToVariantExit0;
   }
   memcpy( V_ARRAY( pBlob )->pvData, pSourceBlob, SourceBlobSize );
   *ppDestBlob = pBlob;
   return S_OK;

BlobToVariantExit0:
   if ( pmem == NULL ) {
      delete pBlob;
   } else {
      pmem->Free( pBlob );
   }
   return res;
}




//=======================================================================
// Clone a Wide String 
//
//=======================================================================
WCHAR * WSTRClone(const WCHAR *oldstr, IMalloc *pmem /* =NULL */ )
{
   WCHAR *newstr;
   if (pmem) 
      newstr = (WCHAR*)pmem->Alloc(sizeof(WCHAR) * (wcslen(oldstr) + 1));
   else 
      newstr = new WCHAR[wcslen(oldstr) + 1];

   if(newstr) 
      wcscpy(newstr, oldstr);
   return newstr;
}




//=======================================================================
// Free a Wide String 
//
//=======================================================================
void WSTRFree( WCHAR * c, IMalloc *pmem /* =NULL */ )
{
   if(c == NULL) return;

   if(pmem) pmem->Free(c);
   else     delete [] c;
}




//=======================================================================
// Init a variant to an array of something
//=======================================================================
HRESULT VariantInitArrayOf( VARIANT    *pVar, 
                            long       size, 
                            VARTYPE    vt )
{
      SAFEARRAYBOUND SArgsabound[1];

   SArgsabound[0].cElements = size;
   SArgsabound[0].lLbound = 0;

   //VariantInit(pVar);
   VariantClear(pVar);
   VariantInit(pVar);
   V_VT( pVar ) = VT_ARRAY | vt;
   V_ARRAY( pVar ) = SafeArrayCreate( vt, 1, SArgsabound ); 

   if ( V_ARRAY( pVar ) == NULL ) {
      return E_OUTOFMEMORY;
   } else {
      return S_OK;
   }
}


//=======================================================================
// the invers of VariantInitArrayOf
//=======================================================================
HRESULT VariantUninitArrayOf( VARIANT *pVar )
{
   if ( pVar == NULL ) {
      return E_FAIL;
   }
   //SafeArrayDestroy( pVar->parray );
   //VariantInit(pVar);
   VariantClear(pVar);
   VariantInit(pVar);
   return S_OK;
}


//=======================================================================
// Copy a OPCITEMDEF structure
//=======================================================================
HRESULT CopyOPCITEMDEF( OPCITEMDEF *dest, OPCITEMDEF *src )
{

   if ( src->szAccessPath != NULL ) {
      dest->szAccessPath = WSTRClone( src->szAccessPath, NULL );
      if ( dest->szAccessPath == NULL ) {
         goto CopyOPCITEMDEFExit0;
      }
   } else {
      dest->szAccessPath = NULL;
   }

   if ( src->szItemID != NULL ) {
      dest->szItemID = WSTRClone( src->szItemID, NULL );
      if ( dest->szItemID == NULL ) {
         goto CopyOPCITEMDEFExit1;
      }
   } else {
      dest->szItemID = NULL;
   }

   dest->bActive = src->bActive;
   dest->hClient = src->hClient;

   dest->dwBlobSize = src->dwBlobSize;
   if ( src->pBlob != NULL ) {
      dest->pBlob = new BYTE[ src->dwBlobSize ];
      if ( dest->pBlob == NULL ) {
         goto CopyOPCITEMDEFExit2;
      }
      memcpy( dest->pBlob, src->pBlob, src->dwBlobSize );
   } else {
      dest->pBlob = NULL;
   }

   dest->vtRequestedDataType = src->vtRequestedDataType;
   dest->wReserved = src->wReserved;

   return S_OK;

CopyOPCITEMDEFExit2:
   WSTRFree( dest->szItemID, NULL );

CopyOPCITEMDEFExit1:
   WSTRFree( dest->szAccessPath, NULL );

CopyOPCITEMDEFExit0:
   return E_OUTOFMEMORY;
}



//=======================================================================
// Free OPCITEMDEF structure
// (only its fields)
//=======================================================================
HRESULT FreeOPCITEMDEF( OPCITEMDEF *idef )
{
   if ( idef->szAccessPath != NULL ) {
      WSTRFree( idef->szAccessPath, NULL );
   }

   if ( idef->szItemID != NULL ) {
      WSTRFree( idef->szItemID, NULL );
   }

   if ( idef->pBlob != NULL ) {
      delete [] idef->pBlob;
   }

   return S_OK;
}



//=======================================================================
// Convert FILETIME to DATE
//=======================================================================
HRESULT FileTimeToDATE( FILETIME *srcFT, DATE &destDATE )
{
   SYSTEMTIME  systime;

   FILETIME    localFT;

   if (FileTimeToLocalFileTime( srcFT, &localFT) == 0) {
      return HRESULT_FROM_WIN32( GetLastError() );
   }
   if (FileTimeToSystemTime( &localFT, &systime ) == 0) {
      return HRESULT_FROM_WIN32( GetLastError() );
   }
   if (SystemTimeToVariantTime( &systime, &destDATE ) == 0) {
      // Note :   Regardless of what the documentation says, the return
      //          value is zero if the function fails and nonzero if
      //          the function succeeds.
      return E_FAIL;      
   }
   return S_OK;
}
//DOM-IGNORE-END
