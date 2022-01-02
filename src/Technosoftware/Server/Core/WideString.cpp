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

//-----------------------------------------------------------------------------
// INLCUDE
//-----------------------------------------------------------------------------
#include <stdafx.h>
#include "UtilityDefs.h"
#include "WideString.h"

//-----------------------------------------------------------------------------
// CODE
//-----------------------------------------------------------------------------

//=============================================================================
// Construction / Destruction
//=============================================================================
WideString::WideString()
{
   str_ = NULL;
}

WideString::~WideString()
{
   if (str_) {
      delete [] str_;
   }
}


//-----------------------------------------------------------------------------
// OPERATIONS
//-----------------------------------------------------------------------------

HRESULT WideString::SetString( LPCWSTR str )
{
   LPWSTR strold = str_;                       // store old value

   str_ = new WCHAR [wcslen( str ) + 1];
   if (!str_) {
      str_ = strold;                           // restore to old value
      return E_OUTOFMEMORY;
   }
   wcscpy( str_, str );						// set new value

   if (strold) {                                // release old value
      delete [] strold;
   }
   return S_OK;
}


HRESULT WideString::AppendString( LPCWSTR str )
{
   LPWSTR strold = str_;                       // store old value

   str_ = new WCHAR [wcslen( strold ) + wcslen( str ) + 1];
   if (!str_) {
      str_ = strold;                           // restore to old value
      return E_OUTOFMEMORY;
   }
   wcscpy( str_, strold );                     // set existing value
   wcscat( str_, str );                        // append new value

   if (strold) {                                // release old value
      delete [] strold;
   }
   return S_OK;
}


LPWSTR WideString::Copy()
{
   LPWSTR copy = new WCHAR [wcslen( str_ ) + 1];
   if (copy) {
      wcscpy( copy, str_ );
   }
   return copy;
}


LPWSTR WideString::CopyCOM()
{
   LPWSTR copy = ComAlloc<WCHAR>( (DWORD)wcslen( str_ ) + 1 );
   if (copy) {
      wcscpy( copy, str_ );
   }
   return copy;
}


BSTR WideString::CopyBSTR()
{
   return ::SysAllocString( str_ );
}
//DOM-IGNORE-END
