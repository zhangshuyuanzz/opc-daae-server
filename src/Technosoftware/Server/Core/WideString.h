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

#ifndef __WideString_H_
#define __WideString_H_

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000


//-----------------------------------------------------------------------
// CLASS
//-----------------------------------------------------------------------
class WideString
{
// Construction / Destruction
public:

   /**
    * @fn	WideString();
    *
    * @brief	Default constructor.
    */

   WideString();

   /**
    * @fn	~WideString();
    *
    * @brief	Destructor.
    */

   ~WideString();

// Attributes
public:

   /**
    * @fn	inline operator LPCWSTR() const
    *
    * @brief	Cast that converts the given string to a LPCWSTR.
    *
    * @return	The result of the operation.
    */

   inline operator LPCWSTR() const { return str_; }

   /**
    * @fn	inline operator LPWSTR()
    *
    * @brief	Cast that converts the given string  to a LPWSTR.
    *
    * @return	The result of the operation.
    */

   inline operator LPWSTR() { return str_; }

// Operations
public:

   /**
    * @fn	HRESULT SetString( LPCWSTR str );
    *
    * @brief	This method stores a copy of the specified string. 
    * 			If the new string cannot be copied then the old string is still valid.
    *
    * @param	str	The string.
    *
    * @return	A hResult.
    */

   HRESULT  SetString( LPCWSTR str );

   /**
    * @fn	HRESULT AppendString( LPCWSTR str );
    *
    * @brief	This method appends a copy of the specified string to the existing string. 
    * 			If the new string cannot be appended then the old string is still valid.
    *
    * @param	str	The string.
    *
    * @return	A hResult.
    */

   HRESULT  AppendString( LPCWSTR str );

   /**
    * @fn	LPWSTR Copy();
    *
    * @brief	This method returns a copy of the wide string.
    * 			Memory is allocated from the heap.
    *
    * @return	A LPWSTR.
    */

   LPWSTR   Copy();

   /**
    * @fn	LPWSTR CopyCOM();
    *
    * @brief	This method returns a copy of the wide string.
    * 			Memory is allocated from the COM memory manager.
    *
    * @return	A LPWSTR.
    */

   LPWSTR   CopyCOM();

   /**
    * @fn	BSTR CopyBSTR();
    *
    * @brief	This method returns a BSTR copy of the wide string.
    *
    * @return	A BSTR.
    */

   BSTR     CopyBSTR();

   // Comparison operators
   BOOL     operator==( LPCWSTR str ) const;
   BOOL     operator==( WideString & str ) const;

// Implementation
protected:
   LPWSTR   str_;
};

/**
 * @fn	inline BOOL WideString::operator==( LPCWSTR str ) const
 *
 * @brief	Comparison operators.
 *
 * @param	str	The string.
 *
 * @return	true if the parameters are considered equivalent.
 */

inline BOOL WideString::operator==( LPCWSTR str ) const
{
   return (wcscmp( str_, str ) == 0) ? TRUE : FALSE;
}

/**
 * @fn	inline BOOL WideString::operator==( WideString & str ) const
 *
 * @brief	Equality operator.
 *
 * @param [in,out]	str	The string.
 *
 * @return	true if the parameters are considered equivalent.
 */

inline BOOL WideString::operator==( WideString & str ) const
{
   return (wcscmp( str_, str.str_ ) == 0) ? TRUE : FALSE;
}
//DOM-IGNORE-END

#endif //__WideString_H_
