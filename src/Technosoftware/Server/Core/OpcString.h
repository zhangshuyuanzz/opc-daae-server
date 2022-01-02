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

#ifndef _COpcString_H_
#define _COpcString_H_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "OpcDefs.h"

#define OPC_EMPTY_STRING _T("")

//==============================================================================
// Class:   COpcString
// PURPOSE: Implements a string class.

class OPCUTILS_API COpcString
{
    OPC_CLASS_NEW_DELETE_ARRAY();

public:

    //==========================================================================
    // Operators

    // Constructor
    COpcString();
    COpcString(LPCSTR szStr);
    COpcString(LPCWSTR wszStr);
    COpcString(const GUID& cGuid);

    // Copy Constructor
    COpcString(const COpcString& cStr);

    // Destructor
    ~COpcString();

    // Cast
    operator LPCSTR() const;
    operator LPCWSTR() const;

    // Assignment
    COpcString& operator=(const COpcString& cStr);

    // Append
    COpcString& operator+=(const COpcString& cStr);

    // Index
    TCHAR& operator[](UINT uIndex);
    TCHAR  operator[](UINT uIndex) const;

    // Comparison
    int Compare(const COpcString& cStr) const;

    bool operator==(LPCSTR szStr) const {return (Compare(szStr) == 0);}
    bool operator<=(LPCSTR szStr) const {return (Compare(szStr) <= 0);}
    bool operator <(LPCSTR szStr) const {return (Compare(szStr)  < 0);}
    bool operator!=(LPCSTR szStr) const {return (Compare(szStr) != 0);}
    bool operator >(LPCSTR szStr) const {return (Compare(szStr)  > 0);}
    bool operator>=(LPCSTR szStr) const {return (Compare(szStr) >= 0);}  
    
    bool operator==(LPCWSTR szStr) const {return (Compare(szStr) == 0);}
    bool operator<=(LPCWSTR szStr) const {return (Compare(szStr) <= 0);}
    bool operator <(LPCWSTR szStr) const {return (Compare(szStr)  < 0);}
    bool operator!=(LPCWSTR szStr) const {return (Compare(szStr) != 0);}
    bool operator >(LPCWSTR szStr) const {return (Compare(szStr)  > 0);}
    bool operator>=(LPCWSTR szStr) const {return (Compare(szStr) >= 0);}

    bool operator==(const COpcString& szStr) const {return (Compare(szStr) == 0);}
    bool operator<=(const COpcString& szStr) const {return (Compare(szStr) <= 0);}
    bool operator <(const COpcString& szStr) const {return (Compare(szStr)  < 0);}
    bool operator!=(const COpcString& szStr) const {return (Compare(szStr) != 0);}
    bool operator >(const COpcString& szStr) const {return (Compare(szStr)  > 0);}
    bool operator>=(const COpcString& szStr) const {return (Compare(szStr) >= 0);}

    // Addition
    OPCUTILS_API friend COpcString operator+(const COpcString& cStr1, LPCSTR szStr2);
    OPCUTILS_API friend COpcString operator+(const COpcString& cStr1, LPCWSTR wszStr2);
    OPCUTILS_API friend COpcString operator+(const COpcString& cStr1, const COpcString& cStr2);
    OPCUTILS_API friend COpcString operator+(LPCSTR  szStr1,          const COpcString& cStr2);
    OPCUTILS_API friend COpcString operator+(LPCWSTR wszStr1,         const COpcString& cStr2);

    //==========================================================================
    // Public Methods

    // GetLength
    UINT GetLength() const;

    // IsEmpty
    bool IsEmpty() const;

    // Empty
    void Empty() {Free();}

    // ToGuid
    bool ToGuid(GUID& tGuid) const;

    // FromGuid
    void FromGuid(const GUID& tGuid);

    // GetBuffer
    LPTSTR GetBuffer();

    // SetBuffer
    void SetBuffer(UINT uLength);

    // Find
    int Find(LPCTSTR tsTarget) const;

    // ReverseFind
    int ReverseFind(LPCTSTR tsTarget) const;

    // SubStr
    COpcString SubStr(UINT uStart, UINT uCount = -1) const;

    // Trim
    COpcString& Trim();

    // ToLower
    COpcString ToLower(UINT uIndex = -1);

    // ToUpper
    COpcString ToUpper(UINT uIndex = -1);

    // Clone
    static LPSTR Clone(LPCSTR szStr);

    // Clone
    static LPWSTR Clone(LPCWSTR wszStr);

    // ToMultiByte
    static LPSTR ToMultiByte(LPCWSTR wszStr, int iwszLen = -1);

    // ToUnicode
    static LPWSTR ToUnicode(LPCSTR szStr, int iszLen = -1);

private:

    // TStrBuf
    struct TStrBuf
    {
        UINT   uRefs;
        LPSTR  szStr;
        LPWSTR wszStr;
    };

    //==========================================================================
    // Private Methods
    
    // Set
    void Set(LPCSTR szStr);

    // Set
    void Set(LPCWSTR wszStr);

    // Set
    void Set(const COpcString& cStr);
    
    // Free
    void Free();

    // Alloc
    static TStrBuf* Alloc(UINT uLength);

    //==========================================================================
    // Private Members

    TStrBuf* m_pBuf;
};

//==============================================================================
// FUNCTION: Comparisons
// PURPOSE:  Compares two strings.

OPCUTILS_API inline bool operator==(LPCSTR szStr1, const COpcString& cStr2) {return (cStr2.Compare(szStr1) == 0);}
OPCUTILS_API inline bool operator<=(LPCSTR szStr1, const COpcString& cStr2) {return (cStr2.Compare(szStr1) <= 0);}
OPCUTILS_API inline bool operator <(LPCSTR szStr1, const COpcString& cStr2) {return (cStr2.Compare(szStr1)  < 0);}
OPCUTILS_API inline bool operator!=(LPCSTR szStr1, const COpcString& cStr2) {return (cStr2.Compare(szStr1) != 0);}
OPCUTILS_API inline bool operator >(LPCSTR szStr1, const COpcString& cStr2) {return (cStr2.Compare(szStr1)  > 0);}
OPCUTILS_API inline bool operator>=(LPCSTR szStr1, const COpcString& cStr2) {return (cStr2.Compare(szStr1) >= 0);}  

OPCUTILS_API inline bool operator==(LPCWSTR szStr1, const COpcString& cStr2) {return (cStr2.Compare(szStr1) == 0);}
OPCUTILS_API inline bool operator<=(LPCWSTR szStr1, const COpcString& cStr2) {return (cStr2.Compare(szStr1) <= 0);}
OPCUTILS_API inline bool operator <(LPCWSTR szStr1, const COpcString& cStr2) {return (cStr2.Compare(szStr1)  < 0);}
OPCUTILS_API inline bool operator!=(LPCWSTR szStr1, const COpcString& cStr2) {return (cStr2.Compare(szStr1) != 0);}
OPCUTILS_API inline bool operator >(LPCWSTR szStr1, const COpcString& cStr2) {return (cStr2.Compare(szStr1)  > 0);}
OPCUTILS_API inline bool operator>=(LPCWSTR szStr1, const COpcString& cStr2) {return (cStr2.Compare(szStr1) >= 0);}

#endif // _COpcString_H_

