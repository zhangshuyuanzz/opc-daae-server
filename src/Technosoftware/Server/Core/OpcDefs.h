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

#ifndef _OpcDefs_H_
#define _OpcDefs_H_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include <crtdbg.h>
#include <malloc.h>

//==============================================================================
// MACRO:   OPC_ASSERT
// PURPOSE: Halts program execution if condition is not met.

#ifdef _DEBUG
extern void OpcAssert(bool bCondition, LPCWSTR szMessage);
#define OPC_ASSERT(x) OpcAssert(x, L#x)
#else
#define OPC_ASSERT(x)
#endif

//==============================================================================
// MACRO:   TRY, LEAVE. FINALLY
// PURPOSE: Macros to facilitate cleanup after error conditions.

#pragma warning(disable: 4102) 

#ifndef OPC_NO_EXCEPTIONS

#define TRY try 
#define CATCH catch (DWORD)
#define CATCH_FINALLY catch (DWORD) {} finally:
#define FINALLY finally:
#define THROW() throw (DWORD)0;
#define THROW_(xResult, xValue) { (xResult) = (xValue); throw (DWORD)0; }
#define GOTOFINALLY() goto finally;

#else

#define TRY {    
#define CATCH } goto finally; onerror:
#define CATCH_FINALLY onerror: finally:
#define FINALLY finally:
#define THROW() goto onerror;
#define THROW_(xResult, xValue) { (xResult) = (xValue); goto onerror; }
#define GOTOFINALLY() goto finally;

#endif

#define OPCUTILS_API

//==============================================================================
// FUNCTION: OpcAlloc
// PURPOSE:  Allocates a block of memory.

OPCUTILS_API void* OpcAlloc(size_t tSize);

//==============================================================================
// FUNCTION: OpcFree
// PURPOSE:  Frees a block of memory.

OPCUTILS_API void OpcFree(void* pBlock);

//==============================================================================
// MACRO:   OpcArrayAlloc
// PURPOSE: Allocates an array of simple data (i.e. not classes).

#define OpcArrayAlloc(xType,xLength) (xType*)OpcAlloc((xLength)*sizeof(xType))

//==============================================================================
// FUNCTION: OpcStrDup
// PURPOSE:  Duplicates a string.

OPCUTILS_API CHAR* OpcStrDup(LPCSTR szValue);
OPCUTILS_API WCHAR* OpcStrDup(LPCWSTR szValue);

//==============================================================================
// FUNCTION: OpcArrayDup
// PURPOSE:  Duplicates an array of types (bitwise copy must be valid).

#define OpcArrayDup(xCopy, xType, xLength, xValue) \
{ \
    UINT uLength = (xLength)*sizeof(xType); \
    xCopy = (xType*)OpcAlloc(uLength); \
 \
    if (xValue != NULL) memcpy(xCopy, xValue, uLength); \
    else memset(xCopy, 0, uLength); \
}

//==============================================================================
// MACRO:   OPC_CLASS_NEW_DELETE
// PURPOSE: Implements a class new and delete operators.

#define OPC_CLASS_NEW_DELETE() \
public: void* operator new(size_t nSize) {return OpcAlloc(nSize);} \
public: void operator delete(void* p) {OpcFree(p);}

#define OPC_CLASS_NEW_DELETE_EX(xCallType) \
xCallType void* operator new(size_t nSize) {return OpcAlloc(nSize);} \
xCallType void operator delete(void* p) {OpcFree(p);}

//==============================================================================
// MACRO:   OPC_CLASS_NEW_DELETE_ARRAY
// PURPOSE: Implements a class new and delete instance and array operators.

#define OPC_CLASS_NEW_DELETE_ARRAY() \
OPC_CLASS_NEW_DELETE() \
public: void* operator new[](size_t nSize) {return OpcAlloc(nSize);} \
public:  void operator delete[](void* p) {OpcFree(p);}

#define OPC_CLASS_NEW_DELETE_ARRAY_EX(xCallType) \
OPC_CLASS_NEW_DELETE_EX(xCallType) \
xCallType void* operator new[](size_t nSize) {return OpcAlloc(nSize);} \
xCallType void operator delete[](void* p) {OpcFree(p);}

//==============================================================================
// FUNCTION: operator
// PURPOSE:  Implements various operators for CY structures.

inline bool operator==(const CY& a, const CY& b) 
{
	return (a.int64 == b.int64); 
}

inline bool operator!=(const CY& a, const CY& b) 
{ 
	return (a.int64 != b.int64); 
}

inline bool operator<(const CY& a, const CY& b)
{
	return (a.int64 < b.int64); 
}

inline bool operator<=(const CY& a, const CY& b)
{
	return (a.int64 <= b.int64); 
}

inline bool operator>(const CY& a, const CY& b)
{
	return (a.int64 > b.int64); 
}

inline bool operator>=(const CY& a, const CY& b)
{
	return (a.int64 >= b.int64); 
}

//==============================================================================
// MACRO:    OPC_MASK_XXX
// PURPOSE:  Facilitate manipulation of bit masks.

#define OPC_MASK_ISSET(xValue, xMask) (((xValue) & (xMask)) == (xMask))
#define OPC_MASK_SET(xValue, xMask)   xValue |= (xMask)
#define OPC_MASK_UNSET(xValue, xMask) xValue &= ~(xMask)

#endif // _OpcDefs_H_

