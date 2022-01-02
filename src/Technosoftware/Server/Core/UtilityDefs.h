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

#ifndef __UtilityDefs_H_
#define __UtilityDefs_H_

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000


//-------------------------------------------------------------------------
// DEFINES
//-------------------------------------------------------------------------
//    The handle is invalid.
#define  E_INVALID_HANDLE  ((HRESULT)HRESULT_FROM_WIN32( ERROR_INVALID_HANDLE ))

//    The object is already in the list.
#define  E_OBJECT_IN_LIST  ((HRESULT)HRESULT_FROM_WIN32( ERROR_OBJECT_IN_LIST ))


/////////////////////////////////////////////////////////////////////////////
// Class CRefClass declaration
/////////////////////////////////////////////////////////////////////////////
class CRefClass
{
// Construction
public:
   CRefClass() : m_lRefCount(0) {}
protected:
   virtual ~CRefClass() {}                      // Destructor only called by Release()

// Operations
public:                                         
   inline void AddRef()                         // Works like COM AddRef
      { InterlockedIncrement( &m_lRefCount ); }

   inline void Release()                        // Works like COM Release
      {
         if (InterlockedDecrement( &m_lRefCount ) == 0 )
            delete this;
      }

// Implementation
private:
   LONG  m_lRefCount;
};


/////////////////////////////////////////////////////////////////////////////
// Class CSimplePtrArray declaration
/////////////////////////////////////////////////////////////////////////////
// Simple Array of Pointers
//    Behaviour is identical with CSimpleArray from ATL except that the
//    instances are deleted too.
/////////////////////////////////////////////////////////////////////////////
template <class T>
class CSimplePtrArray : public CSimpleArray< T >
{
public:
   // Overrides
   // Note :   The function members of the base class are not virtual.
   //          Therefore the destructor and the function Remove() which uses
   //          the functions RemoveAll() and RemoveAt() are implemendted too.
   ~CSimplePtrArray()
   {
      RemoveAll();
   }
   BOOL Remove(T& t)
   {
      int nIndex = Find(t);
      if(nIndex == -1)
         return FALSE;
      return RemoveAt(nIndex);
   }
   BOOL RemoveAt(int nIndex)
   {
      ATLASSERT(nIndex >= 0 && nIndex < m_nSize);
      if (nIndex < 0 || nIndex >= m_nSize) {
         return FALSE;
      }
      delete m_aT[nIndex];

      if(nIndex != (m_nSize - 1)) {
         memmove( (void*)(m_aT + nIndex), (void*)(m_aT + nIndex + 1), (m_nSize - (nIndex + 1)) * sizeof(T) );
      }
      m_nSize--;
      return TRUE;
   }
   void RemoveAll()
   {
      if(m_aT != NULL)
      {
         for(int i = 0; i < m_nSize; i++)
            delete m_aT[i];
         free(m_aT);
         m_aT = NULL;
      }
      m_nSize = 0;
      m_nAllocSize = 0;
   }
};


/////////////////////////////////////////////////////////////////////////////
// Class CSimpleValQueue declaration
/////////////////////////////////////////////////////////////////////////////
// Queue Template for Queues of simple data types
/////////////////////////////////////////////////////////////////////////////
template <class T>
class CSimpleValQueue
{
// Construction / Destruction
public:
   CSimpleValQueue() : m_aT(NULL), m_dwSize(0), m_dwAllOPCize(0) {}
   ~CSimpleValQueue() { if (m_aT) free( m_aT); }

// Operators
public:
   HRESULT PreAllocate( DWORD dwAllOPCize )
      {
         T* aT = (T*)realloc( m_aT, dwAllOPCize * sizeof (T) );
         if (aT) {
            m_dwAllOPCize = dwAllOPCize;
            m_aT = aT;
            return S_OK;
         }
         return E_OUTOFMEMORY;
      }

   HRESULT Add( T t)
      {
         if (m_dwSize == m_dwAllOPCize) {
            DWORD dwNewAllOPCize = m_dwAllOPCize + 10;
            T* aT = (T*)realloc( m_aT, dwNewAllOPCize * sizeof (T) );
            if (!aT) return E_OUTOFMEMORY;

            m_dwAllOPCize = dwNewAllOPCize;
            m_aT = aT;
         }
         m_aT[ m_dwSize++ ] = t;
         return S_OK;
      }

   void RemoveFirstN( DWORD dwNum )
      {
         _ASSERTE( dwNum && dwNum <= m_dwSize);
         memmove( m_aT, (void*)&m_aT[dwNum], (m_dwSize - dwNum) * sizeof (T) );
         m_dwSize -= dwNum;
      }

   inline DWORD   GetSize() const { return m_dwSize; }
   inline void    RemoveAll() { m_dwSize = 0; }
   inline T&      operator[]( DWORD dwIndex ) const
      {
         _ASSERTE( dwIndex < m_dwSize);
         return m_aT[ dwIndex ];
      }

// Implementation
protected:
   T*    m_aT;
   DWORD m_dwSize;
   DWORD m_dwAllOPCize;
};


#if (_ATL_VER < 0x0700)                         // Not required for ATL versions 7.0 or higher
/////////////////////////////////////////////////////////////////////////////
// Class CAutoVectorPtr declaration
/////////////////////////////////////////////////////////////////////////////
// Contains a subset of the functionality from the CAutoVectorPtr template
// which is included in ATL since version 7.0.
/////////////////////////////////////////////////////////////////////////////

   template< typename T >
   class CAutoVectorPtr
   {
   public:
      CAutoVectorPtr() throw() : m_p( NULL ) {}
      ~CAutoVectorPtr() throw() { Free(); }

      bool Allocate( size_t nElements ) throw()
         {
            _ASSERTE( m_p == NULL );
            m_p = new T[ nElements ];
            return m_p ? true : false;
         }

      void Free() throw() { try { if (m_p) { delete [] m_p; m_p = NULL; } } catch (...) {} }

      operator T*() const throw() { return( m_p ); }

   protected:
      T* m_p;
   };
#endif

template <class T> T* ComAlloc( DWORD dwNum = 1 ) { return (T*)pIMalloc->Alloc( sizeof (T) * dwNum ); }
static void ComFree( LPVOID p) { if (p) pIMalloc->Free( p ); }

#define  _OPC_CHECK_HRFUNC(f) {HRESULT hr = f; if (FAILED( hr )) throw hr;} 
#define  _OPC_CHECK_HR(hr) {if (FAILED( hr )) throw hr;}
#define  _OPC_CHECK_PTR(p) {if ((p)== NULL) throw E_OUTOFMEMORY;}

//DOM-IGNORE-END

#endif //__UtilityDefs_H_
