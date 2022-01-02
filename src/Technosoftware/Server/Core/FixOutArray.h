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
#ifndef __FIXOUTARRAY_H_
#define __FIXOUTARRAY_H_

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000


// Default template to initialize/destroy members of the class template CFixOutArray.
template <class T> class _FxOElem {
public:
   inline void Init(T* p,BOOL f,T IVal) { *p=IVal; }
   inline void Destroy(T*, BOOL f) {}
};


// Initialize/destroy VARIANT members of the class template CFixOutArray.
template <> class _FxOElem<VARIANT> {
public:
   inline void Init(VARIANT* p, BOOL f, VARIANT IVal) {VariantInit(p); }
   inline void Destroy(VARIANT* p, BOOL f) { VariantClear(p); }
};

// Initialize/destroy LPWSTR members of the class template CFixOutArray.
template <> class _FxOElem<LPWSTR> {
public:
   inline void Init(LPWSTR* p, BOOL f, LPWSTR IVal) { *p=NULL; }
   inline void Destroy(LPWSTR* p, BOOL f)
            {
               if(*p) {
                  f ? pIMalloc->Free(*p) : delete *p;
                  *p=NULL;
               }
            }
};



//-------------------------------------------------------------------------
// CLASS TEMPLATE CFixOutArray
//
//    T           Type of an array element
//    fCOMMemory  TRUE use memory from the COM Memory Manager
//                FALSE use heap memory
//-------------------------------------------------------------------------
template <class T, BOOL fCOMMemory = TRUE> class CFixOutArray
{
// Construction
public :
   CFixOutArray() : m_ppOut( NULL ), m_dwNum( 0 ) {};
   ~CFixOutArray() {};

// Operations
   //----------------------------------------------------------------------
   // Initialize Array
   // ----------------
   //    dwNum    Number of elements
   //    ppOut    Address of a pointer to store the allocated buffer.
   //    Init     Optional initialization value for an element
   //----------------------------------------------------------------------
   void Init( DWORD dwNum, T** ppOut, T Init = T() )
      {
         *ppOut = fCOMMemory ?
                     (T *)pIMalloc->Alloc( dwNum * sizeof (T) ) :
                     new T[ dwNum ];

         if (*ppOut == NULL) {
            throw E_OUTOFMEMORY;
         }
         _FxOElem<T>  Elem;
         for (DWORD i=0; i<dwNum; i++) {
            Elem.Init( &(*ppOut)[ i ], fCOMMemory, Init );
         }
         m_ppOut = ppOut;
         m_dwNum = dwNum;
      }

   //----------------------------------------------------------------------
   // Release array and all elements
   //----------------------------------------------------------------------
   void Cleanup()
      {
         if (m_ppOut) {                         // Release existing array
            _FxOElem<T> Elem;
            while (m_dwNum--) {
               Elem.Destroy( &(*m_ppOut)[ m_dwNum ], fCOMMemory );
            }
            fCOMMemory ? 
                  pIMalloc->Free( *m_ppOut ) :
                  delete [] *m_ppOut;
            *m_ppOut = NULL;
         }
      }

   //----------------------------------------------------------------------
   // Access to an element of the array
   //----------------------------------------------------------------------
   inline T & operator[]( DWORD dwNdx )
         {
            _ASSERTE( m_ppOut );                // Must be initialized
            _ASSERTE( dwNdx < m_dwNum );        // Must be in range
            return (*m_ppOut)[ dwNdx ];
         }

// Implementation
protected:
   T**   m_ppOut;
   DWORD m_dwNum;
};
//DOM-IGNORE-END

#endif // if not defined __FIXOUTARRAY_H
