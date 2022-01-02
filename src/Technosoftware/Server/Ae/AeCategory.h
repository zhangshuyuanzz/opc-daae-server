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

#ifndef __EVENTCATEGORY_H
#define __EVENTCATEGORY_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "WideString.h"                         // for WideString

class AeAttribute;
class AeConditionDefiniton;


//-----------------------------------------------------------------------
// CLASS
//-----------------------------------------------------------------------
class AeCategory
{
// Construction / Destruction
public:
   AeCategory();                            // Default Constructor
   HRESULT Create( DWORD dwID, LPCWSTR szDescr, DWORD dwEventType );
   ~AeCategory();                           // Destructor

// Attributes
public:
   inline DWORD         CatID()     const { return m_dwID; }
   inline WideString&  Descr()           { return m_wsDescr; }
   inline DWORD         EventType() const { return m_dwEventType; }
   inline DWORD         NumOfAttrs()      {  m_csAttrMap.Lock();
                                             DWORD dwNum = m_mapAttributes.GetSize();
                                             m_csAttrMap.Unlock();
                                             return dwNum;
                                          }
   inline DWORD         NumOfInternalAttrs()
                                    const { return m_dwNumOfInternalAttrs; }


// Operations
public:
   HRESULT  AddEventAttribute( AeAttribute* pAttr );
   BOOL     ExistEventAttribute( DWORD dwAttrID );
   BOOL     ExistEventAttribute( LPCWSTR szAttrDescr );
   HRESULT  QueryEventAttributes(   DWORD    *  pdwCount,
                                    DWORD    ** ppdwAttrIDs,
                                    LPWSTR   ** ppszAttrDescs,
                                    VARTYPE  ** ppvtAttrTypes );
   HRESULT  GetAttributeIDs( DWORD* pdwCount, DWORD** ppdwAttrIDs );

   HRESULT  AttachConditionDef( AeConditionDefiniton* pCondDef );
   HRESULT  DetachConditionDef( AeConditionDefiniton* pCondDef );
   HRESULT  QueryConditionNames(    DWORD    *  pdwCount,
                                    LPWSTR   ** ppszConditionNames );

// Implementation
protected:
   DWORD                               m_dwID;        // Event Category ID
   WideString                         m_wsDescr;     // Event Category Description
   DWORD                               m_dwEventType; // Associated Event Type

                                                      // The number of internal handled
                                                      // attributes in m_mapAttributes
   DWORD                               m_dwNumOfInternalAttrs;

   CComAutoCriticalSection             m_csAttrMap;
   CSimpleMap<DWORD, AeAttribute*> m_mapAttributes;

   CComAutoCriticalSection             m_csCondDefRefs;
   CSimpleArray<AeConditionDefiniton*>   m_arCondDefRefs;
};
//DOM-IGNORE-END

#endif // __EVENTCATEGORY_H

