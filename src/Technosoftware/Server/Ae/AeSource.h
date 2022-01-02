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

#ifndef __EventSource_H
#define __EventSource_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "WideString.h"                         // for WideString

class EventArea;
class AeCondition;
class AeEvent;
class AeEventArray;


//-----------------------------------------------------------------------
// CLASS
//-----------------------------------------------------------------------
class AeSource
{
// Construction / Destruction
public:
   AeSource();
   HRESULT Create( EventArea* pArea, LPCWSTR szName, BOOL fIsFullyQualified );
   ~AeSource();

// Attributes
public:
   inline WideString&     Name()            { return m_wsName; }
   inline LPCWSTR          PartialName()     { return m_pszPartialName; }

// Operations
public:
   HRESULT  AddToAdditionalArea( EventArea* pArea );

   HRESULT  AttachCondition( AeCondition* pCond );
   HRESULT  DetachCondition( AeCondition* pCond );
   HRESULT  QueryConditionNames(    DWORD    *  pdwCount,
                                    LPWSTR   ** ppszConditionNames );

   HRESULT  GetConditionState(
   /* [in] */                    LPWSTR         szConditionName,
   /* [in] */                    DWORD          dwNumEventAttrs,
   /* [size_is][in] */           DWORD       *  pdwAttributeIDs,
   /* [out] */                   OPCCONDITIONSTATE
                                             ** ppConditionState );

   HRESULT  AckCondition(
   /* [string][in] */            LPWSTR         szAcknowledgerID,
   /* [string][in] */            LPWSTR         szComment,
   /* [string][in] */            LPWSTR         szConditionName,
   /* [in] */                    FILETIME       ftActiveTime,
   /* [out] */                   AeEvent**     ppEvent );

   HRESULT  EnableConditions(
   /* [in] */                    BOOL           fEnable,
   /* [out] */                   AeEventArray* parEvents );

   HRESULT  GetAreaNames( LPVARIANT pvAreas );

// Implementation
protected:
   WideString                      m_wsName;   // The fully qualified name of the source
   LPCWSTR                          m_pszPartialName;
                                                // Points to the partially source name in
                                                // name_. Do not free this string !

                                                // This Source instance is member of this Areas
   CComAutoCriticalSection          m_csAreaRefs;
   CSimpleArray<EventArea*>        m_arAreaRefs;
   
   CComAutoCriticalSection          m_csCondRefs;
   CSimpleArray<AeCondition*>   m_arCondRefs;

   AeCondition* LookupCondition( LPCWSTR szName );
};
//DOM-IGNORE-END

#endif // __EventSource_H
