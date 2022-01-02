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

#ifndef __OnEvent_H
#define __OnEvent_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "AeAttributeValueMap.h"
#include "AeCondition.h"
#include "UtilityDefs.h"                        // for CRefClass

class AeSource;
class AeCategory;


//-----------------------------------------------------------------------
// CLASS AeEvent
//-----------------------------------------------------------------------
class AeEvent : public ONEVENTSTRUCT, public CRefClass
{                                      
                                                // For initialization of the attribute value map
   friend HRESULT AeCondition::CreateEventInstance( AeEvent** ppEvent );

// Construction
public:
   AeEvent();

   // Create simple or tracking event
   HRESULT Create( AeCategory* pCat, AeSource *pSource,
                   LPCWSTR szParMessage, DWORD dwParSeverity, LPCWSTR szParActorID,
                   LPVARIANT pvAttrValues, LPFILETIME pft );

   // Note:
   //    Condition related events are craeted by
   //    AeCondition::CreateEventInstance().

// Destruction
private:                                        
   ~AeEvent() { Cleanup(); }                   // Destructor only called by Release()

// Operations
public:
   inline LPVARIANT LookupAttributeValue( DWORD dwAttrID )
         { return m_mapAttrValues.Lookup( dwAttrID );}

// Implementation
protected:
   AttributeValueMap m_mapAttrValues;               // Copy of all current attribute values
   void Cleanup();
};



//-----------------------------------------------------------------------
// CLASS AeEventArray
//-----------------------------------------------------------------------
class AeEventArray
{
// Construction
public:
   AeEventArray();
   HRESULT  Create( DWORD dwMax );

// Destruction
public:
   ~AeEventArray();

// Attributes
public:
   inline DWORD NumOfEvents() const { return m_dwNum; }
   inline AeEvent** EventPtrArray() const { return m_parOnEvent; }

// Operations
public:
   HRESULT  Add( AeEvent* pEvent );

// Implementation
protected:
   DWORD       m_dwMax;
   DWORD       m_dwNum;
   AeEvent**  m_parOnEvent;
};



//-----------------------------------------------------------------------
// CLASS AeSubscribedEvent
//-----------------------------------------------------------------------
class AeSubscribedEvent : public ONEVENTSTRUCT
{
// Construction
public:
   AeSubscribedEvent();
   HRESULT Create( AeEvent* pOnEvent, DWORD dwNumOfAttrIDs, DWORD dwAttrIDs[] );

// Destruction
public:
   ~AeSubscribedEvent();
};
//DOM-IGNORE-END

#endif // __OnEvent_H

