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

#ifndef __EVENTCONDITION_H
#define __EVENTCONDITION_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "WideString.h"                         // for WideString
#include "AeConditionDefinition.h"

class AeSource;
class AeEvent;
class AttributeValueMap;
class AeConditionChangeStates;
class _variant_t;


//-----------------------------------------------------------------------
// CLASS AeCondition
//-----------------------------------------------------------------------
class AeCondition {
// Construction
public:
   AeCondition();                           // Default Constructor
                                                // Initializer
   HRESULT Create( DWORD dwCondID, AeConditionDefiniton* pCondDef, AeSource* pSource );

// Destruction
public:
   ~AeCondition();

// Attributes
public:
   inline WideString&  Name() { return m_pCondDef->Name(); }
   inline BOOL IsActive() const { return (m_wNewState & OPC_CONDITION_ACTIVE) ? TRUE : FALSE; }
   inline BOOL IsEnabled() const { return (m_wNewState & OPC_CONDITION_ENABLED) ? TRUE : FALSE; }
   inline BOOL IsAcked() const { return (m_wNewState & OPC_CONDITION_ACKED) ? TRUE : FALSE; }

// Operations
public:
   HRESULT  Enable( BOOL fEnable, AeEvent** ppEvent );
   HRESULT  AcknowledgeByServer( LPCWSTR szComment, AeEvent** ppEvent, BOOL useCurrentTime = TRUE);
   HRESULT  Acknowledge( LPCWSTR AcknowledgerID, LPCWSTR szComment, FILETIME ftActiveTime, AeEvent** ppEvent, BOOL useCurrentTime = TRUE);
   HRESULT  GetState( DWORD dwNumEventAttrs, DWORD* pdwAttributeIDs, OPCCONDITIONSTATE** pState );
   HRESULT  CreateEventInstance( AeEvent** ppEvent );
   HRESULT  ChangeState( AeConditionChangeStates& cs, AeEvent** ppEvent, BOOL useCurrentTime = TRUE);
   BOOL     IsAttachedTo( LPCWSTR szSource, LPCWSTR szConditionName );
   BOOL     IsAttachedTo( LPCWSTR szSource, DWORD dwEventCategory,
                          LPCWSTR szConditionName, LPCWSTR szSubConditionName,
                          LPDWORD pdwSubCondDefID, AeCategory** ppCat );

// Implementation
protected:
   // Data Members
   // It is not required to protect the data members of this class
   // with ciritical sections beacause the access of the condition
   // map is protected.
   WORD        m_wQuality;                // Current Quality
   WORD        m_wChangeMask;
   WORD        m_wNewState;               // Condition State : active, enabled, acked
   DWORD       m_dwSeverity;              // Current Severity
   BOOL        m_fAckRequired;            // Current acknowledge requirement
   FILETIME    m_ftLastAckTime;
   FILETIME    m_ftSubCondLastActive;
   FILETIME    m_ftCondLastActive;
   FILETIME    m_ftCondLastInactive;
   FILETIME    m_ftTime;                  // time that the condition transitioned
                                          // into the new state or sub condition.
   WideString m_wsAcknowledgerID;
   WideString m_wsAckComment;
   WideString m_wsMessage;               // Current Message

   DWORD                   m_dwCondID;
   AeSource*           m_pSource;
   AeConditionDefiniton*     m_pCondDef;

   DWORD                   m_dwNumOfAttrs;
   _variant_t*             m_pavAttrValues;

   AeSubConditionDefiniton*  m_pActiveSubCond;
};
//DOM-IGNORE-END

#endif // __EVENTCONDITION_H

