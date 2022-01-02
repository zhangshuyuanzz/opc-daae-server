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

#ifdef   _OPC_SRV_AE                            // Alarms & Events Server

 //DOM-IGNORE-BEGIN
 //-------------------------------------------------------------------------
 // INLCUDE
 //-------------------------------------------------------------------------
#include "stdafx.h"
#include "UtilityDefs.h"
#include "AeEvent.h"
#include "AeSource.h"
#include "AeCategory.h"
#include "AeCondition.h"
#include "AeAttribute.h"
#include "AeAttributeValueMap.h"
#include "CoreMain.h"

//-------------------------------------------------------------------------
// CODE AeCondition
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AeCondition::AeCondition()
{
    m_wQuality = OPC_QUALITY_GOOD;
    m_wChangeMask = 0;
    m_wNewState = OPC_CONDITION_ENABLED | OPC_CONDITION_ACKED;
    m_dwNumOfAttrs = 0;
    m_pavAttrValues = NULL;

    memset(&m_ftLastAckTime, 0, sizeof(m_ftLastAckTime));
    memset(&m_ftSubCondLastActive, 0, sizeof(m_ftSubCondLastActive));
    memset(&m_ftCondLastActive, 0, sizeof(m_ftCondLastActive));
    memset(&m_ftCondLastInactive, 0, sizeof(m_ftCondLastInactive));
    memset(&m_ftTime, 0, sizeof(m_ftTime));
}



//=========================================================================
// Initializer
// -----------
//    Must be called after construction.
//    Also the internal default attributes 'ACK COMMENT' and
//    'AREAS' are initalized.
//=========================================================================
HRESULT AeCondition::Create(DWORD dwCondID, AeConditionDefiniton* pCondDef, AeSource* pSource)
{
    m_dwCondID = dwCondID;
    m_pCondDef = pCondDef;
    m_pSource = pSource;

    HRESULT hres = m_wsAckComment.SetString(L"");
    if (FAILED(hres)) return hres;

    hres = m_wsAcknowledgerID.SetString(L"");
    if (FAILED(hres)) return hres;

    m_pActiveSubCond = pCondDef->DefaultSubCondition();
    m_dwSeverity = m_pActiveSubCond->Severity();

    hres = m_wsMessage.SetString(m_pActiveSubCond->Description());
    if (FAILED(hres)) return hres;

    m_dwNumOfAttrs = pCondDef->Category()->NumOfAttrs();
    m_pavAttrValues = new _variant_t[m_dwNumOfAttrs];
    if (!m_pavAttrValues) return E_OUTOFMEMORY;

    try {                                        // Initialize Internal handled attributes
        m_pavAttrValues[ATTRNDX_ACKCOMMENT] = m_wsAckComment;
        hres = m_pSource->GetAreaNames(&m_pavAttrValues[ATTRNDX_AREAS]);
        if (FAILED(hres)) return hres;
    }
    catch (_com_error &e) {                      // catch _variant_t operation errors
        return e.Error();
    }
    return S_OK;                                 // all succeeded
}



//=========================================================================
// Destructor
//=========================================================================
AeCondition::~AeCondition()
{
    if (m_pavAttrValues) {
        delete[] m_pavAttrValues;
    }
}


//-------------------------------------------------------------------------
// OPERATIONS
//-------------------------------------------------------------------------

//=========================================================================
// Enable
// ------
//    Enables or disables the condition.
//    The return code is S_FALSE and no event instance is created
//    if the condition is:
//       - already in the specified state
//       - disabled
//       - inactive and acked
//
// Parameters:
//    IN
//       fEnable              if TRUE the condition is enabled; otherwise
//                            disabled.
//    OUT
//       ppEvent              Created Event instance if return code
//                            is S_OK. Note: this function does not
//                            notifiy the clients if the condition
//                            is acknowledged.
// Return Code:
//    S_OK                    All succeeded
//    S_FALSE                 Condition is already in the specified state
//                            or disabled or 'inactive and acked'.
//                            No event instance is created.
//    E_XXX                   Error occured.
//=========================================================================
HRESULT AeCondition::Enable(BOOL fEnable, AeEvent** ppEvent)
{
    *ppEvent = NULL;
    if (IsEnabled() == fEnable) {
        return S_FALSE;                           // Already in the specified state
    }

    HRESULT  hres = S_FALSE;
    FILETIME ftNow;

    hres = CoFileTimeNow(&ftNow);
    if (FAILED(hres)) return hres;

    if (fEnable) {
        m_wNewState |= OPC_CONDITION_ENABLED;

        if (IsActive()) {
            m_ftCondLastInactive = ftNow;
            m_wNewState &= ~OPC_CONDITION_ACTIVE;
        }
        if (!IsAcked()) {
            hres = AcknowledgeByServer(L"Acknowledged by Server when Enabled", ppEvent);
            if (FAILED(hres)) return hres;
            if (hres == S_OK) {
                (*ppEvent)->Release();           // Release the event instance because no
                *ppEvent = NULL;                 // event must be generated.
                m_ftLastAckTime = ftNow;         // Guarantees that all changes have the same time
            }
        }
    }
    else {
        m_wNewState &= ~OPC_CONDITION_ENABLED;
    }
    hres = S_FALSE;                           // Return Code if no event is created.
                                              // Notification must be sent only if the
                                              // condition is enabled and the state is
                                              // 'Active' or 'Inactive and Unacked'.
    if (IsActive() || !IsAcked()) {
        m_ftTime = ftNow;                      // Time of the condition state transition is now
        m_wChangeMask = OPC_CHANGE_ENABLE_STATE;
        // Create an event instance for client notification
        hres = CreateEventInstance(ppEvent);
    }
    return hres;
}



//=========================================================================
// AcknowledgeByServer                                               INLINE
// -------------------
//    The conditin is acknowledged automatically by the server.
//
// Parameters:
//    IN
//       szComment            An optional text string or NULL-Pointer
//                            if there is no comment.
//    OUT
//       ppEvent              Created Event instance if return code
//                            is S_OK. Note: this function does not
//                            notifiy the clients if the condition
//                            is acknowledged.
//=========================================================================
HRESULT AeCondition::AcknowledgeByServer(LPCWSTR szComment, AeEvent** ppEvent, BOOL useCurrentTime /* = TRUE */)
{
    return Acknowledge(L"", szComment ? szComment : L"", m_ftCondLastActive, ppEvent, useCurrentTime);
}



//=========================================================================
// Acknowledge
// -----------
//    Acknowledges the condition.
//
// Parameters:
//    IN
//       szAcknowledgerID     Identifying who is acknowledging the
//                            condition. It's a NULL-String if
//                            acknowledged by Server.
//       szComment            An optional text string. Pass a NULL-String
//                            if there is no comment.
//       ftActiveTime         Identifies the condition by time.
//    OUT
//       ppEvent              Created Event instance if return code
//                            is S_OK. Note: this function does not
//                            notify the clients if the condition
//                            is acknowledged.
//=========================================================================
HRESULT AeCondition::Acknowledge(LPCWSTR szAcknowledgerID, LPCWSTR szComment, FILETIME ftActiveTime, AeEvent** ppEvent, BOOL useCurrentTime /* = TRUE */)
{
    HRESULT hres = S_OK;

    _ASSERTE(szComment);                        // Must not be NULL
    *ppEvent = NULL;
    // Checks if this condition has to be acknowledged
    if (CompareFileTime(&ftActiveTime, &m_ftCondLastActive) != 0)
        return OPC_E_INVALIDTIME;
    if (IsAcked()) return OPC_S_ALREADYACKED;   // Already acknowledged
    if (!IsEnabled()) return E_INVALIDARG;      // The condition is unacknowledged but disabled


    if (useCurrentTime) {
        hres = CoFileTimeNow(&m_ftTime);        // Time of the event occurence
        if (FAILED(hres)) return hres;
    }
    m_ftLastAckTime = m_ftTime;                 // Time of acknowledge

    hres = m_wsAcknowledgerID.SetString(szAcknowledgerID);
    if (FAILED(hres)) return hres;

    m_wChangeMask = 0;                          // Reset the change mask

    if (*szComment) {                           // Only if comment is allowed
        hres = m_wsAckComment.SetString(szComment);
        if (FAILED(hres)) return hres;
        // Update the 'ACK COMMENT' attribute if changed
        try {
            if (wcscmp(V_BSTR(&m_pavAttrValues[ATTRNDX_ACKCOMMENT]), szComment) != 0) {
                m_pavAttrValues[ATTRNDX_ACKCOMMENT] = szComment;
                m_wChangeMask |= OPC_CHANGE_ATTRIBUTE;
            }
        }
        catch (_com_error &e) {                 // Catch _variant_t operation errors
            return e.Error();
        }
    }

    m_wNewState |= OPC_CONDITION_ACKED;
    m_wChangeMask |= OPC_CHANGE_ACK_STATE;

    // Acknowledge Notification
    AeBaseServer* pServerHandler = ::OnGetAeBaseServer();
    _ASSERTE(pServerHandler != NULL);

    DWORD dwSubCondID;
    hres = m_pCondDef->GetSubCondDefID(m_pActiveSubCond, &dwSubCondID);
    if (FAILED(hres)) return hres;

    hres = pServerHandler->OnAcknowledgeNotification(m_dwCondID, dwSubCondID);
    if (FAILED(hres)) return hres;

    return CreateEventInstance(ppEvent);       // Create an event instance for client notification
}



//=========================================================================
// GetState
// --------
//    Gets the current stat of the condition.
//    Memory is allocated from the Global COM Memory Manager.
//=========================================================================
HRESULT AeCondition::GetState(DWORD dwNumEventAttrs, DWORD* pdwAttributeIDs,
    OPCCONDITIONSTATE** pState)
{
    HRESULT  hres = S_OK;
    DWORD    i;

    *pState = NULL;

    OPCCONDITIONSTATE* pS = NULL;

    try {
        pS = ComAlloc<OPCCONDITIONSTATE>();
        if (!pS) throw E_OUTOFMEMORY;

        memset(pS, 0, sizeof(OPCCONDITIONSTATE));

        pS->wState = m_wNewState;

        if (IsActive()) {
            pS->szActiveSubCondition = m_pActiveSubCond->Name().CopyCOM();
            pS->szASCDefinition = m_pActiveSubCond->Definition().CopyCOM();
            pS->dwASCSeverity = m_pActiveSubCond->Severity();
            pS->szASCDescription = m_pActiveSubCond->Description().CopyCOM();

            if (!pS->szActiveSubCondition || !pS->szASCDefinition || !pS->szASCDescription) {
                throw E_OUTOFMEMORY;
            }
        }
        else {
            // According to OPC AE specification 1.03 we have to return
            // NULL pointers and not NUL strings if the Event Condition is inactive.
            pS->szActiveSubCondition = NULL;
            pS->szASCDefinition = NULL;
            pS->dwASCSeverity = 0;
            pS->szASCDescription = NULL;
        }

        pS->wQuality = m_wQuality;
        pS->ftLastAckTime = m_ftLastAckTime;
        pS->ftSubCondLastActive = m_ftSubCondLastActive;
        pS->ftCondLastActive = m_ftCondLastActive;
        pS->ftCondLastInactive = m_ftCondLastInactive;

        // 22-jan-2002 MT
        // According to OPC AE specification 1.03 we have to return
        // a NULL pointer and not a NUL string if there is no ID.
        // Don't copy the string if the string is empty.
        if (wcslen(m_wsAcknowledgerID)) {
            pS->szAcknowledgerID = m_wsAcknowledgerID.CopyCOM();
            if (!pS->szAcknowledgerID) throw E_OUTOFMEMORY;
        }

        // 22-jan-2002 MT
        // According to OPC AE specification 1.03 we have to return
        // a NULL pointer and not a NUL string if there is no ID
        // Don't copy the string if the string is empty.
        if (wcslen(m_wsAckComment)) {
            pS->szComment = m_wsAckComment.CopyCOM();
            if (!pS->szComment) throw E_OUTOFMEMORY;
        }

        // Sub Condition Info
        hres = m_pCondDef->GetSubConditionInfo(
            &pS->dwNumSCs,
            &pS->pszSCNames,
            &pS->pszSCDefinitions,
            &pS->pdwSCSeverities,
            &pS->pszSCDescriptions);
        if (FAILED(hres)) throw hres;

        // Event Attributes
        pS->dwNumEventAttrs = dwNumEventAttrs;

        if (!(pS->pEventAttributes = ComAlloc<VARIANT>(dwNumEventAttrs)))
            throw E_OUTOFMEMORY;

        for (i = 0; i < dwNumEventAttrs; i++) {
            VariantInit(&pS->pEventAttributes[i]);
        }

        if (!(pS->pErrors = ComAlloc<HRESULT>(dwNumEventAttrs)))
            throw E_OUTOFMEMORY;
        // Get ID array from condition definition
        DWORD*   pdwAttrIDs = NULL;
        DWORD    dwCount;
        hres = m_pCondDef->Category()->GetAttributeIDs(&dwCount, &pdwAttrIDs);
        if (FAILED(hres)) throw hres;

        for (i = 0; i < dwNumEventAttrs; i++) {
            pS->pErrors[i] = E_FAIL;
            // Search specified ID and copy value if attribute exist
            for (DWORD idcnt = 0; idcnt < dwCount; idcnt++) {
                if (pdwAttrIDs[idcnt] == pdwAttributeIDs[i]) {
                    pS->pErrors[i] = VariantCopy(&pS->pEventAttributes[i], &m_pavAttrValues[idcnt]);
                }
            }
        }
        delete[] pdwAttrIDs;                     // No longer used
        *pState = pS;                             // All succeeded
    }
    catch (HRESULT hresEx) {
        if (pS) {
            if (pS->szActiveSubCondition) pIMalloc->Free(pS->szActiveSubCondition);
            if (pS->szASCDefinition)      pIMalloc->Free(pS->szASCDefinition);
            if (pS->szASCDescription)     pIMalloc->Free(pS->szASCDescription);
            if (pS->szAcknowledgerID)     pIMalloc->Free(pS->szAcknowledgerID);
            if (pS->szComment)            pIMalloc->Free(pS->szComment);
            if (pS->pErrors)              pIMalloc->Free(pS->pErrors);

            // Note :   pszSCNames, pszSCDefinitions,
            //          pdwSCSeverities, pszSCDescriptions are
            //          handled by function GetSubConditionInfo().

            if (pS->pEventAttributes) {
                for (i = 0; i < dwNumEventAttrs; i++) {
                    VariantClear(&pS->pEventAttributes[i]);
                }
                pIMalloc->Free(pS->pEventAttributes);
            }
            pIMalloc->Free(pS);
        }
        *pState = NULL;
        hres = hresEx;
    }

    return hres;
}



//=========================================================================
// ChangeState
// -----------
//    Changes one or more states of the codition. If at least one state
//    has changed and all succeeded then an event instance is created
//    and returned.
//
// Parameters:
//    IN
//       AeConditionChangeStates reference with the used attributes
//
//       ActiveState()        If TRUE the condition becomes active;
//                            otherwise inactive.
//       SubCondID()          ID of the current sub condition which
//                            is active. This parameter is ignored if
//                            the condition becomes inactive.
//       Quality()            Quality of the condition
//       AttrValuesPtr()      Array of attribute values. The size must be
//                            identical with the number of defined
//                            attributes of the associated condition
//                            definition.
//
//       DEFAULT OVERRIDES
//          If the following pointers are NULL-pointers then the default
//          values from the associated condition definition are used.
//
//       Message()            Event notification message
//       SeverityPtr()        Severity
//       AckRequiredPtr()     If TRUE then the event must be
//                            acknowledged; otherwise acknowledgement is
//                            not required. This parameter is ignored if
//                            the condition becomes inactive.
//       TimeStampPtr()       Time of the condition state transition.
//    OUT
//       ppEvent              Created Event instance if return code
//                            is S_OK.
//
// Return Code:
//    S_OK                    All succeeded
//    S_FALSE                 No condition states has changed or condition
//                            is disabled.
//                            No event instance is created.
//    E_INVALIDARG            Invalid number of attribute values.
//    E_XXX                   Error occured.
//=========================================================================
HRESULT AeCondition::ChangeState(AeConditionChangeStates& cs, AeEvent** ppEvent, BOOL useCurrentTime /* = TRUE */)
{
    *ppEvent = NULL;                             // Initialize out-parameter

                                                 // The number of specified attribute values
                                                 // must be identical with the number of defined
                                                 // attributes of the associated condition definition.
    if (cs.AttrCount() != (m_pCondDef->Category()->NumOfAttrs() -
        m_pCondDef->Category()->NumOfInternalAttrs())) {
        return E_INVALIDARG;
    }

    HRESULT hres = S_OK;
    if (cs.TimeStampPtr()) {
        m_ftTime = *cs.TimeStampPtr();            // Time of the condition state transition is already defined
    }
    else {
        hres = CoFileTimeNow(&m_ftTime);        // Time of the condition state transition is now
        if (FAILED(hres)) return hres;
    }

    m_wChangeMask = 0;                           // Reset the change mask

    //
    // Active Sub Condition
    //
    if (m_pCondDef->IsMultiState() && cs.ActiveState()) {
        AeSubConditionDefiniton* pNewActiveSubCond = m_pCondDef->SubCondition(cs.SubCondID());
        if (!pNewActiveSubCond) {                 // There is no sub condition definition with the specified ID
            return E_INVALID_HANDLE;
        }
        if (m_pActiveSubCond != pNewActiveSubCond) {
            m_ftSubCondLastActive = m_ftTime;
            m_pActiveSubCond = pNewActiveSubCond;  // Transition into new sub condition
            m_wChangeMask |= OPC_CHANGE_SUBCONDITION;
        }
    }

    //
    // Active State
    //
    if (IsActive() != cs.ActiveState()) {
        if (cs.ActiveState()) {
            m_ftCondLastActive = m_ftTime;
            m_ftSubCondLastActive = m_ftTime;
            m_wNewState |= OPC_CONDITION_ACTIVE;
        }
        else {
            m_ftCondLastInactive = m_ftTime;
            m_wNewState &= ~OPC_CONDITION_ACTIVE;
        }
        m_wChangeMask |= OPC_CHANGE_ACTIVE_STATE;
    }

    //
    // Quality
    //
    if (m_wQuality != cs.Quality()) {
        m_wChangeMask |= OPC_CHANGE_QUALITY;
        m_wQuality = cs.Quality();
    }

    //
    // Attribute Values
    //
    if (cs.AttrValuesPtr()) {
        try {                                     // Initialize Internal handled attributes

            for (DWORD i = m_pCondDef->Category()->NumOfInternalAttrs(), z = 0; i < m_dwNumOfAttrs; i++, z++) {
                if (m_pavAttrValues[i] != cs.AttrValuesPtr()[z]) {
                    m_pavAttrValues[i] = cs.AttrValuesPtr()[z];
                    m_wChangeMask |= OPC_CHANGE_ATTRIBUTE;
                }
            }
        }
        catch (_com_error &e) {                   // Catch _variant_t operation errors
            return e.Error();
        }
    }

    // The optional parameters may override the default values

    //
    // Message
    //
    LPCWSTR szNewMessage = cs.Message() ? cs.Message() : m_pActiveSubCond->Description();
    if (wcscmp(m_wsMessage, szNewMessage)) {
        hres = m_wsMessage.SetString(szNewMessage);
        if (FAILED(hres)) return hres;
        m_wChangeMask |= OPC_CHANGE_MESSAGE;
    }

    //
    // Severity
    //
    DWORD dwNewSeverity = cs.SeverityPtr() ? *cs.SeverityPtr() : m_pActiveSubCond->Severity();
    if (dwNewSeverity < 1 || dwNewSeverity > 1000) return E_INVALIDARG;
    if (m_dwSeverity != dwNewSeverity) {
        m_wChangeMask |= OPC_CHANGE_SEVERITY;
        m_dwSeverity = dwNewSeverity;
    }

    //
    // Acknowledge Required Flag
    //

    BOOL fNewAckRequired = FALSE;

    if (IsActive()) {

        fNewAckRequired = cs.AckRequiredPtr() ? *cs.AckRequiredPtr() : m_pActiveSubCond->AckRequired();

        if (!IsAcked() && !fNewAckRequired) {
            WORD wCM = m_wChangeMask;              // Save change mask
                                                   // Note: AcknowledgeByServer resets the change mask.

            hres = AcknowledgeByServer(L"Acknowledge No Longer Required", ppEvent, useCurrentTime);
            if (FAILED(hres)) return hres;
            // Restore change mask
            m_wChangeMask = (*ppEvent)->wChangeMask | wCM;
            (*ppEvent)->Release();
            *ppEvent = NULL;

        }
    }
    // Conditions must be acknowledged if required before new
    // states can be set.
    if (fNewAckRequired && IsAcked()) {
        m_wNewState &= ~OPC_CONDITION_ACKED;      // New state is not acknowledged
        m_wChangeMask |= OPC_CHANGE_ACK_STATE;
    }

    if (!(IsEnabled() && m_wChangeMask))
        return S_FALSE;                           // Don't generate an event because the condition is disabled
                                                  // or nothing has changed

    hres = CreateEventInstance(ppEvent);       // Create an event instance for client notification
    return hres;
}



//=========================================================================
// CreateEventInstance
// -------------------
//    Creates an event instance with the current values of the
//    condition.
//
// Parameters:
//    OUT
//       ppEvent              Created Event instance if all succeeded.
//=========================================================================
HRESULT AeCondition::CreateEventInstance(AeEvent** ppEvent)
{
    HRESULT     hres = S_OK;
    AeEvent*   pE = NULL;
    DWORD*      pdwAttrIDs = NULL;

    *ppEvent = NULL;

    try {
        pE = new AeEvent;                        // Create an event instance
        if (!pE) throw E_OUTOFMEMORY;

        pE->wChangeMask = m_wChangeMask;
        pE->wNewState = m_wNewState;
        pE->szSource = m_pSource->Name().Copy();
        pE->ftTime = m_ftTime;
        pE->szMessage = m_wsMessage.Copy();
        pE->dwEventType = OPC_CONDITION_EVENT;
        pE->dwEventCategory = m_pCondDef->Category()->CatID();
        pE->dwSeverity = m_dwSeverity;
        pE->szConditionName = m_pCondDef->Name().Copy();
        pE->szSubconditionName = m_pActiveSubCond->Name().Copy();
        pE->wQuality = m_wQuality;
        pE->bAckRequired = IsAcked() ? FALSE : TRUE;
        pE->ftActiveTime = m_ftCondLastActive;
        pE->dwCookie = (DWORD)this;    // Cookie must be unique
        pE->szActorID = m_wsAcknowledgerID.Copy();
        pE->wReserved = 0;

        if (!pE->szConditionName || !pE->szSubconditionName ||
            !pE->szActorID || !pE->szSource || !pE->szMessage) {
            throw E_OUTOFMEMORY;
        }

        // Add all attribute IDs and values 
        // The IDs are from the specified category and the current values from the user.
        // The values must be in the same order and have the same types as specified by
        // the category.

        DWORD dwCount;

        hres = m_pCondDef->Category()->GetAttributeIDs(&dwCount, &pdwAttrIDs);
        if (FAILED(hres)) throw hres;

        hres = pE->m_mapAttrValues.Create(dwCount);
        if (FAILED(hres)) throw hres;

        for (DWORD i = 0; i < dwCount; i++) {
            hres = pE->m_mapAttrValues.SetAtIndex(i, pdwAttrIDs[i], &m_pavAttrValues[i]);
            if (FAILED(hres)) throw hres;
        }
    }
    catch (HRESULT hresEx) {
        hres = hresEx;                            // catch own exceptions
    }
    catch (...) {                                 // catch all other exception
        hres = E_FAIL;                            // e.g. if the attribute value array is invalid
    }

    if (FAILED(hres)) {
        if (pE) {
            pE->Release();
        }
    }
    else {
        *ppEvent = pE;
    }

    if (pdwAttrIDs)                              // Delete the temporary array of Attributes IDs.
        delete[] pdwAttrIDs;

    return hres;
}



//=========================================================================
// IsAttachedTo
// ------------
//    Checks if this condition is attached to the specified 
//    source/condition/sub condition/category.
//
// Parameters:
//    IN
//       szSource
//       dwEventCategory
//       szConditionName
//       szSubConditionName
//    OUT
//       pdwSubCondDefID      The ID of the sub condition specified
//                            by name.
//       ppCat                Pointer to the Event Category specified
//                            by ID.
//
// Return Code:
//    TRUE if this condition is attached to the specified
//    source/condition/sub condition/category; otherwise FALSE.
//=========================================================================
BOOL AeCondition::IsAttachedTo(LPCWSTR szSource, DWORD dwEventCategory,
    LPCWSTR szConditionName, LPCWSTR szSubConditionName,
    LPDWORD pdwSubCondDefID, AeCategory** ppCat)
{
    if (m_pCondDef->Category()->CatID() != dwEventCategory)     return FALSE;
    if (wcscmp(m_pSource->Name(), szSource))                  return FALSE;
    if (wcscmp(m_pCondDef->Name(), szConditionName))          return FALSE;
    if (FAILED(m_pCondDef->GetSubCondDefID(szSubConditionName, pdwSubCondDefID))) {
        return FALSE;
    }
    *ppCat = m_pCondDef->Category();
    return TRUE;
}



//=========================================================================
// IsAttachedTo
// ------------
//    Checks if this condition is attached to the specified 
//    source/condition.
//
// Parameters:
//    IN
//       szSource
//       szConditionName
//
// Return Code:
//    TRUE if this condition is attached to the specified
//    source/condition; otherwise FALSE.
//=========================================================================
BOOL AeCondition::IsAttachedTo(LPCWSTR szSource, LPCWSTR szConditionName)
{
    if (wcscmp(m_pSource->Name(), szSource))            return FALSE;
    if (wcscmp(m_pCondDef->Name(), szConditionName))    return FALSE;
    return TRUE;
}
//DOM-IGNORE-END
#endif