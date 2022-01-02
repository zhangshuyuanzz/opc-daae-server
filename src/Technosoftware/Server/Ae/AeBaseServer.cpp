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
 //-----------------------------------------------------------------------
 // INLCUDE
 //-----------------------------------------------------------------------
#include "stdafx.h"
#include "AeBaseServer.h"
#include "FixOutArray.h"
#include "AeCategory.h"
#include "AeAttribute.h"
#include "AeComSubscriptionManager.h"
#include "AeEvent.h"
#include "MatchPattern.h"

//-------------------------------------------------------------------------
// CODE AeConditionChangeStates
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AeConditionChangeStates::AeConditionChangeStates()
{
    memset(this, 0, sizeof(AeConditionChangeStates));
}

AeConditionChangeStates::AeConditionChangeStates(const AeConditionChangeStates& cs)
{
    memcpy(this, &cs, sizeof(cs));
}

AeConditionChangeStates::AeConditionChangeStates(
    // mandatory
    DWORD       dwCondID,
    DWORD       dwSubCondID,
    BOOL        fActiveState,
    WORD        wQuality,
    // optional
    DWORD       dwAttrCount,   /* = 0    */
    LPVARIANT   pvAttrValues,  /* = NULL */
    LPCWSTR     szMessage,     /* = NULL */
    DWORD*      pdwSeverity,   /* = NULL */
    LPBOOL      pfAckRequired, /* = NULL */
    LPFILETIME  pftTimeStamp   /* = NULL */
)
{
    m_dwCondID = dwCondID;
    m_dwSubCondID = dwSubCondID;
    m_fActiveState = fActiveState;
    m_wQuality = wQuality;
    m_dwAttrCount = dwAttrCount;
    m_pvAttrValues = pvAttrValues;
    m_szMessage = szMessage;
    m_pdwSeverity = pdwSeverity;
    m_pfAckRequired = pfAckRequired;
    m_pftTimeStamp = pftTimeStamp;
}



//-------------------------------------------------------------------------
// CODE AeBaseServer
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AeBaseServer::AeBaseServer()
{
}



//=========================================================================
// Initializer
// -----------
//    Must be called after construction.
//=========================================================================
HRESULT AeBaseServer::Create()
{
    return m_RootArea.Create(NULL, AREAID_ROOT, L"Root");
}



//=========================================================================
// Destructor
//=========================================================================
AeBaseServer::~AeBaseServer()
{
    int i;

    // Cleanup Condition Map
    {
        AeCondition* pCond;
        m_csCondMap.Lock();
        for (i = 0; i < m_mapConditions.GetSize(); i++) {
            pCond = m_mapConditions.GetValueAt(i);
            if (pCond) delete pCond;
        }
        m_mapConditions.RemoveAll();
        m_csCondMap.Unlock();
    }
    // Cleanup Condition Definition Map
    {
        AeConditionDefiniton* pCondDef;
        m_csCondDefMap.Lock();
        for (i = 0; i < m_mapConditionDefs.GetSize(); i++) {
            pCondDef = m_mapConditionDefs.GetValueAt(i);
            if (pCondDef) delete pCondDef;
        }
        m_mapConditionDefs.RemoveAll();
        m_csCondDefMap.Unlock();
    }
    // Cleanup Source Map
    {
        AeSource* pSrc;
        m_csSrcMap.Lock();
        for (i = 0; i < m_mapSources.GetSize(); i++) {
            pSrc = m_mapSources.GetValueAt(i);
            if (pSrc) delete pSrc;
        }
        m_mapSources.RemoveAll();
        m_csSrcMap.Unlock();
    }
    // Cleanup Event Category Map
    {
        AeCategory* pCat;
        m_csCatMap.Lock();
        for (i = 0; i < m_mapCategories.GetSize(); i++) {
            pCat = m_mapCategories.GetValueAt(i);
            if (pCat) delete pCat;
        }
        m_mapCategories.RemoveAll();
        m_csCatMap.Unlock();
    }
}



//-------------------------------------------------------------------------
// OPERATIONS
//-------------------------------------------------------------------------

//=========================================================================
// AddEventCategory
// ----------------
//    Adds a new category to the event server.
//    The ID of the event category must be unique.
//=========================================================================
HRESULT AeBaseServer::AddEventCategory(DWORD dwCatID, LPCWSTR szDescr, DWORD dwEventType)
{
    _ASSERTE(szDescr != NULL);                 // Must not be NULL
    _ASSERTE(dwEventType == OPC_SIMPLE_EVENT || // Must be one of these types
        dwEventType == OPC_TRACKING_EVENT ||
        dwEventType == OPC_CONDITION_EVENT);

    HRESULT hres = S_OK;
    AeCategory* pCat = NULL;

    m_csCatMap.Lock();
    try {
        if (m_mapCategories.Lookup(dwCatID))    // ID must be unique
            throw E_OBJECT_IN_LIST;

        pCat = new AeCategory;                // Category ID is unique. Create a new instance.
        if (!pCat) throw E_OUTOFMEMORY;
        // Initialize new instance
        hres = pCat->Create(dwCatID, szDescr, dwEventType);
        if (FAILED(hres)) throw hres;

        if (!m_mapCategories.Add(dwCatID, pCat))// Add the category to the server
            throw E_OUTOFMEMORY;
    }
    catch (HRESULT hresEx) {
        if (pCat)
            delete pCat;
        hres = hresEx;
    }
    m_csCatMap.Unlock();

    return hres;
}



//=========================================================================
// AddEventAttribute
// -----------------
//    Adds a vendor specific attribute to the specified category.
//    The ID of the event attribute must be unique.
//    The default attributes 'ACK COMMENT' and 'AREAS' are added and
//    handled internally. ATTRID_ACKCOMMENT and ATTRID_AREAS are reserved
//    attributes IDs.
//=========================================================================
HRESULT AeBaseServer::AddEventAttribute(DWORD dwCatID, DWORD dwAttrID, LPCWSTR szDescr, VARTYPE vt)
{
    _ASSERTE(szDescr != NULL);                 // Must not be NULL

    HRESULT           hres = S_OK;
    AeAttribute*  pAttr = NULL;

    m_csCatMap.Lock();
    try {
        if ((dwAttrID == ATTRID_ACKCOMMENT) || (dwAttrID == ATTRID_AREAS))
            throw E_INVALIDARG;                    // Internally used attribute ID specified

        AeCategory* pCat = NULL;
        // Check if the attribute ID is unique
        for (int i = 0; i < m_mapCategories.GetSize(); i++) {
            if (pCat = m_mapCategories.GetValueAt(i)) {
                if (pCat->ExistEventAttribute(dwAttrID)) {
                    throw E_OBJECT_IN_LIST;          // Attribute with the specified ID already exist
                }
            }
        }
        // Get the specified category
        pCat = m_mapCategories.Lookup(dwCatID);
        if (!pCat) throw E_INVALID_HANDLE;        // There is no category with the specified ID

        pAttr = new AeAttribute;              // Attribute ID is unique. Create a new instance.
        if (!pAttr) throw E_OUTOFMEMORY;
        // Initialize new instance
        hres = pAttr->Create(dwAttrID, szDescr, vt);
        if (FAILED(hres)) throw hres;

        hres = pCat->AddEventAttribute(pAttr);  // Add the attribute to the specified category
        if (FAILED(hres)) throw hres;
    }
    catch (HRESULT hresEx) {
        if (pAttr)
            delete pAttr;
        hres = hresEx;
    }
    m_csCatMap.Unlock();

    return hres;
}



//=========================================================================
// AddSingleStateConditionDef
// --------------------------
//    Adds a single state condition definition.
//    The ID of the condition definition must be unique. The parameters
//    specifies the default condition states.
//=========================================================================
HRESULT AeBaseServer::AddSingleStateConditionDef(
    DWORD dwCatID, DWORD dwCondDefID,
    LPCWSTR szName, LPCWSTR szDef, DWORD dwSeverity, LPCWSTR szDescr,
    BOOL fAckRequired)
{
    HRESULT              hres = S_OK;
    AeConditionDefiniton*  pCondDef = NULL;

    m_csCatMap.Lock();
    m_csCondDefMap.Lock();
    try {
        // ID must be unique
        if (m_mapConditionDefs.Lookup(dwCondDefID))
            throw E_OBJECT_IN_LIST;

        if (LookupConditionDef(szName))         // Name must be unique
            throw E_OBJECT_IN_LIST;
        // Get the specified category
        AeCategory* pCat = m_mapCategories.Lookup(dwCatID);
        if (!pCat) throw E_INVALID_HANDLE;        // There is no category with the specified ID

        if (pCat->EventType() != OPC_CONDITION_EVENT)
            throw E_FAIL;                          // Category must be specified for condition-related events

        pCondDef = new AeConditionDefiniton;        // Create new condition definition instance
        if (!pCondDef) throw E_OUTOFMEMORY;

        hres = pCondDef->Create(pCat, dwCondDefID, szName, szDef, dwSeverity, szDescr, fAckRequired);
        if (FAILED(hres)) throw hres;
        // Attach the condition definition to the specified category
        hres = pCat->AttachConditionDef(pCondDef);
        if (FAILED(hres)) throw hres;
        // Add the new condition definition to the server
        if (!m_mapConditionDefs.Add(dwCondDefID, pCondDef)) {
            pCat->DetachConditionDef(pCondDef);  // Detach the condition definition from the category
            throw E_OUTOFMEMORY;
        }
    }
    catch (HRESULT hresEx) {
        if (pCondDef)
            delete pCondDef;
        hres = hresEx;
    }
    m_csCatMap.Unlock();
    m_csCondDefMap.Unlock();

    return hres;
}



//=========================================================================
// AddMultiStateConditionDef
// -------------------------
//    Adds a multi state condition definition.
//    The ID of the condition definition must be unique. The parameters
//    specifies the default condition states.
//    At least one sub condition must be added with the function
//    AddSubConditionDef().
//=========================================================================
HRESULT AeBaseServer::AddMultiStateConditionDef(DWORD dwCatID, DWORD dwCondDefID, LPCWSTR szName)
{
    HRESULT              hres = S_OK;
    AeConditionDefiniton*  pCondDef = NULL;

    m_csCatMap.Lock();
    m_csCondDefMap.Lock();
    try {
        // ID must be unique
        if (m_mapConditionDefs.Lookup(dwCondDefID))
            throw E_OBJECT_IN_LIST;

        if (LookupConditionDef(szName))         // Name must be unique
            throw E_OBJECT_IN_LIST;
        // Get the specified category
        AeCategory* pCat = m_mapCategories.Lookup(dwCatID);
        if (!pCat) throw E_INVALID_HANDLE;        // There is no category with the specified ID

        if (pCat->EventType() != OPC_CONDITION_EVENT)
            throw E_FAIL;                          // Category must be specified for condition-related events

        pCondDef = new AeConditionDefiniton;        // Create new condition definition instance
        if (!pCondDef) throw E_OUTOFMEMORY;

        hres = pCondDef->Create(pCat, dwCondDefID, szName);
        if (FAILED(hres)) throw hres;
        // Attach the condition definition to the specified category
        hres = pCat->AttachConditionDef(pCondDef);
        if (FAILED(hres)) throw hres;
        // Add the new condition definition to the server
        if (!m_mapConditionDefs.Add(dwCondDefID, pCondDef)) {
            pCat->DetachConditionDef(pCondDef);  // Detach the condition definition from the category
            throw E_OUTOFMEMORY;
        }
    }
    catch (HRESULT hresEx) {
        if (pCondDef)
            delete pCondDef;
        hres = hresEx;
    }
    m_csCatMap.Unlock();
    m_csCondDefMap.Unlock();

    return hres;
}



//=========================================================================
// AddSubConditionDef
// ------------------
//    Adds a sub condition to a multi state condition definition.
//    The ID of the sun condition definition must be unique within the
//    multi state condition. The parameters specifies the default
//    sub condition states.
//    This function must be called at least twice for each multi
//    state condition definition added with AddMultiStateConditionDef().
//=========================================================================
HRESULT AeBaseServer::AddSubConditionDef(
    DWORD dwCondDefID, DWORD dwSubCondDefID,
    LPCWSTR szName, LPCWSTR szDef, DWORD dwSeverity, LPCWSTR szDescr,
    BOOL fAckRequired)
{
    HRESULT              hres = S_OK;
    AeConditionDefiniton*  pCondDef = NULL;

    m_csCondDefMap.Lock();
    try {
        if (dwSubCondDefID == 0)                  // 0 is used by single state conditions
            throw E_INVALIDARG;
        // Condition with specified ID must exist
        pCondDef = m_mapConditionDefs.Lookup(dwCondDefID);
        if (!pCondDef) throw E_INVALID_HANDLE;    // There is no condition definition with the specified ID

        hres = pCondDef->AddSubCondDef(dwSubCondDefID, szName, szDef, dwSeverity, szDescr, fAckRequired);

        if (FAILED(hres)) throw hres;
    }
    catch (HRESULT hresEx) {
        hres = hresEx;
    }
    m_csCondDefMap.Unlock();
    return hres;
}



//=========================================================================
// AddArea
// -------
//    Adds an area to the process area space.
//=========================================================================
HRESULT AeBaseServer::AddArea(DWORD dwParentAreaID, DWORD dwAreaID, LPCWSTR szName)
{
    _ASSERTE(szName != NULL);                  // Must not be NULL

    HRESULT     hres = S_OK;
    EventArea* pNewArea = NULL;

    try {
        if ((dwAreaID == AREAID_ROOT) || (dwAreaID == AREAID_UNSPECIFIED))
            throw E_INVALIDARG;                    // Internally used area ID specified

        EventArea* pParentArea = NULL;

        hres = m_RootArea.GetArea(dwParentAreaID, &pParentArea);
        if (FAILED(hres)) throw hres;           // The specified parent area doesn't exist

        pNewArea = new EventArea;                // Area ID is unique. Create a new instance.
        if (!pNewArea) throw E_OUTOFMEMORY;
        // Initialize new instance and add it to the parent area
        hres = pNewArea->Create(pParentArea, dwAreaID, szName);
        if (FAILED(hres)) throw hres;
    }
    catch (HRESULT hresEx) {
        if (pNewArea)
            delete pNewArea;
        hres = hresEx;
    }

    return hres;
}



//=========================================================================
// AddSource
// ---------
//    Adds an event source object to the process area space.
//
// Parameters:
//    dwAreaID                Identifier of the parent process area.
//                            Use AREAID_ROOT to add a top-level source
//                            object.
//    dwSrcID                 Unique source identifier
//    szName                  Name of the source. If fMultiSource is TRUE
//                            then this parameter specifies the fully
//                            qualified source name; otherwise only the
//                            partial source name.
//    fMultiSource            TRUE if this source is shared by multiple
//                            areas (see function AddExistingSource()).
//=========================================================================
HRESULT AeBaseServer::AddSource(DWORD dwAreaID, DWORD dwSrcID,
    LPCWSTR szName, BOOL fMultiSource /* = FALSE */)
{
    _ASSERTE(szName != NULL);                  // Must not be NULL

    HRESULT        hres = S_OK;
    AeSource*  pSrc = NULL;

    m_csSrcMap.Lock();
    try {
        if (m_mapSources.Lookup(dwSrcID))       // ID must be unique
            throw E_OBJECT_IN_LIST;

        EventArea* pArea = NULL;
        hres = m_RootArea.GetArea(dwAreaID, &pArea);
        if (FAILED(hres)) throw hres;           // There is no area with the specified ID

                                                  // Check if the name is unique
        if (fMultiSource) {
            if (LookupSource(szName)) {          // The fully qualified source name already specified
                throw E_OBJECT_IN_LIST;
            }
        }
        // Get the fully qualified source name
        LPWSTR pszQualifiedName = NULL;
        hres = pArea->GetQualifiedSourceName(szName, &pszQualifiedName);
        if (pszQualifiedName)
            pIMalloc->Free(pszQualifiedName);
        // Name already exist
        if (SUCCEEDED(hres)) throw E_OBJECT_IN_LIST;
        if (hres == E_INVALIDARG)                 // OK, the source name is unique
            hres = S_OK;
        if (FAILED(hres)) throw hres;           // Other error occured

        pSrc = new AeSource;                  // Create a new instance
        if (!pSrc) throw E_OUTOFMEMORY;
        // Initialize new instance
        hres = pSrc->Create(pArea, szName, fMultiSource);
        if (FAILED(hres)) throw hres;

        if (!m_mapSources.Add(dwSrcID, pSrc))   // Add the source to the server
            throw E_OUTOFMEMORY;
    }
    catch (HRESULT hresEx) {
        if (pSrc)
            delete pSrc;
        hres = hresEx;
    }
    m_csSrcMap.Unlock();

    return hres;
}



//=========================================================================
// AddExistingSource
// -----------------
//    Adds an existing event source object to an additional process area.
//
// Parameters:
//    dwAreaID                Identifier of the parent process area.
//                            Use AREAID_ROOT to add a top-level source
//                            object.
//    dwSrcID                 Unique source identifier. The source with
//                            this ID alreday must be created and added to
//                            another area. Use function AddSource() with
//                            flag 'fMultiSource' set.
//=========================================================================
HRESULT AeBaseServer::AddExistingSource(DWORD dwAreaID, DWORD dwSrcID)
{
    HRESULT  hres = S_OK;

    m_csSrcMap.Lock();
    try {
        EventArea*    pArea = NULL;
        AeSource*  pSrc = NULL;

        hres = m_RootArea.GetArea(dwAreaID, &pArea);
        if (FAILED(hres)) throw hres;           // There is no area with the specified ID

        pSrc = m_mapSources.Lookup(dwSrcID);
        if (!pSrc) throw E_INVALID_HANDLE;       // There is no source with the specified ID

        hres = pSrc->AddToAdditionalArea(pArea);
    }
    catch (HRESULT hresEx) {
        hres = hresEx;
    }
    m_csSrcMap.Unlock();

    return hres;
}



//=========================================================================
// AddCondition
// ------------
//    Adds a condition object to the process area space.
//
// Parameters:
//    dwAreaID                Identifier of the parent process area.
//                            Use AREAID_ROOT to add a top-level source
//                            object.
//    dwSrcID                 Unique source identifier. The source with
//                            this ID alreday must be created and added to
//                            another area. Use function AddSource() with
//                            flag 'fMultiSource' set.
//=========================================================================
HRESULT AeBaseServer::AddCondition(DWORD dwSrcID, DWORD dwCondDefID, DWORD dwCondID)
{
    HRESULT           hres = S_OK;
    AeCondition*  pCond = NULL;

    m_csSrcMap.Lock();
    m_csCondDefMap.Lock();
    m_csCondMap.Lock();
    try {

        AeSource* pSrc = m_mapSources.Lookup(dwSrcID);
        if (!pSrc) throw E_INVALID_HANDLE;       // There is no source with the specified ID

        AeConditionDefiniton* pCondDef = m_mapConditionDefs.Lookup(dwCondDefID);
        if (!pCondDef) throw E_INVALID_HANDLE;    // There is no condition definition with the specified ID

        if (m_mapConditions.Lookup(dwCondID))   // ID must be unique
            throw E_OBJECT_IN_LIST;

        pCond = new AeCondition;              // Create a new instance.
        if (!pCond) throw E_OUTOFMEMORY;
        // Initialize new instance
        hres = pCond->Create(dwCondID, pCondDef, pSrc);
        if (FAILED(hres)) throw hres;

        hres = pSrc->AttachCondition(pCond);    // Attach the condition the specified source
        if (FAILED(hres)) throw hres;
        // Add the condition to the server
        if (!m_mapConditions.Add(dwCondID, pCond)) {
            pSrc->DetachCondition(pCond);        // Detach the condition  from the source
            throw E_OUTOFMEMORY;
        }
    }
    catch (HRESULT hresEx) {
        if (pCond)
            delete pCond;
        hres = hresEx;
    }
    m_csCondMap.Unlock();
    m_csCondDefMap.Unlock();
    m_csSrcMap.Unlock();

    return hres;
}



//=========================================================================
// ProcessEvent
// ------------
//    Creates a new simple or tracking event and process it.
//
// Parameters:
//    dwCatID                 Category identifier. For simple events the
//                            event type of the category must be of
//                            type OPC_SIMPLE_EVENT and for tracking events
//                            of type OPC_TRACKING_EVENT.
//
//    dwSrcID                 Source identifier.
//
//    szMessage               Text which describes the event.
//
//    dwSeverity              The urgency of the event. Range is 1-1000.
//
//    szActorID               If NULL-Pointer, a simple event will be
//                            generated; otherwise a tracking event.
//
//    dwAttrCount             Number of attribute values specified in
//                            pvAttrValues. This number must be identical
//                            with the number of attributes added to
//                            the specified category. This parameter
//                            is only used for cross-check.
//
//    pvAttrValues            Array of attribute values. The order and
//                            the types must be identical with the
//                            attributes of the specified category.
//
//    pft                     Time that the event occured. If NULL-Pointer
//                            then the current time will be used.
//
//=========================================================================
HRESULT AeBaseServer::ProcessEvent(DWORD dwCatID, DWORD dwSrcID, LPCWSTR szMessage, DWORD dwSeverity, LPCWSTR szActorID,
    DWORD dwAttrCount, LPVARIANT pvAttrValues, LPFILETIME pft)
{
    _ASSERTE(szMessage != NULL);               // Must not be NULL

    if (dwSeverity < MIN_LOW_SEVERITY || dwSeverity > 1000)
        return E_INVALIDARG;

    HRESULT     hres = S_OK;
    AeEvent*   pEvent = NULL;
    int         i;

    m_csServers.Lock();
    i = m_arServers.GetSize();
    m_csServers.Unlock();
    if (!i) return S_OK;                         // There are no clients connected

    m_csSrcMap.Lock();
    m_csCatMap.Lock();
    try {
        // Check if there is a source with the specified ID
        AeSource* pSrc = m_mapSources.Lookup(dwSrcID);
        if (!pSrc) throw E_INVALID_HANDLE;        // There is no source with the specified ID

        AeCategory* pCat = m_mapCategories.Lookup(dwCatID);
        if (!pCat) throw E_INVALID_HANDLE;        // There is no category with the specified ID

                                                  // Invalid number of attributes specified
        _ASSERTE(dwAttrCount == pCat->NumOfAttrs() - pCat->NumOfInternalAttrs());
        if (dwAttrCount != (pCat->NumOfAttrs() - pCat->NumOfInternalAttrs()))
            throw E_INVALIDARG;

        pEvent = new AeEvent;                    // Create an event instance
        if (!pEvent) throw E_OUTOFMEMORY;
        // Initialize event instance as simple or tracking event
        hres = pEvent->Create(pCat, pSrc, szMessage, dwSeverity, szActorID, pvAttrValues, pft);
        if (FAILED(hres)) throw hres;
    }
    catch (HRESULT hresEx) {
        hres = hresEx;
    }
    m_csCatMap.Unlock();
    m_csSrcMap.Unlock();

    if (SUCCEEDED(hres)) {                     // Event successfully created
        hres = FireEvent(pEvent);               // Process it for all connected clients
    }

    if (pEvent) {                                // Now the event is processed for all connected
        pEvent->Release();                        // clients and no longer used by this part
    }
    return hres;
}



//=========================================================================
// ProcessConditionStateChanges
// ----------------------------
//    Changes the state of one or more conditions.
//
// Parameters:
//    IN
//       dwCount              The number of conditions to be changed.
//       pCondStateChanges    Array of class AeConditionChangeStates with the
//                            condition states.
//    OUT
//       errors              Array of HRESULTs indicating the success of
//                            the individial condition changes.
//                            This parameter is optional. 
//
// Return Code:
//    S_OK                    All succeeded
//    S_FALSE                 The operation succeeded but not for all
//                            specified conditions. Refer to individual
//                            error returns for more information.
//    E_XXX                   Error occured.
//=========================================================================
HRESULT AeBaseServer::ProcessConditionStateChanges(DWORD dwCount, AeConditionChangeStates* pCondStateChanges, BOOL useCurrentTime /*= TRUE */, HRESULT* errors /*= NULL */)
{
    DWORD                i;
    HRESULT              hres, hresRet = S_OK;
    AeEvent*            pEvent;
    AeEventArray        arEvents;
    AeCondition*     pCond;
    AeConditionChangeStates*   pCS;

    m_csCondMap.Lock();
    try {
        hres = arEvents.Create(dwCount);
        if (FAILED(hres)) throw hres;

        for (i = 0; i < dwCount; i++) {             // Handle all specified conditions
            pCS = &pCondStateChanges[i];
            // Search the condition
            pCond = m_mapConditions.Lookup(pCS->CondID());
            if (pCond) {                          // Change the states
                hres = pCond->ChangeState(*pCS, &pEvent, useCurrentTime);
                if (hres == S_OK) {              // If hres is S_FALSE then the nothing has changed
                    hres = arEvents.Add(pEvent);
                    if (FAILED(hres)) {
                        pEvent->Release();
                    }
                }
            }
            else {
                hres = E_INVALID_HANDLE;
            }
            if (errors) {
                errors[i] = hres;
            }
            if (FAILED(hres)) {
                hresRet = S_FALSE;
            }
        }

        hres = FireEvents(arEvents);            // Notify all subscribed clients
    }
    catch (HRESULT hresEx) {
        hresRet = hresEx;
    }
    m_csCondMap.Unlock();

    return hresRet;
}



//=========================================================================
// AckCondition
// ------------
//    Acknowledges a condition internally by the server.
//=========================================================================
HRESULT AeBaseServer::AckCondition(DWORD dwCondID, LPCWSTR szComment /*= NULL */)
{
    HRESULT hres = E_INVALID_HANDLE;             // Error code if there is no condition with the specified ID

    m_csCondMap.Lock();
    AeCondition* pCond = m_mapConditions.Lookup(dwCondID);
    if (pCond) {
        AeEvent *pEvent;
        hres = pCond->AcknowledgeByServer(szComment, &pEvent);
        if (hres == S_OK) {                       // Ack state has changed only if return code is S_OK.
            hres = FireEvent(pEvent);
            pEvent->Release();                     // Now the event is processed for all connected   
        }                                         // clients and no longer used by this part
    }
    m_csCondMap.Unlock();
    return hres;
}



//=========================================================================
// FireShutdownRequest
// -------------------
//    Fires a 'Shutdown Request' to all subscribed clients.
//
// Parameters:
//    szReason                A text string witch indicates the reason for
//                            the shutdown or a NULL pointer if there is 
//                            no reason provided.
//                            This parameter is optional. 
//=========================================================================
void AeBaseServer::FireShutdownRequest(LPCWSTR szReason /* = NULL */)
{
    m_csServers.Lock();
    for (int i = 0; i < m_arServers.GetSize(); i++) {
        m_arServers[i]->FireShutdownRequest(szReason);
    }
    m_csServers.Unlock();
}



//-------------------------------------------------------------------------
// IMPLEMENTATION
//-------------------------------------------------------------------------

//=========================================================================
// QueryEventCategories
// --------------------
//    Gets the Event Categories supported by the server.
//    Out parameters must be initialized and tested outside of this
//    function.
//=========================================================================
HRESULT AeBaseServer::QueryEventCategories(
    /* [in] */                    DWORD          dwEventType,
    /* [out] */                   DWORD       *  pdwCount,
    /* [size_is][size_is][out] */ DWORD       ** ppdwEventCategories,
    /* [size_is][size_is][out] */ LPWSTR      ** ppszEventCategoryDescs)
{
    CFixOutArray< DWORD >   aID;
    CFixOutArray< LPWSTR >  aDescr;

    HRESULT hres = S_OK;

    m_csCatMap.Lock();
    try {
        AeCategory* pCat = NULL;
        // Calculate the size of the result arrays
        if (dwEventType == OPC_ALL_EVENTS) {
            *pdwCount = m_mapCategories.GetSize();
        }
        else {
            for (int i = 0; i < m_mapCategories.GetSize(); i++) {
                pCat = m_mapCategories.GetValueAt(i);
                if (!pCat) throw E_FAIL;

                if (pCat->EventType() & dwEventType) {
                    (*pdwCount)++;
                }
            }
        }
        // Allocate and initialize the result arrays
        aID.Init(*pdwCount, ppdwEventCategories);
        aDescr.Init(*pdwCount, ppszEventCategoryDescs);

        DWORD dwFound = 0;                        // Fill the result arrays
        for (int i = 0; i < m_mapCategories.GetSize(); i++) {

            pCat = m_mapCategories.GetValueAt(i);
            if (!pCat) throw E_FAIL;

            if (pCat->EventType() & dwEventType) {
                _ASSERTE(dwFound < *pdwCount);
                aID[dwFound] = pCat->CatID();
                aDescr[dwFound] = pCat->Descr().CopyCOM();
                if (aDescr[dwFound] == NULL) throw E_OUTOFMEMORY;
                dwFound++;
            }
        }
    }
    catch (HRESULT hresEx) {
        aID.Cleanup();
        aDescr.Cleanup();
        hres = hresEx;
    }
    m_csCatMap.Unlock();

    return hres;
}



//=========================================================================
// QueryConditionNames
// -------------------
//    Gets the name of the conditions of the specified Category.
//    Out parameters must be initialized and tested outside of this
//    function.
//=========================================================================
HRESULT AeBaseServer::QueryConditionNames(
    /* [in] */                    DWORD          dwEventCategory,
    /* [out] */                   DWORD       *  pdwCount,
    /* [size_is][size_is][out] */ LPWSTR      ** ppszConditionNames)
{
    HRESULT hres = S_OK;

    m_csCatMap.Lock();
    try {
        AeCategory* pCat = m_mapCategories.Lookup(dwEventCategory);
        if (!pCat) throw E_INVALIDARG;            // There is no category with the specified ID

                                                  // Only Categories related to type 'Condition Event' can have Conditions
        if (pCat->EventType() != OPC_CONDITION_EVENT)
            throw E_INVALIDARG;

        hres = pCat->QueryConditionNames(
            pdwCount,
            ppszConditionNames);

        if (FAILED(hres)) throw hres;
    }
    catch (HRESULT hresEx) {
        hres = hresEx;
    }
    m_csCatMap.Unlock();

    return hres;
}



//=========================================================================
// QuerySubConditionNames
// -----------------------
//    Gets the name of all sub conditions of the specified condition.
//    Out parameters must be initialized and tested outside of this
//    function.
//=========================================================================
HRESULT AeBaseServer::QuerySubConditionNames(
    /* [in] */                    LPWSTR         szConditionName,
    /* [out] */                   DWORD       *  pdwCount,
    /* [size_is][size_is][out] */ LPWSTR      ** ppszSubConditionNames)
{
    HRESULT hres = S_OK;

    m_csCondDefMap.Lock();
    try {
        AeConditionDefiniton* pCondDef = LookupConditionDef(szConditionName);
        if (!pCondDef) throw E_INVALIDARG;

        hres = pCondDef->QuerySubConditionNames(
            pdwCount,
            ppszSubConditionNames);
    }
    catch (HRESULT hresEx) {
        hres = hresEx;
    }
    m_csCondDefMap.Unlock();

    return hres;
}



//=========================================================================
// QueryEventAttributes
// --------------------
//    Calls the QueryEventAttributes() function of the specified category.
//    Out parameters must be initialized and tested outside of this
//    function.
//=========================================================================
HRESULT AeBaseServer::QueryEventAttributes(
    /* [in] */                    DWORD          dwEventCategory,
    /* [out] */                   DWORD       *  pdwCount,
    /* [size_is][size_is][out] */ DWORD       ** ppdwAttrIDs,
    /* [size_is][size_is][out] */ LPWSTR      ** ppszAttrDescs,
    /* [size_is][size_is][out] */ VARTYPE     ** ppvtAttrTypes)
{
    HRESULT hres = S_OK;

    m_csCatMap.Lock();
    try {
        AeCategory* pCat = m_mapCategories.Lookup(dwEventCategory);
        if (!pCat) throw E_INVALIDARG;            // There is no category with the specified ID

        hres = pCat->QueryEventAttributes(
            pdwCount,
            ppdwAttrIDs,
            ppszAttrDescs,
            ppvtAttrTypes);

        if (FAILED(hres)) throw hres;
    }
    catch (HRESULT hresEx) {
        hres = hresEx;
    }
    m_csCatMap.Unlock();

    return hres;
}



//=========================================================================
// QuerySourceConditions
// ---------------------
//    Gets the name of the conditions of the specified source.
//    Out parameters must be initialized and tested outside of this
//    function.
//=========================================================================
HRESULT AeBaseServer::QuerySourceConditions(
    /* [in] */                    LPWSTR         szSource,
    /* [out] */                   DWORD       *  pdwCount,
    /* [size_is][size_is][out] */ LPWSTR      ** ppszConditionNames)
{
    HRESULT  hres;

    m_csSrcMap.Lock();
    try {
        AeSource* pSrc = LookupSource(szSource);
        if (!pSrc) throw E_INVALIDARG;

        hres = pSrc->QueryConditionNames(
            pdwCount,
            ppszConditionNames);

        if (FAILED(hres)) throw hres;
    }
    catch (HRESULT hresEx) {
        hres = hresEx;
    }
    m_csSrcMap.Unlock();

    return hres;
}



//=========================================================================
// GetConditionState
// -----------------
//    Gets current state information for the condition instance
//    corresponding to the source and condition name.
//    Out parameters must be initialized and tested outside of this
//    function.
//=========================================================================
HRESULT AeBaseServer::GetConditionState(
    /* [in] */                    LPWSTR         szSource,
    /* [in] */                    LPWSTR         szConditionName,
    /* [in] */                    DWORD          dwNumEventAttrs,
    /* [size_is][in] */           DWORD       *  pdwAttributeIDs,
    /* [out] */                   OPCCONDITIONSTATE
    ** ppConditionState)
{
    HRESULT  hres;

    m_csSrcMap.Lock();
    try {
        AeSource* pSrc = LookupSource(szSource);
        if (!pSrc) throw E_INVALIDARG;

        hres = pSrc->GetConditionState(
            szConditionName,
            dwNumEventAttrs,
            pdwAttributeIDs,
            ppConditionState);

        if (FAILED(hres)) throw hres;
    }
    catch (HRESULT hresEx) {
        hres = hresEx;
    }
    m_csSrcMap.Unlock();

    return hres;
}



//=========================================================================
// AckCondition
// ------------
//    Acknowledges one or more conidtions in the event sever.
//    Out parameters must be initialized and tested outside of this
//    function.
//=========================================================================
HRESULT AeBaseServer::AckCondition(
    /* [in] */                    DWORD          dwCount,
    /* [string][in] */            LPWSTR         szAcknowledgerID,
    /* [string][in] */            LPWSTR         szComment,
    /* [size_is][in] */           LPWSTR      *  pszSource,
    /* [size_is][in] */           LPWSTR      *  pszConditionName,
    /* [size_is][in] */           FILETIME    *  pftActiveTime,
    /* [size_is][in] */           DWORD       *  pdwCookie,
    /* [size_is][size_is][out] */ HRESULT     ** ppErrors)
{
    HRESULT                 hres;
    AeSource*           pSrc;
    CFixOutArray< HRESULT > aErr;
    AeEventArray           arEvents;
    AeEvent*               pEvent;

    m_csSrcMap.Lock();
    try {
        hres = arEvents.Create(dwCount);
        if (FAILED(hres)) throw hres;

        aErr.Init(dwCount, ppErrors, S_OK);     // Initialize with default value

        for (DWORD i = 0; i < dwCount; i++) {       // Acknowledge all specified conditions

            if (pdwCookie[i]) {                    // There is cookie specified. The cookie specifies the condition
                AeCondition* pCond = (AeCondition*)(pdwCookie[i]);
                m_csCondMap.Lock();
                // Test if the specific condition is still available
                if (m_mapConditions.FindVal(pCond) != -1) {
                    if (!pCond->IsAttachedTo(pszSource[i], pszConditionName[i])) {
                        aErr[i] = E_INVALIDARG;       // Invalid Condition or Source Name
                    }
                    else {
                        aErr[i] = pCond->Acknowledge(szAcknowledgerID, szComment, pftActiveTime[i], &pEvent);
                    }
                }
                else {
                    aErr[i] = E_INVALIDARG;          // Invalid Cookie
                }
                m_csCondMap.Unlock();
            }
            else {                                 // There is no cookie specified
                pSrc = LookupSource(pszSource[i]);
                if (pSrc) {
                    aErr[i] = pSrc->AckCondition(
                        szAcknowledgerID,
                        szComment,
                        pszConditionName[i],
                        pftActiveTime[i],
                        &pEvent);
                }
                else {
                    aErr[i] = E_INVALIDARG;          // Source name doesn't exist
                }
            }
            if (aErr[i] == S_OK) {                 // Ack state has changed only if return code is S_OK.
                aErr[i] = arEvents.Add(pEvent);
                if (FAILED(aErr[i])) {
                    pEvent->Release();
                    hres = S_FALSE;                  // Not all entries in ppErrors are S_OK
                }
            }
            else {
                hres = S_FALSE;                     // Not all entries in ppErrors are S_OK
            }
        }
        FireEvents(arEvents);                   // Notify all subscribed clients
    }
    catch (HRESULT hresEx) {
        aErr.Cleanup();
        hres = hresEx;
    }
    m_csSrcMap.Unlock();

    return hres;
}



//=========================================================================
// EnableConditions
// ----------------
//    Enables or disables all conditions which belongs to the specified
//    sources or area names.
//
// Parameters:
//    fEnable                 If TRUE the conditions are enabled;
//                            otherwise disabled.
//    fBySource               If TRUE pszNames specifies source names;
//                            otherwise area names.
//    dwCount                 Number of source/area names in pszNames.
//    pszNames                Array of source respectively area names.
//=========================================================================
HRESULT AeBaseServer::EnableConditions(
    /* [in] */                    BOOL           fEnable,
    /* [in] */                    BOOL           fBySource,
    /* [in] */                    DWORD          dwCount,
    /* [size_is][in] */           LPWSTR      *  pszNames)
{
    HRESULT        hresRet = S_OK;
    HRESULT        hres = S_OK;
    DWORD          i;

    if (dwCount == 0) {
        return E_INVALIDARG;
    }

    m_csCondMap.Lock();                          // The source/area functions uses the condition array
    if (fBySource) {                             // Change Condition State by Source
        AeSource*  pSrc;

        m_csSrcMap.Lock();

        for (i = 0; i < dwCount; i++) {
            pSrc = LookupSource(pszNames[i]);
            if (!pSrc) {
                hresRet = E_INVALIDARG;
            }
            else {
                AeEventArray  arEvents;

                hres = pSrc->EnableConditions(fEnable, &arEvents);
                if (FAILED(hres)) {
                    hresRet = hres;
                }
                FireEvents(arEvents);             // Notify all subscribed clients
            }
        }

        m_csSrcMap.Unlock();
    }
    else {                                       // Change Condition State by Area
        EventArea* pArea;

        for (i = 0; i < dwCount; i++) {

            if (pszNames[i] == NULL || *pszNames[i] == NULL) {
                hresRet = E_INVALIDARG;             // E_INVALIDARG if invalid Area is specified
                continue;
            }

            hres = m_RootArea.PositionTo(pszNames[i], &pArea);
            if (SUCCEEDED(hres)) {
                CSimpleValArray<AeEventArray*> arparEvents;
                hres = pArea->EnableConditions(fEnable, arparEvents);
                if (FAILED(hres)) {
                    hresRet = hres;
                }
                for (int z = 0; z < arparEvents.GetSize(); z++) {
                    FireEvents(*(arparEvents[z])); // Notify all subscribed clients
                    delete arparEvents[z];
                }
                arparEvents.RemoveAll();
            }
            else if (hres == OPC_E_INVALIDBRANCHNAME) {
                hresRet = E_INVALIDARG;             // E_INVALIDARG if invalid Area is specified
            }
            else {
                hresRet = hres;                     // other error
            }
        }
    }
    m_csCondMap.Unlock();                        // The source/area functions uses the condition array

    return hresRet;
}



//=========================================================================
// TranslateToItemIDs
// ------------------
//    Translates the attributes specified by the parameters to
//    OPC Data Access ItemIDs.
//    Out parameters must be initialized and tested outside of this
//    function.
//=========================================================================
HRESULT AeBaseServer::TranslateToItemIDs(
    /* [in] */                    LPWSTR         szSource,
    /* [in] */                    DWORD          dwEventCategory,
    /* [in] */                    LPWSTR         szConditionName,
    /* [in] */                    LPWSTR         szSubconditionName,
    /* [in] */                    DWORD          dwCount,
    /* [size_is][in] */           DWORD       *  pdwAssocAttrIDs,
    /* [size_is][size_is][out] */ LPWSTR      ** ppszAttrItemIDs,
    /* [size_is][size_is][out] */ LPWSTR      ** ppszNodeNames,
    /* [size_is][size_is][out] */ CLSID       ** ppCLSIDs)
{
    CFixOutArray< LPWSTR >  aItemID;
    CFixOutArray< LPWSTR >  aNode;
    CFixOutArray< CLSID >   aCLSID;

    HRESULT hres = S_OK;

    m_csCondMap.Lock();

    try {
        aItemID.Init(dwCount, ppszAttrItemIDs); // Initialize the result parameters
        aNode.Init(dwCount, ppszNodeNames);
        aCLSID.Init(dwCount, ppCLSIDs);

        AeCondition*  pCond;
        AeCategory*   pCat;
        DWORD             dwSubCondDefID;
        int               condcnt;                // Condition index counter
        DWORD             attrcnt;                // Attribute index counter

        hres = E_INVALIDARG;

        // Search the specified condition / category

        for (condcnt = 0; condcnt < m_mapConditions.GetSize(); condcnt++) {
            pCond = m_mapConditions.m_aVal[condcnt];
            if (pCond->IsAttachedTo(szSource, dwEventCategory, szConditionName, szSubconditionName, &dwSubCondDefID, &pCat)) {
                hres = S_OK;
                break;                              // Found
            }
        }

        if (FAILED(hres)) throw hres;           // Not found

        // Translate all specified attributes to ItemIDs

        for (attrcnt = 0; attrcnt < dwCount; attrcnt++) {

            if (pCat->ExistEventAttribute(pdwAssocAttrIDs[attrcnt])) {
                // Found
                hres = OnTranslateToItemIdentifier(
                    m_mapConditions.m_aKey[condcnt], dwSubCondDefID,
                    pdwAssocAttrIDs[attrcnt],
                    &aItemID[attrcnt], &aNode[attrcnt], &aCLSID[attrcnt]);
            }
            else {
                hres = E_INVALIDARG;                // Not found. Invalid attribute.
            }
            if (FAILED(hres)) throw hres;
        }
    }
    catch (HRESULT hresEx) {
        aItemID.Cleanup();
        aNode.Cleanup();
        aCLSID.Cleanup();
        hres = hresEx;
    }

    m_csCondMap.Unlock();

    return hres;
}



//=========================================================================
// GetEventsForRefresh
// -------------------
//    Creates and returns the events appropriate to the active and
//    inactive, unacknowledged conidtions in the event sever.
//    Checks the 'cancel refresh' flag of the event subscription 
//    if in progress.
//    This function is called from the EventRefreshThread.
//
// Parameters:
//    pSubscr                 Used to check the 'cancel refresh' flag
//                            of the event subscription.
//    arEventPtrs             Whrere to return the resulting events.
//=========================================================================
HRESULT AeBaseServer::GetEventsForRefresh(
    /* [in] */                    AeComSubscriptionManager* pSubscr,
    /* [out] */                   CSimpleArray<AeEvent*>& arEventPtrs)
{
    HRESULT           hres = S_OK;
    AeCondition*  pCond;
    AeEvent*         pEvent;
    int               i;

    m_csCondMap.Lock();
    try {
        for (i = 0; i < m_mapConditions.GetSize(); i++) {
            // Terminate function if refresh canceled
            if (pSubscr->IsCancelRefreshActive()) throw S_OK;

            pCond = m_mapConditions.GetValueAt(i);

            if (pCond->IsEnabled() &&
                (pCond->IsActive() || !pCond->IsAcked())) {
                // Enabled condition which is active
                // or inactive, unacknowledged
                // Create an event instance
                hres = pCond->CreateEventInstance(&pEvent);
                if (FAILED(hres)) throw hres;

                if (!arEventPtrs.Add(pEvent)) throw E_OUTOFMEMORY;
            }
        }
    }
    catch (HRESULT hresEx) {
        for (i = 0; i < arEventPtrs.GetSize(); i++) {
            arEventPtrs[i]->Release();
        }
        arEventPtrs.RemoveAll();
        hres = hresEx;
    }
    m_csCondMap.Unlock();

    return hres;
}



//=========================================================================
// ExistEventAttribute
// -------------------
//    Returns TRUE if the specified Event Attribute exist as member of
//    the specified Event Category; otherwise FALSE.
//=========================================================================
BOOL AeBaseServer::ExistEventAttribute(DWORD dwEventCategory, DWORD dwAttrID)
{
    BOOL fExist = FALSE;

    m_csCatMap.Lock();

    AeCategory* pCat = m_mapCategories.Lookup(dwEventCategory);
    if (pCat) {
        fExist = pCat->ExistEventAttribute(dwAttrID);
    }

    m_csCatMap.Unlock();
    return fExist;
}



//=========================================================================
// ExistEventCategory
// ------------------
//    Returns TRUE if the specified Event Category exist; otherwise FALSE.
//=========================================================================
BOOL AeBaseServer::ExistEventCategory(DWORD dwEventCategory)
{
    m_csCatMap.Lock();
    BOOL fExist = (m_mapCategories.FindKey(dwEventCategory) == -1) ? FALSE : TRUE;
    m_csCatMap.Unlock();
    return fExist;
}



//=========================================================================
// ExistSource
// -----------
//    Returns TRUE if the Source with the specified name exist in the
//    process area space; otherwise FALSE.
//
// Parameters:
//    szName                  Name of the source.
//                            It is possible to specify a source name using
//                            the wildcard sytax described in Appendix A
//                            of the AE Specification.
//=========================================================================
BOOL AeBaseServer::ExistSource(LPCWSTR szName)
{
    DWORD dwSize;
    BOOL  fSourceExist = FALSE;

    m_csSrcMap.Lock();
    dwSize = m_mapSources.GetSize();
    while (dwSize--) {
        if (MatchPattern(m_mapSources.m_aVal[dwSize]->Name(), szName)) {
            fSourceExist = TRUE;
            break;
        }
    }
    m_csSrcMap.Unlock();
    return fSourceExist;
}



//=========================================================================
// ExistArea
// ---------
//    Returns TRUE if the Area with the specified name exist in the
//    process area space; otherwise FALSE.
//
// Parameters:
//    szName                  Name of the area.
//                            It is possible to specify an area name using
//                            the wildcard sytax described in Appendix A
//                            of the AE Specification.
//=========================================================================
BOOL AeBaseServer::ExistArea(LPCWSTR szName)
{
    return m_RootArea.ExistArea(szName);
}



//=========================================================================
// FireEvent
// ---------
//    Fires one event to all subscribed clients.
//=========================================================================
HRESULT AeBaseServer::FireEvent(AeEvent* pEvent)
{
    m_csServers.Lock();
    for (long i = 0; i < m_arServers.GetSize(); i++) {
        m_arServers[i]->ProcessEvents(1, &pEvent);// Process the event for this client
    }
    m_csServers.Unlock();
    return S_OK;
}



//=========================================================================
// FireEvent
// ---------
//    Fires one ore more events to all subscribed clients.
//=========================================================================
HRESULT AeBaseServer::FireEvents(AeEventArray& Events)
{
    m_csServers.Lock();
    for (long i = 0; i < m_arServers.GetSize(); i++) {
        m_arServers[i]->ProcessEvents(
            Events.NumOfEvents(), Events.EventPtrArray()
        ); // Process the event for this client
    }
    m_csServers.Unlock();
    return S_OK;
}



//=========================================================================
// LookupSource
// ------------
//    Gets the source with the specified source name.
//    This function assumes that m_mapSources is already locked by
//    the caller.
//=========================================================================
AeSource* AeBaseServer::LookupSource(LPCWSTR szName)
{
    AeSource* pSrc = NULL;

    if (szName) {
        for (int i = 0; i < m_mapSources.GetSize(); i++) {
            if (m_mapSources.m_aVal[i]->Name() == szName) {
                pSrc = m_mapSources.m_aVal[i];
                break;
            }
        }
    }
    return pSrc;
}



//=========================================================================
// LookupConditionDef
// ------------------
//    Gets the condition definition the specified condition definition
//    name. This function assumes that m_mapConditionDefs is already
//    locked by the caller.
//=========================================================================
AeConditionDefiniton* AeBaseServer::LookupConditionDef(LPCWSTR szName)
{
    AeConditionDefiniton* pCondDef = NULL;

    for (int i = 0; i < m_mapConditionDefs.GetSize(); i++) {
        if (m_mapConditionDefs.m_aVal[i]->Name() == szName) {
            pCondDef = m_mapConditionDefs.m_aVal[i];
            break;
        }
    }
    return pCondDef;
}



//=========================================================================
// AddServerToList
// ---------------
//    Adds an event server instance to the event server list.
//    This method is called when a client connects.
//=========================================================================
HRESULT AeBaseServer::AddServerToList(AeComBaseServer* pServer)
{
    _ASSERTE(pServer != NULL);

    HRESULT hres;
    m_csServers.Lock();
    hres = m_arServers.Add(pServer) ? S_OK : E_OUTOFMEMORY;
    m_csServers.Unlock();
    return hres;
}



//=========================================================================
// RemoveServerFromList
// --------------------
//    Removes a server instance from the event server list.
//    This method is called when a client has properly disconnected,
//    that is freed all the COM interfaces it owns.
//=========================================================================
HRESULT AeBaseServer::RemoveServerFromList(AeComBaseServer* pServer)
{
    _ASSERTE(pServer != NULL);

    HRESULT hres;
    m_csServers.Lock();
    hres = m_arServers.Remove(pServer) ? S_OK : E_FAIL;
    m_csServers.Unlock();
    return hres;
}
//DOM-IGNORE-END


#endif