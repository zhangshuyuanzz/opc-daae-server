/*
 * Copyright (c) 2020 Technosoftware GmbH. All rights reserved
 * Web: https://technosoftware.com 
 * 
 * The source code in this file is covered under a dual-license scenario:
 *   - Owner of a purchased license: RPL 1.5
 *   - GPL V3: everybody else
 *
 * RPL license terms accompanied with this source code.
 * See https://technosoftware.com/license/RPLv15License.txt
 *
 * GNU General Public License as published by the Free Software Foundation;
 * version 3 of the License are accompanied with this source code.
 * See https://technosoftware.com/license/GPLv3License.txt
 *
 * This source code is distributed in the hope that it will be useful, but
 * WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY
 * or FITNESS FOR A PARTICULAR PURPOSE.
 */

/*
 * OPC Server SDK C++ generic server interface.
 *
 * IMPORTANT: DON'T CHANGE THIS FILE.
 */

//-----------------------------------------------------------------------------
// INCLUDES
//-----------------------------------------------------------------------------
#include "stdafx.h"
#include "DaServer.h"
#ifdef _OPC_SRV_AE
#include "AeServer.h"
#endif
#ifdef _OPC_EVALUATION_VERSION
#include "LicenseHandler.h"
#endif
#include "IClassicBaseNodeManager.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace IClassicBaseNodeManager
{

static CSimplePtrArray<WideString*> m_apPropDescr;

//----------------------------------------------------------------------------
// These structures, enumerations and classes match the OPC specifications
// or the server EXE file.
//----------------------------------------------------------------------------

//----------------------------------------------------------------------------
// OPC DA/AE Server SDK C++ API Callback Methods
// (Called by the customization assembly)
//----------------------------------------------------------------------------

/**
 * @fn  HRESULT CreateOneItem( LPWSTR itemId, DWORD accessRights, LPVARIANT value, BOOL active, OPCEUTYPE euType, LPVARIANT euInformation, DeviceItem** deviceItem )
 *
 * @brief   Creates one item in the server data structure. This function is called from the user
 *          DLL for each item created.
 *
 * @param   itemId              Identifier for the item.
 * @param   accessRights        The access rights.
 * @param   value               The value.
 * @param   active              true to active.
 * @param   euType              Type of the Engineering Unit.
 * @param   euInformation       Information describing the pv eu.
 * @param [in,out]  deviceItem  If non-null, the d items.
 *
 * @return  The new one item.
 */

HRESULT CreateOneItem(
    LPWSTR          itemId,
    DWORD           accessRights,
    LPVARIANT       value,
    BOOL            active,
    OPCEUTYPE       euType,
    LPVARIANT       euInformation,
    DeviceItem**    deviceItem
    )

{
    HRESULT hres;

    *deviceItem = new DeviceItem; // Allocate space for new server item.
    if (*deviceItem == nullptr) {
        hres = E_OUTOFMEMORY;
    }
    else { // Initialize new Item instance
        hres = (*deviceItem)->Create(itemId, accessRights, value, active, 0, nullptr, nullptr, euType, euInformation);
        if (SUCCEEDED(hres)) { // Add new created item to the item list
            hres = gpDataServer->AddDeviceItem(*deviceItem);
        }
    }
    VariantClear(value);
    return hres;
}

HRESULT DLLCALL AddItem(LPWSTR itemId, DaAccessRights accessRights, LPVARIANT initValue, bool active, DaEuType euType, double minValue, double maxValue, void** deviceItem)
{
    HRESULT hres = S_OK;

    DeviceItem* pDeviceItem = nullptr;

    // Initialize new Item instance
    if (static_cast<OPCEUTYPE>(euType) == OPC_NOENUM)
    {
        hres = CreateOneItem(
            itemId, // ItemId
            accessRights, // DaAccessRights
            initValue, // Initial Value
            active, // Active State
            OPC_NOENUM, // EU Type
            nullptr, // EU Info
            &pDeviceItem // The created classic device item
            );
    }
    else
    {
        if (static_cast<OPCEUTYPE>(euType) == OPC_ANALOG)
        {
            // EU Info must contain exaclty two doubles
            // corresponding to the LOW and HI EU range.
            VARIANT varEUInfo;
            VariantInit(&varEUInfo);
            V_VT(&varEUInfo) = VT_ARRAY | VT_R8;

            SAFEARRAYBOUND rgs;
            rgs.cElements = 2;
            rgs.lLbound = 0;

            try {
                CHECK_PTR(V_ARRAY(&varEUInfo) = SafeArrayCreate(VT_R8, 1, &rgs))
                    // put initialized variant into array
                    double dLow = minValue;
                double dHi = maxValue;
                long lElIndex;

                lElIndex = 0;                       // LOW EU range
                CHECK_RESULT(SafeArrayPutElement(V_ARRAY(&varEUInfo), &lElIndex, &dLow))
                    lElIndex++;                     // HI EU range
                CHECK_RESULT(SafeArrayPutElement(V_ARRAY(&varEUInfo), &lElIndex, &dHi))
            }
            catch (HRESULT) {                       // cleanup if failure
                VariantClear(&varEUInfo);
                throw;
            }
            hres = CreateOneItem(
                itemId,                             // ItemId
                static_cast<DWORD>(accessRights),   // DaAccessRights
                initValue,                          // Initial Value
                active,                             // Active State
                OPC_ANALOG,                         // EU Type
                &varEUInfo,                         // EU Info
                &pDeviceItem                        // The created classic device item
                );
            VariantClear(&varEUInfo);

        }
    }
    VariantClear(initValue);
    if (deviceItem != nullptr)
    {
        *deviceItem = pDeviceItem;
    }
    return hres;
}

HRESULT DLLCALL AddItem(LPWSTR itemId, DaAccessRights accessRights, LPVARIANT initValue, void** deviceItem)
{
    return AddItem(itemId, accessRights, initValue, true, NoEnum, 0.0, 0.0, deviceItem);
}

HRESULT DLLCALL AddAnalogItem(LPWSTR itemId, DaAccessRights accessRights, LPVARIANT initValue, double minValue, double maxValue, void** deviceItem)
{
    return AddItem(itemId, accessRights, initValue, true, Analog, minValue, maxValue, deviceItem);
}

HRESULT DLLCALL RemoveItem(void* deviceItem)
{
    HRESULT hres;

    if (deviceItem == nullptr) {
        return(S_FALSE);
    }
    hres = gpDataServer->RemoveDeviceItem(static_cast<DeviceItem*>(deviceItem));
    return hres;
};

HRESULT DLLCALL DeleteItem(void* deviceItem)
{
    if (deviceItem == nullptr) {
        return(S_FALSE);
    }
    gpDataServer->DeleteDeviceItem(static_cast<DeviceItem*>(deviceItem));
    return S_OK;
};

HRESULT DLLCALL AddProperty(int propertyId, LPWSTR description, LPVARIANT valueType)
{
    HRESULT hres;
    DaItemProperty* pProp;

    hres = gpDataServer->GetItemProperty(propertyId, &pProp);
    if (FAILED(hres)) {

        WideString* p1 = new WideString;
        hres = p1->SetString(description);
        if (FAILED(hres)) {
            return hres;
        }

        if (m_apPropDescr.Add(p1))
        {
            pProp = new DaItemProperty(propertyId, V_VT(valueType), static_cast<LPWSTR>(*p1));
            hres = gpDataServer->AddItemProperty(pProp);
        }
        else
        {
            return S_FALSE;
        }
    }

    return hres;
};


HRESULT DLLCALL SetItemValue(
    void* deviceItemHandle,
    LPVARIANT newValue,
    short quality,
    FILETIME timestamp
    )
{
    HRESULT hres;
    DeviceItem* pItem = static_cast<DeviceItem*>(deviceItemHandle);

    gpDataServer->readWriteLock_.BeginWriting(); // We are manipulating the DeviceItem make
    // sure no one else gets access during update.
    if (pItem == nullptr) {
        gpDataServer->readWriteLock_.EndWriting(); // UnLock DeviceItem
        hres = S_FALSE;
        return(hres);
    }

    if (newValue == nullptr || V_VT(newValue) == VT_NULL) {
        // Update only quality and time stamp
        hres = pItem->set_ItemQuality(quality, &timestamp);
    }
    else {
        hres = pItem->set_ItemValue(newValue, quality, &timestamp);
    }
    gpDataServer->readWriteLock_.EndWriting(); // UnLock DeviceItem

    return hres;
}

void DLLCALL SetServerState(ServerState serverState)
{
    gpDataServer->SetServerState(static_cast<OPCSERVERSTATE>(serverState));
};

void DLLCALL GetActiveItems(int * numItemHandles, void* ** deviceItemHandles)
{
    DWORD dwSize;
    DeviceItem* pItem;
    int numHandles = 0;

    gpDataServer->m_ItemListLock.BeginReading(); // Protect item list access
    // Handle all items
    dwSize = static_cast<DWORD>(gpDataServer->m_arServerItems.GetSize());
    void** activeItemHandles;

    activeItemHandles = new void*[dwSize];
    for (DWORD i = 0; i < dwSize; i++) {
        pItem = gpDataServer->m_arServerItems[i];
        if (pItem == nullptr) {
            continue; // Handle next item
        }

        if (pItem->get_ActiveCount() > 0) {
            activeItemHandles[numHandles] = pItem;
            numHandles++;
        }
    } // handle all items
    gpDataServer->m_ItemListLock.EndReading(); // Release item list protection

    *numItemHandles = numHandles;
    *deviceItemHandles = activeItemHandles;
};

void DLLCALL GetClients(int * numClientHandles, void* ** clientHandles, LPWSTR ** clientNames)
{
    gpDataServer->GetClients(numClientHandles, clientHandles, clientNames);
}

void DLLCALL GetGroups(void * clientHandle, int * numGroupHandles, void* ** groupHandles, LPWSTR ** groupNames)
{
    gpDataServer->GetGroups(clientHandle, numGroupHandles, groupHandles, groupNames);
}

void GetGroupState(void * groupHandle, DaGroupState * groupState)
{
    gpDataServer->GetGroupState(groupHandle, groupState);
}

void DLLCALL FireShutdownRequest(LPCWSTR reason)
{
    gpDataServer->FireShutdownRequest(reason);
#ifdef _OPC_SRV_AE
    gpEventServer->FireShutdownRequest(reason);
#endif
};

HRESULT DLLCALL AddSimpleEventCategory(int categoryId, LPWSTR categoryDescription)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->AddSimpleEventCategory(categoryId, categoryDescription);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL AddTrackingEventCategory(int categoryId, LPWSTR categoryDescription)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->AddTrackingEventCategory(categoryId, categoryDescription);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL AddConditionEventCategory(int categoryId, LPWSTR categoryDescription)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->AddConditionEventCategory(categoryId, categoryDescription);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL AddEventAttribute(int categoryId, int eventAttribute, LPWSTR attributeDescription, VARTYPE dataType)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->AddEventAttribute(categoryId, eventAttribute, attributeDescription, dataType);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL AddSingleStateConditionDefinition(int categoryId, int conditionId, LPWSTR name, LPWSTR condition, int severity, LPWSTR description, bool ackRequired)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->AddSingleStateConditionDef(categoryId, conditionId, name, condition, severity, description, ackRequired);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL AddMultiStateConditionDefinition(int categoryId, int conditionId, LPWSTR name)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->AddMultiStateConditionDef(categoryId, conditionId, name);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL AddSubConditionDefinition(int conditionId, int subConditionId, LPWSTR name, LPWSTR condition, int severity, LPWSTR description, bool ackRequired)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->AddSubConditionDef(conditionId, subConditionId, name, condition, severity, description, ackRequired);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL AddArea(int parentAreaId, int areaId, LPWSTR name)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->AddArea(parentAreaId, areaId, name);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL AddSource(int areaId, int sourceId, LPWSTR sourceName, bool multiSource)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->AddSource(areaId, sourceId, sourceName, multiSource);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL AddExistingSource(int areaId, int sourceId)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->AddExistingSource(areaId, sourceId);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL AddCondition(int sourceId, int conditionDefinitionId, int conditionId)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->AddCondition(sourceId, conditionDefinitionId, conditionId);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL ProcessSimpleEvent(int categoryId, int sourceId, LPWSTR message, int severity, int attributeCount, LPVARIANT attributeValues, LPFILETIME timeStamp)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->ProcessSimpleEvent(categoryId, sourceId, message, severity, attributeCount, attributeValues, timeStamp);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL ProcessTrackingEvent(int categoryId, int sourceId, LPWSTR message, int severity, LPWSTR actorId, int attributeCount, LPVARIANT attributeValues, LPFILETIME timeStamp)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->ProcessTrackingEvent(categoryId, sourceId, message, severity, actorId, attributeCount, attributeValues, timeStamp);
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL ProcessConditionStateChanges(int count, AeConditionState* conditionStateChanges)
{
#ifdef _OPC_SRV_AE
    HRESULT hResult;
    AeConditionChangeStates* cs = new AeConditionChangeStates[count];

    for (int i = 0; i < count; i++)
    {
        cs[i] = AeConditionChangeStates((conditionStateChanges[i]).CondID(), (conditionStateChanges[i]).SubCondID(), (conditionStateChanges[i]).ActiveState(), (conditionStateChanges[i]).Quality(), (conditionStateChanges[i]).AttrCount(), (conditionStateChanges[i]).AttrValuesPtr());
    }

    hResult = gpEventServer->ProcessConditionStateChanges(count, cs);
    delete cs;

    return hResult;
#else
    return E_NOTIMPL;
#endif
}

HRESULT DLLCALL AckCondition(int conditionId, LPWSTR comment)
{
#ifdef _OPC_SRV_AE
    return gpEventServer->AckCondition(conditionId, comment);
#else
    return E_NOTIMPL;
#endif
}

}