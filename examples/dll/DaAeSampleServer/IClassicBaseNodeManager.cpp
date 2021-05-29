/*
 * Copyright (c) 2011-2021 Technosoftware GmbH. All rights reserved
 * Web: https://technosoftware.com 
 * 
 * License:
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:

 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.

 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 * SPDX-License-Identifier: MIT
 */

/*
 * OPC Server SDK C++ generic server interface.
 *
 * IMPORTANT: DON'T CHANGE THIS FILE.
 */

//-----------------------------------------------------------------------------
// INCLUDES
//-----------------------------------------------------------------------------
#include "StdAfx.h"
#include "IClassicBaseNodeManager.h"

namespace IClassicBaseNodeManager
{
//-----------------------------------------------------------------------------
// GLOBALS (DON'T CHANGE)
//-----------------------------------------------------------------------------
AddItemPtr								addItemCallback;
RemoveItemPtr							removeItemCallback;
AddPropertyPtr							addPropertyCallback;
SetItemValuePtr							setItemValueCallback;
SetServerStatePtr						setServerStateCallback;
GetActiveItemsPtr						getActiveItemsCallback;

AddSimpleEventCategoryPtr				addSimpleEventCategoryCallback;
AddTrackingEventCategoryPtr				addTrackingEventCategoryCallback;
AddConditionEventCategoryPtr			addConditionEventCategoryCallback;
AddEventAttributePtr					addEventAttributeCallback;
AddSingleStateConditionDefinitionPtr	addSingleStateConditionDefinitionCallback;
AddMultiStateConditionDefinitionPtr		addMultiStateConditionDefinitionCallback;
AddSubConditionDefinitionPtr			addSubConditionDefinitionCallback;
AddAreaPtr								addAreaCallback;
AddSourcePtr							addSourceCallback;
AddExistingSourcePtr					addExistingSourceCallback;
AddConditionPtr							addConditionCallback;
ProcessSimpleEventPtr					processSimpleEventCallback;
ProcessTrackingEventPtr					processTrackingEventCallback;
ProcessConditionStateChangesPtr		    processConditionStateChangesCallback;
AckConditionPtr							ackConditionCallback;

FireShutdownRequestPtr                  fireShutdownRequestCallback;
GetClientsPtr                           getClientsCallback;
GetGroupsPtr                            getGroupsCallback;
GetGroupStatePtr                        getGroupStateCallback;
GetItemStatesPtr                        getItemStatesCallback;


//----------------------------------------------------------------------------
// These structures, enumerations and classes match the OPC specifications 
// or the server EXE file.
//----------------------------------------------------------------------------

//----------------------------------------------------------------------------
// OPC DA/AE Server SDK DLL API Callback Methods 
// (Called by the customization assembly)
//----------------------------------------------------------------------------

HRESULT AddItem(LPWSTR itemID, DaAccessRights accessRights, LPVARIANT initValue, void ** deviceItemHandle)
{
   return addItemCallback(itemID, accessRights, initValue, true, DaEuType::NoEnum, 0.0, 0.0, deviceItemHandle);
}

HRESULT AddAnalogItem(LPWSTR itemID, DaAccessRights accessRights, LPVARIANT initValue, double minValue, double maxValue, void ** deviceItemHandle)
{
	return addItemCallback(itemID, accessRights, initValue, true, DaEuType::Analog, minValue, maxValue, deviceItemHandle);
}

HRESULT RemoveItem(void* deviceItemHandle)
{
	return removeItemCallback(deviceItemHandle);
}

HRESULT AddProperty ( int propertyID, LPWSTR description, LPVARIANT valueType  )
{
    return addPropertyCallback(propertyID, description, valueType);
}

HRESULT SetItemValue(void* deviceItemHandle, LPVARIANT newValue, short quality, FILETIME timestamp)
{
	return setItemValueCallback(deviceItemHandle, newValue, quality, timestamp);
}

void SetServerState( ServerState serverState )
{
    setServerStateCallback(serverState);
}

void GetActiveItems( int * dwNumItemHandles, void * **  appItemHandles)
{
    getActiveItemsCallback(dwNumItemHandles, appItemHandles);
}

void GetClients(int * numClientHandles, void* ** clientHandles, LPWSTR ** clientNames)
{
    getClientsCallback(numClientHandles, clientHandles, clientNames);
}

void GetGroups(void * clientHandle, int * numGroupHandles, void* ** groupHandles, LPWSTR ** groupNames)
{
    getGroupsCallback(clientHandle, numGroupHandles, groupHandles, groupNames);
}

void GetGroupState(void * groupHandle, DaGroupState * groupState)
{
    getGroupStateCallback(groupHandle, groupState);
}

void GetItemStates(void * groupHandle, int * numDaItemStates, DaItemState* * daItemStates)
{
    getItemStatesCallback(groupHandle, numDaItemStates, daItemStates);
}

void FireShutdownRequest(LPCWSTR reason)
{
    fireShutdownRequestCallback(reason);
};
   
DLLEXP HRESULT DLLCALL OnDefineDaCallbacks(
						AddItemPtr			        addItem, 
						RemoveItemPtr		        removeItem, 
						AddPropertyPtr		        addProperty, 
						SetItemValuePtr		        setItemValue, 
						SetServerStatePtr	        setServerState,
						GetActiveItemsPtr	        getActiveItems,
                        FireShutdownRequestPtr      fireShutdownRequest,
                        GetClientsPtr               getClients,
                        GetGroupsPtr                getGroups,
                        GetGroupStatePtr            getGroupState,
                        GetItemStatesPtr            getItemStates )
{
    addItemCallback = addItem;
	removeItemCallback = removeItem;
	addPropertyCallback = addProperty;
	setItemValueCallback = setItemValue;
	setServerStateCallback = setServerState;
	getActiveItemsCallback = getActiveItems;
    fireShutdownRequestCallback = fireShutdownRequest;
    getClientsCallback = getClients;
    getGroupsCallback = getGroups;
    getGroupStateCallback = getGroupState;
    getItemStatesCallback = getItemStates;
    return S_OK;
}


DLLEXP HRESULT DLLCALL OnDefineAeCallbacks( 
						AddSimpleEventCategoryPtr				addSimpleEventCat, 
						AddTrackingEventCategoryPtr				addTrackingEventCat,
						AddConditionEventCategoryPtr			addConditionEventCat, 
						AddEventAttributePtr					addEventAttr,
						AddSingleStateConditionDefinitionPtr	addSingleStateConditionDef,  
						AddMultiStateConditionDefinitionPtr		addMultiStateConditionDef,
						AddSubConditionDefinitionPtr			addSubConditionDef, 
						AddAreaPtr								addArea, 
						AddSourcePtr							addSource, 
						AddExistingSourcePtr					addExistingSource,
						AddConditionPtr							addCondition, 
						ProcessSimpleEventPtr					processSimpleEvent, 
						ProcessTrackingEventPtr					processTrackingEvent,
						ProcessConditionStateChangesPtr		    processConditionStateChanges, 
						AckConditionPtr							ackCondition )
{
	addSimpleEventCategoryCallback = addSimpleEventCat;
	addTrackingEventCategoryCallback = addTrackingEventCat;
	addConditionEventCategoryCallback = addConditionEventCat;
	addEventAttributeCallback = addEventAttr;
	addSingleStateConditionDefinitionCallback = addSingleStateConditionDef;
	addMultiStateConditionDefinitionCallback = addMultiStateConditionDef;
	addSubConditionDefinitionCallback = addSubConditionDef;
	addAreaCallback = addArea;
	addSourceCallback = addSource;
	addExistingSourceCallback = addExistingSource;
	addConditionCallback = addCondition;
	processSimpleEventCallback = processSimpleEvent;
	processTrackingEventCallback = processTrackingEvent;
	processConditionStateChangesCallback = processConditionStateChanges;
	ackConditionCallback = ackCondition;

    return S_OK;
}

HRESULT AddSimpleEventCategory( int categoryID, LPWSTR categoryDescription)
{
    return addSimpleEventCategoryCallback(categoryID, categoryDescription);
}

HRESULT AddTrackingEventCategory( int categoryID, LPWSTR categoryDescription)
{
	return addTrackingEventCategoryCallback(categoryID, categoryDescription);
}

HRESULT AddConditionEventCategory( int categoryID, LPWSTR categoryDescription)
{
    return addConditionEventCategoryCallback(categoryID, categoryDescription);
}

HRESULT AddEventAttribute( int categoryID, int eventAttribute, LPWSTR attributeDescription, VARTYPE dataType )
{
    return addEventAttributeCallback(categoryID, eventAttribute, attributeDescription, dataType);
}

HRESULT AddSingleStateConditionDefinition( int categoryId, int conditionId, LPWSTR name, LPWSTR condition, int severity, LPWSTR description, bool ackRequired )
{
	return addSingleStateConditionDefinitionCallback( categoryId, conditionId, name, condition, severity, description, ackRequired );
}

HRESULT AddMultiStateConditionDefinition( int categoryId, int conditionId, LPWSTR name )
{
	return addMultiStateConditionDefinitionCallback( categoryId, conditionId, name );
}

HRESULT AddSubConditionDefinition( int conditionId, int subConditionId, LPWSTR name, LPWSTR condition, int severity, LPWSTR description, bool ackRequired )
{
	return addSubConditionDefinitionCallback( conditionId, subConditionId, name, condition, severity, description, ackRequired );
}

HRESULT AddArea( int parentAreaId, int areaId, LPWSTR name )
{
	return addAreaCallback( parentAreaId, areaId, name );
}

HRESULT AddSource( int areaId, int sourceId, LPWSTR sourceName, bool multiSource )
{
	return addSourceCallback( areaId, sourceId, sourceName, multiSource );
}

HRESULT AddExistingSource( int areaId, int sourceId )
{
	return addExistingSourceCallback( areaId, sourceId );
}

HRESULT AddCondition( int sourceId, int conditionDefinitionId, int conditionId )
{
	return addConditionCallback( sourceId, conditionDefinitionId, conditionId );
}

HRESULT ProcessSimpleEvent( int categoryId, int sourceId, LPWSTR message, int severity, int attributeCount, LPVARIANT attributeValues, LPFILETIME timeStamp )
{
	return processSimpleEventCallback( categoryId, sourceId, message, severity, attributeCount, attributeValues, timeStamp );
}

HRESULT ProcessTrackingEvent( int categoryId, int sourceId, LPWSTR message, int severity, LPWSTR actorId, int attributeCount, LPVARIANT attributeValues, LPFILETIME timeStamp )
{
	return processTrackingEventCallback( categoryId, sourceId, message, severity, actorId, attributeCount, attributeValues, timeStamp );
}

HRESULT ProcessConditionStateChanges(int count, AeConditionState* AeConditionStateChanges)
{
    return processConditionStateChangesCallback(count, AeConditionStateChanges);
}

HRESULT AckCondition( int conditionId, LPWSTR comment )
{
	return ackConditionCallback( conditionId, comment );
}

}