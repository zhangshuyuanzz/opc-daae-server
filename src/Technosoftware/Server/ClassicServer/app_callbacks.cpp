/*
 * Copyright (c) 2011-2020 Technosoftware GmbH. All rights reserved
 * Web: https://technosoftware.com
 *
 * Purpose:
 *
 * The Software is subject to the Technosoftware GmbH Source Code License Agreement,
 * which can be found here:
 * https://technosoftware.com/documents/Source_License_Agreement.pdf
 */

 //-----------------------------------------------------------------------------
 // INLCUDE
 //-----------------------------------------------------------------------------
#include "stdafx.h"                             // Generic server part headers
#include "CoreMain.h"

												// Application specific definitions
#ifdef NDEBUG
#include <stdio.h>
#endif
#include "DaServer.h"
#include "Logger.h"

//=============================================================================
// CreateOneItem
// -------------
//    Creates one item in the server data structure.
//    This function is called from the user DLL for each item created.
//=============================================================================
HRESULT		CreateOneItem(	LPWSTR						szItemID,
						  DWORD						dwAccessRights,
						  LPVARIANT					pvValue,
						  BOOL						fActive,
						  OPCEUTYPE					eEUType,
							LPVARIANT					pvEUInfo,
						  void**					deviceItem
)

{
	HRESULT           hres;
	DeviceItem*    pItem;

	pItem = new DeviceItem;                   // Allocate space for new server item.
	if (pItem == NULL) {
		hres = E_OUTOFMEMORY;
	}
	else {                                       // Initialize new Item instance
		hres = pItem->Create(szItemID, dwAccessRights, pvValue, fActive, 0, NULL, NULL, eEUType, pvEUInfo);
		if (SUCCEEDED(hres)) {                  // Add new created item to the item list
			hres = gpDataServer->AddDeviceItem(pItem);
		}
	}
	*deviceItem = (void*)pItem;
	VariantClear(pvValue);
	USES_CONVERSION;
	LOGFMTT("CreateOneItem() for item %s finished with hres = 0x%x.", W2A(szItemID), hres);
	return hres;
}

