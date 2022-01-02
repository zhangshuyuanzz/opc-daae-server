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

//DOM-IGNORE-BEGIN

#include "stdafx.h"
#include <comdef.h>                             // for _variant_t
#include "DaItemProperty.h"


//=========================================================================
// Constructor
//=========================================================================
DaItemProperty::DaItemProperty( DWORD dwID, VARTYPE vt, LPWSTR pwszNewDescr)
{
	_ASSERTE(dwID != 0);                 // No invalid ID
	_ASSERTE(pwszNewDescr != NULL);      // Must be defined

	dwPropID = dwID;
	vtDataType = vt;
	pwszDescr = pwszNewDescr;
}


//=========================================================================
// Destructor
//=========================================================================
DaItemProperty::~DaItemProperty()
{
}

//DOM-IGNORE-END
