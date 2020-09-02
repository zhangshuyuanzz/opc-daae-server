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

//-----------------------------------------------------------------------------
// INLCUDE
//-----------------------------------------------------------------------------
#include "stdafx.h"                             // Generic server part headers
#include "DaDeviceItem.h"
                                                // Application specific definitions
#include "DaServer.h"

using namespace IClassicBaseNodeManager;

extern bool 	 gUseOnAddItem;
extern bool 	 gUseOnRemoveItem;

//-----------------------------------------------------------------------------
// CODE
//-----------------------------------------------------------------------------

DeviceItem::DeviceItem()
{
   //
   // TODO : Initialize here your own item specific data
   //
}


DeviceItem::~DeviceItem()
{
   //
   // TODO : Release here your own item specific data
   //
}


//=========================================================================
// AttachActiveCount
//=========================================================================
HRESULT DeviceItem::AttachActiveCount( void )
{
	   HRESULT        hres;

   	hres = DaDeviceItem::AttachActiveCount();

	if (gUseOnAddItem)
	{
		if (DaDeviceItem::get_ActiveCount() > 1)
		{
			return hres;
		}

		OnAddItem(this);
	}
	return hres;
}




//=========================================================================
// DetachActiveCount
//=========================================================================
HRESULT DeviceItem::DetachActiveCount(void)
{
	   HRESULT        hres;

	hres = DaDeviceItem::DetachActiveCount();
	if (gUseOnRemoveItem)
	{
		if (DaDeviceItem::get_ActiveCount() == 0)
		{
			OnRemoveItem(this);
		}
	}
	return hres;
}
