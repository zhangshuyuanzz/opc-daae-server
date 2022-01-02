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
#include "DaComServer.h"
#include "CoreMain.h"

/////////////////////////////////////////////////////////////////////////////
//                       Initialize Static Members
/////////////////////////////////////////////////////////////////////////////

DaComServer::_ThreadModel::AutoCriticalSection      DaComServer::criticalSectionCreate_;

DaComBaseServer*                                    DaComServer::creator_ = NULL;

DaBaseServer*                                       DaComServer::daBaseServerToConnect_ = NULL;




/////////////////////////////////////////////////////////////////////////////
//                       Internal DaComServer Instance Creation
/////////////////////////////////////////////////////////////////////////////


HRESULT DaComServer::CreateInstanceInternal(
    /* [in]          */ DaComBaseServer *       creator,
    /* [in]          */ DaBaseServer  *         daBaseServerToConnect,
    /* [in]          */ REFIID                  refIid,
    /* [iid_is][out] */ LPUNKNOWN *             unknown)
{
    _ASSERTE(unknown != NULL);                  // Must not be NULL
    _ASSERTE(daBaseServerToConnect != NULL);

    criticalSectionCreate_.Lock();              // Lock DaComServer object creation by (D)COM

    DaComServer::creator_ = creator;
    DaComServer::daBaseServerToConnect_ = daBaseServerToConnect;

    HRESULT hres = DaComServer::_CreatorClass::CreateInstance(  NULL,
                                                                refIid,
                                                                reinterpret_cast<LPVOID *>(unknown) );

    criticalSectionCreate_.Unlock();            // Unlock DaComServer object creation by (D)COM
    return hres;
}



/////////////////////////////////////////////////////////////////////////////
//                       DaComServer   CONSTRUCTOR
/////////////////////////////////////////////////////////////////////////////
DaComServer::DaComServer(void)
{
    unknownMarshaler_ = NULL;

    criticalSectionCreate_.Lock();              // Lock access of static member 'g_pServerHandler'

    if (DaComServer::creator_ == NULL) {
        // Created by (D)COM. The Default Server handler is used.

        // Since the first created instance is always connected to
        // the default Server Handler the class to identify server
        // instances must be created first.
        serverInstanceHandle_ = new DaServerInstanceHandle;
        if (serverInstanceHandle_ == NULL) {
            creationFailed_ = TRUE;
        }
        // Link to the default Server Instance.
        daBaseServer_ = OnGetDaBaseServer(NULL);
        if (daBaseServer_ == NULL) {
            creationFailed_ = TRUE;
        }
    }
    else {                                      // Created internal with CreateInstanceInternal(...)
                                                // (e.g by CALL-R Functions).

       // Now server instances created by the same client.

                                                // Use existing connection and attach this new created instance.
        serverInstanceHandle_ = DaComServer::creator_->serverInstanceHandle_;
        daBaseServer_ = DaComServer::daBaseServerToConnect_;
    }

    // Must be initialized again for each object creation if not
    // the default server is take.
    DaComServer::creator_ = NULL;
    DaComServer::daBaseServerToConnect_ = NULL;

    criticalSectionCreate_.Unlock();            // Unlock access of global member 'g_pCreator'
}
//DOM-IGNORE-END
