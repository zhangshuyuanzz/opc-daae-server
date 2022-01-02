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
#include "DaComBaseServer.h"
#include "DaGenericServer.h"
#include "Logger.h"

/////////////////////////////////////////////////////////////////////////////
// DaComBaseServer

//=========================================================================
// DaComBaseServer Constructor
//
//=========================================================================
DaComBaseServer::DaComBaseServer()
{
    HRESULT res;

    LOGFMTI("New client connects to server");
    // initialize server if not yet done
    creationFailed_ = FALSE;
    daBaseServer_ = NULL;
    res = core_generic_main.InitializeServer();
    if (FAILED(res)) {
        creationFailed_ = TRUE;
    }
    daGenericServer_ = NULL;
}



//=========================================================================
// DaComBaseServer Destructor
//
//=========================================================================
DaComBaseServer::~DaComBaseServer()
{
    LPWSTR pName;
    if (SUCCEEDED(daBaseServer_->GetName(&pName))) {
		USES_CONVERSION;
		LOGFMTI("Client disconnects from Server '%s'", W2A(pName));
        delete pName;
    }
}



//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+  ...
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
HRESULT DaComBaseServer::FinalConstructBase()
{
    if (creationFailed_) {
        return E_FAIL;
    }

    HRESULT hres = OpcCommon::Create();        // Initialize base class
    if (FAILED(hres)) return hres;

    daGenericServer_ = new DaGenericServer();
    if (daGenericServer_ == NULL) {
        return E_OUTOFMEMORY;
    }

    hres = daGenericServer_->Create(daBaseServer_, this);
    if (SUCCEEDED(hres)) {
        hres = DaBrowse::Create(&daGenericServer_->m_BrowseData, daBaseServer_);
    }
    if (FAILED(hres)) {
        delete daGenericServer_;
        daGenericServer_ = NULL;
    }
    return hres;
}



//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+  ...
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
void DaComBaseServer::FinalReleaseBase()
{
    if (daGenericServer_ != NULL) {
        daGenericServer_->Kill(TRUE);
        daGenericServer_ = NULL;
    }
}

//DOM-IGNORE-END


