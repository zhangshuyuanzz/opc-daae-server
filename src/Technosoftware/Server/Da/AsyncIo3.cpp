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

 //-----------------------------------------------------------------------
 // INLCUDE
 //-----------------------------------------------------------------------
#include "stdafx.h"
#include "DaGroup.h"
#include "DaGenericItem.h"
#include "FixOutArray.h"
#include "DataCallbackThread.h"
#include "Logger.h"

///////////////////////////////////////////////////////////////////////////
//////////////////////////// IOPCAsyncIO3 /////////////////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCAsyncIO3::ReadMaxAge                                       INTERFACE
// ------------------------
//    Reads one or more Values, Qualities and TimeStamps for
//    the specified Items asynchronously.
//=========================================================================
STDMETHODIMP DaGroup::ReadMaxAge(
    /* [in] */                    DWORD             dwCount,
    /* [size_is][in] */           OPCHANDLE      *  phServer,
    /* [size_is][in] */           DWORD          *  pdwMaxAge,
    /* [in] */                    DWORD             dwTransactionID,
    /* [out] */                   DWORD          *  pdwCancelID,
    /* [size_is][size_is][out] */ HRESULT        ** ppErrors)
{
    LOGFMTI("IOPCAsyncIO3::ReadMaxAge");

    return ReadAsync(dwCount,
        phServer,
        pdwMaxAge,
        dwTransactionID,
        pdwCancelID,
        ppErrors);
}



//=========================================================================
// IOPCAsyncIO3::WriteVQT                                         INTERFACE
// ----------------------
//    Writes one or more Values, Qualities and TimeStamps for
//    the specified Items asynchronously.
//=========================================================================
STDMETHODIMP DaGroup::WriteVQT(
    /* [in] */                    DWORD             dwCount,
    /* [size_is][in] */           OPCHANDLE      *  phServer,
    /* [size_is][in] */           OPCITEMVQT     *  pItemVQT,
    /* [in] */                    DWORD             dwTransactionID,
    /* [out] */                   DWORD          *  pdwCancelID,
    /* [size_is][size_is][out] */ HRESULT        ** ppErrors)
{
    LOGFMTI("IOPCAsyncIO3::WriteVQT");

    return WriteAsync(dwCount,
        phServer,
        pItemVQT,
        dwTransactionID,
        pdwCancelID,
        ppErrors);
}



//=========================================================================
// IOPCAsyncIO3::RefreshMaxAge                                    INTERFACE
// ---------------------------
//    Forces a callback to IOPCDataCallback::OnDataChange for all
//    active items in the group. The MaxAge value will determine where
//    the data is obtained.
//=========================================================================
STDMETHODIMP DaGroup::RefreshMaxAge(
    /* [in] */                    DWORD             dwMaxAge,
    /* [in] */                    DWORD             dwTransactionID,
    /* [out] */                   DWORD          *  pdwCancelID)
{
    LOGFMTI("IOPCAsyncIO3::RefreshMaxAge");

    return Refresh2OrRefreshMaxAge(
        NULL,
        &dwMaxAge,
        dwTransactionID,
        pdwCancelID);
}

///////////////////////////////////////////////////////////////////////////


//-------------------------------------------------------------------------
// IMPLEMENTATTION
//-------------------------------------------------------------------------

//=========================================================================
// ReadAsync                                                      INTERFACE
// ---------
//    Implementation for IOPCAsyncIO2::Read() and
//    IOPCAsyncIO3::ReadMaxAge().
//
//    Parameter pdwMaxAge may be NULL.
//    If pdwMaxAge is NULL then the function is called from
//    for IOPCAsyncIO2::Read(); otherwise from IOPCAsyncIO3::ReadMaxAge().
//=========================================================================
HRESULT DaGroup::ReadAsync(
    /* [in] */                    DWORD             dwCount,
    /* [size_is][in] */           OPCHANDLE      *  phServer,
    /* [size_is][in] */           DWORD          *  pdwMaxAge,
    /* [in] */                    DWORD             dwTransactionID,
    /* [out] */                   DWORD          *  pdwCancelID,
    /* [size_is][size_is][out] */ HRESULT        ** ppErrors)
{
    HRESULT        hrRet;                     // Returned result
    DaGenericGroup* pGGroup;

    *pdwCancelID = 0;
    *ppErrors = NULL;

    hrRet = GetGenericGroup(&pGGroup);         // Check group state and get the pointer
    if (FAILED(hrRet)) {
        LOGFMTE("ReadAsync() failed with error: No generic Group");
        return hrRet;
    }

    DaDeviceItem**  ppDItems = NULL;           // Array of Device Item Pointers
    OPCITEMSTATE*  pItemStates = NULL;           // Array of OPC Item States
    DWORD*         pdwMaxAgeDefs = NULL;         // Array of Max Age definitions which corresponds
                                                 // to the valid items
    DWORD          dwCountValid = 0;              // Number of valid items

    try {
        USES_CONVERSION;
        LOGFMTI("ReadAsync() %ld items of group %s", dwCount, W2A(pGGroup->m_Name));

        // Check arguments
        if (dwCount == 0) {                       // There must be items
            LOGFMTE("ReadAsync() failed with invalid argument(s): dwCount is 0");
            throw E_INVALIDARG;
        }
        if (*m_vec.begin() == NULL) {             // There mus be a registered callback function
            LOGFMTE("ReadAsync() failed with error: The client has not registered a callback through IConnectionPoint::Advise");
            throw CONNECT_E_NOCONNECTION;
        }

        // Get the device items associated with the specified server item handles
        hrRet = pGGroup->GetDItemsAndStates(
            dwCount,                // Number of requested items
            phServer,               // Server item handles of the requested items
            pdwMaxAge,              // Max Ages Input
            OPC_READABLE,           // Returened items must have read access
            ppErrors,
            &dwCountValid,
            &ppDItems,
            &pItemStates,
            &pdwMaxAgeDefs          // Max Ages Output
            );
        _OPC_CHECK_HR(hrRet);
        if (dwCountValid == 0) throw S_FALSE;     // If there are no valid items then the error array
                                                  // is returned but no callback must be invoked

        // Create the data callback object
        DataCallbackThread* pThread = new DataCallbackThread(pGGroup);
        _OPC_CHECK_PTR(pThread);

        // Create the thread for the read operation
        HRESULT hrRead;                           // Store the Custom Read result in a own
                                                  // variable and do not overwrite the result from
                                                  // GetDItemsAndStates()

        if (pdwMaxAge) {                          // Called by IOPCAsyncIO3::ReadMaxAge()
            hrRead = pThread->CreateCustomReadMaxAge(
                dwCountValid,           // Number of items to read
                dwTransactionID,
                ppDItems,               // Array of DeviceItems
                pItemStates,            // OPCITEMSTATE result array 
                pdwMaxAgeDefs,          // Array with Max Age definitions
                NULL,
                pdwCancelID);          // Out parameter

        }
        else {                                    // Called by IOPCAsyncIO2::Read()
            hrRead = pThread->CreateCustomRead(
                dwCountValid,           // Number of items to read
                dwTransactionID,
                ppDItems,               // Array of DeviceItems
                pItemStates,            // OPCITEMSTATE result array 
                NULL,
                pdwCancelID);          // out parameter
        }

        if (FAILED(hrRead)) {
            delete pThread;
            throw hrRead;
        }

        // Note :   Do no longer use the pointer pThread if the Create() function
        //          succeeded because the object deletes itself.
        //          Also all provided arrays will be released.

        // 'hrRet' is used as return value. The value is either S_OK or S_FALSE.
    }
    catch (HRESULT hrEx) {
        if (pItemStates) {
            delete[] pItemStates;
        }
        if (ppDItems) {
            while (dwCountValid--) {
                ppDItems[dwCountValid]->Detach();
            }
            delete[] ppDItems;
        }
        if (pdwMaxAgeDefs) {
            delete[] pdwMaxAgeDefs;
        }
        if (hrEx != S_FALSE) {                    // The error array will be returned on S_FALSE
            ComFree(*ppErrors);
            *ppErrors = NULL;
        }
        *pdwCancelID = 0;
        hrRet = hrEx;
    }
    ReleaseGenericGroup();
    return hrRet;
}



//=========================================================================
// WriteAsync                                                      INTERNAL
// ----------
//    Implementation for IOPCAsyncIO2::Write() and
//    IOPCAsyncIO3::WriteVQT().
//=========================================================================
HRESULT DaGroup::WriteAsync(
    /* [in] */                    DWORD             dwCount,
    /* [size_is][in] */           OPCHANDLE      *  phServer,
    /* [size_is][in] */           OPCITEMVQT     *  pItemVQT,
    /* [in] */                    DWORD             dwTransactionID,
    /* [out] */                   DWORD          *  pdwCancelID,
    /* [size_is][size_is][out] */ HRESULT        ** ppErrors)
{
    HRESULT        hrRet;                     // Returned result
    HRESULT        hrTmp;                     // Temporary results
    DaGenericGroup* pGGroup;

    *pdwCancelID = 0;
    *ppErrors = NULL;

    hrRet = GetGenericGroup(&pGGroup);         // Check group state and get the pointer
    if (FAILED(hrRet)) {
        LOGFMTE("WriteAsync() failed with error: No generic Group");
        return hrRet;
    }

    DaDeviceItem**  ppDItems = NULL;           // Array of Device Item Pointers
    OPCITEMSTATE*  pItemStates = NULL;           // Array of OPC Item States
    OPCITEMVQT*    pVQTsToWrite = NULL;          // Array of Values, Qualities and TimeStamps to write
    DWORD          dwCountValid = 0;              // Number of valid items

    try {
        USES_CONVERSION;
        LOGFMTI("WriteAsync() %ld items of group %s", dwCount, W2A(pGGroup->m_Name));

        // Check arguments
        if (dwCount == 0) {                       // There must be items
            LOGFMTE("WriteAsync() failed with invalid argument(s): dwCount is 0");
            throw E_INVALIDARG;
        }
        if (*m_vec.begin() == NULL) {             // There mus be a registered callback function
            LOGFMTE("WriteAsync() failed with error: The client has not registered a callback through IConnectionPoint::Advise");
            throw CONNECT_E_NOCONNECTION;
        }

        // Get the device items associated with the specified server item handles
        hrRet = pGGroup->GetDItemsAndStates(
            dwCount,                // Number of requested items
            phServer,               // Server item handles of the requested items
            NULL,                   // No Max Ages Input
            OPC_WRITEABLE,          // Returned items must have write access
            ppErrors,
            &dwCountValid,
            &ppDItems,
            &pItemStates,
            NULL                    // No Max Ages Output
            );
        _OPC_CHECK_HR(hrRet);
        if (dwCountValid == 0) throw S_FALSE;     // If there are no valid items then the error array
                                                  // is returned but no callback must be invoked

                                                  // Use only the values for valid items
        hrTmp = pGGroup->CopyVQTArrayForValidItems(
            dwCount, *ppErrors, pItemVQT,
            dwCountValid, &pVQTsToWrite);
        _OPC_CHECK_HR(hrTmp);

        // Create the data callback object
        DataCallbackThread* pThread = new DataCallbackThread(pGGroup);
        _OPC_CHECK_PTR(pThread);

        // Create the thread for the read operation
        hrTmp = pThread->CreateCustomWriteVQT(
            dwCountValid,           // Number of items to write
            dwTransactionID,
            ppDItems,               // Array of DeviceItems
            pItemStates,            // OPCITEMSTATE result array 
            pVQTsToWrite,           // Array of Values, Qualities and TimeStamps to write
            NULL,
            pdwCancelID);          // Out parameter

        if (FAILED(hrTmp)) {
            delete pThread;
            throw hrTmp;
        }

        // Note :   Do no longer use the pointer pThread if the Create() function
        //          succeeded because the object deletes itself.
        //          Also all provided arrays will be released.

        // 'hrRet' is used as return value. The value is either S_OK or S_FALSE.
    }
    catch (HRESULT hrEx) {
        if (pItemStates) {
            delete[] pItemStates;
        }
        DWORD i;
        if (ppDItems) {
            i = dwCountValid;
            while (i--) {
                ppDItems[i]->Detach();
            }
            delete[] ppDItems;
        }
        if (pVQTsToWrite) {
            i = dwCountValid;
            while (i--) {
                VariantClear(&pVQTsToWrite[i].vDataValue);
            }
            delete[] pVQTsToWrite;
        }
        if (hrEx != S_FALSE) {                    // The error array will be returned on S_FALSE
            ComFree(*ppErrors);
            *ppErrors = NULL;
        }
        *pdwCancelID = 0;
        hrRet = hrEx;
    }
    ReleaseGenericGroup();
    return hrRet;
}



//=========================================================================
// Refresh2OrRefreshMaxAge                                         INTERNAL
// -----------------------
//    Implementation for IOPCAsyncIO2::Refresh2() and
//    IOPCAsyncIO3::RefreshMaxAge().
//
//    Either parameter pdwMaxAge or pdwSource must be specified, the other
//    parameter must be NULL. Set pdwSource if the function is called from
//    IOPCAsyncIO2::Refresh2() and pdwMaxAge if the function is called
//    from IOPCAsyncIO3::RefreshMaxAge().
//=========================================================================
HRESULT DaGroup::Refresh2OrRefreshMaxAge(
    /* [in] */                    OPCDATASOURCE  *  pdwSource,
    /* [in] */                    DWORD          *  pdwMaxAge,
    /* [in] */                    DWORD             dwTransactionID,
    /* [out] */                   DWORD          *  pdwCancelID)
{
    // Either the MaxAge or the Source must be specified
    if (pdwSource)       _ASSERTE(!pdwMaxAge);
    else if (pdwMaxAge)  _ASSERTE(!pdwSource);
    else                 _ASSERTE(FALSE);

    HRESULT        hr;
    DaGenericGroup* pGGroup;

    *pdwCancelID = 0;

    hr = GetGenericGroup(&pGGroup);            // Check group state and get the pointer
    if (FAILED(hr)) {
        LOGFMTI("Refresh2OrRefreshMaxAge() failed with error: No generic Group");
        return hr;
    }

    DaDeviceItem**  ppDItems = NULL;           // Array of Device Item Pointers
    OPCITEMSTATE*  pItemStates = NULL;           // Array of OPC Item States
    HRESULT*       errors = NULL;
    DWORD          dwCountValid = 0;              // Number of valid items
    CFixOutArray< DWORD, FALSE >  apMaxAges;     // Use heap memory

    try {
        USES_CONVERSION;
        LOGFMTI("Refresh2OrRefreshMaxAge() items of group %s", W2A(pGGroup->m_Name));

        // Check arguments
        if (pdwSource) {
            if ((*pdwSource != OPC_DS_CACHE) && (*pdwSource != OPC_DS_DEVICE)) {
                LOGFMTE("Refresh2OrRefreshMaxAge() failed with invalid argument(s): dwSource with unknown type");
                throw E_INVALIDARG;
            }
        }
        if (*m_vec.begin() == NULL) {             // There mus be a registered callback function
            LOGFMTE("Refresh2OrRefreshMaxAge() failed with error: The client has not registered a callback through IConnectionPoint::Advise");
            throw CONNECT_E_NOCONNECTION;
        }
        if (pGGroup->GetActiveState() == FALSE) {
            LOGFMTE("Refresh2OrRefreshMaxAge() failed with error: The group must be active");
            throw E_FAIL;                          // The group must be active
        }

        // Get the device items associated with all active items of the group
        hr = pGGroup->GetDItemsAndStates(
            &errors,
            &dwCountValid,
            &ppDItems,
            &pItemStates
            );

        if (FAILED(hr))       throw hr;
        if (dwCountValid == 0)  throw E_FAIL;     // All readable items in the group are inactive

        // Create the data callback object
        DataCallbackThread* pThread = new DataCallbackThread(pGGroup);
        if (pThread == NULL)  throw E_OUTOFMEMORY;

        // Specifiy the source to be use
        OPCDATASOURCE dwSource;
        if (pdwMaxAge) {
            // Check if Max Age specifies 'Read from Device' or 'Read from Cache'.
            // If so we use the same functionality like IOPCAsyncIO2::Refresh2() for
            // reasons of performance.
            if (*pdwMaxAge == 0) {
                pdwSource = &(dwSource = OPC_DS_DEVICE);
                pdwMaxAge = NULL;
            }
            else if (*pdwMaxAge == 0xFFFFFFFF) {
                pdwSource = &(dwSource = OPC_DS_CACHE);
                pdwMaxAge = NULL;
            }
            // else: use the Max Age value
        }
        else {
            dwSource = *pdwSource;
        }


        // Create the thread for the refresh operation
        if (pdwMaxAge) {                          // IOPCAsyncIO3::RefreshMaxAge()

            DWORD*   parMaxAges;
            apMaxAges.Init(dwCountValid, &parMaxAges, *pdwMaxAge);

            hr = pThread->CreateCustomRefreshMaxAge(
                dwCountValid,           // Number of items to read
                dwTransactionID,
                ppDItems,               // Array of DeviceItems
                pItemStates,            // OPCITEMSTATE result array 
                parMaxAges,             // Array with Max Age definitions
                NULL,
                pdwCancelID);          // out parameter
        }
        else {                                    // Called by IOPCAsyncIO2::Refresh2() or by
                                                  // IOPCAsyncIO3::RefreshMaxAge() with Max Age 0xFFFFFFFF or 0.
            hr = pThread->CreateCustomRefresh(
                dwSource,
                dwCountValid,           // Number of items to read
                dwTransactionID,
                ppDItems,               // Array of DeviceItems
                pItemStates,            // OPCITEMSTATE result array 
                NULL,
                pdwCancelID);          // out parameter
        }

        if (FAILED(hr)) {
            delete pThread;
            throw hr;
        }

        // Note :   Do no longer use the pointer pThread if the Create() function
        //          succeeded because the object deletes itself.
        //          Also all provided arrays will be released.
    }
    catch (HRESULT hrEx) {
        if (pItemStates) {
            delete[] pItemStates;
        }
        if (ppDItems) {
            while (dwCountValid--) {
                ppDItems[dwCountValid]->Detach();
            }
            delete[] ppDItems;
        }
        apMaxAges.Cleanup();
        *pdwCancelID = 0;
        hr = hrEx;
    }
    ReleaseGenericGroup();

    if (errors) {                                  // Allocated from COM memory by GetDItemsAndStates()
        pIMalloc->Free(errors);
    }
    return hr;
}
//DOM-IGNORE-END
