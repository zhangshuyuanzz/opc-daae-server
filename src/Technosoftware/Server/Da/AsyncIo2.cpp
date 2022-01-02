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
#include "FixOutArray.h"
#include "DataCallbackThread.h"
#include "Logger.h"

///////////////////////////////////////////////////////////////////////////
///////////////////////////////// IOPCAsyncIO2 ////////////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IOPCAsyncIO2::Read                                             INTERFACE
// ------------------
//    Reads one or more items in a group
//=========================================================================
STDMETHODIMP DaGroup::Read(
    /* [in] */                    DWORD             count,
    /* [size_is][in] */           OPCHANDLE*        serverHandle,
    /* [in] */                    DWORD             dwTransactionID,
    /* [out] */                   DWORD* pdwCancelID,
    /* [size_is][size_is][out] */ HRESULT** ppErrors)
{
    LOGFMTI("IOPCAsyncIO2::Read");

    return ReadAsync(count,
        serverHandle,
        NULL,
        dwTransactionID,
        pdwCancelID,
        ppErrors);
}



//=========================================================================
// IOPCAsyncIO2::Write                                            INTERFACE
// -------------------
//    Writes one or more Values for the specified Items asynchronously.
//=========================================================================
STDMETHODIMP DaGroup::Write(
    /* [in] */                    DWORD             dwCount,
    /* [size_is][in] */           OPCHANDLE* phServer,
    /* [size_is][in] */           VARIANT* pItemValues,
    /* [in] */                    DWORD             dwTransactionID,
    /* [out] */                   DWORD* pdwCancelID,
    /* [size_is][size_is][out] */ HRESULT** ppErrors)
{
    LOGFMTI("IOPCAsyncIO2::Write");
#ifdef WITH_SHALLOW_COPY

    HRESULT hr = S_OK;
    try {
        CAutoVectorPtr<OPCITEMVQT> avpVQTs;
        if (!avpVQTs.Allocate(dwCount)) throw E_OUTOFMEMORY;

        memset(avpVQTs, 0, sizeof(OPCITEMVQT) * dwCount);
        for (DWORD i = 0; i < dwCount; i++) {
            avpVQTs[i].vDataValue = pItemValues[i];  // Shallow copy
        }

        hr = WriteAsync(dwCount, phServer, avpVQTs,
            dwTransactionID, pdwCancelID,
            ppErrors);
    }
    catch (HRESULT hrEx) { hr = hrEx; }
    catch (...) { hr = E_FAIL; }
    return hr;
#else
    DWORD    i;
    HRESULT  hr = S_OK;
    CAutoVectorPtr<OPCITEMVQT> avpVQTs;
    try {

        if (!avpVQTs.Allocate(dwCount)) throw E_OUTOFMEMORY;

        memset(avpVQTs, 0, sizeof(OPCITEMVQT) * dwCount);
        for (i = 0; i < dwCount; i++) {
            hr = VariantCopy(&avpVQTs[i].vDataValue, &pItemValues[i]);   // Deep copy
            _OPC_CHECK_HR(hr);
        }

        hr = WriteAsync(dwCount, phServer, avpVQTs,
            dwTransactionID, pdwCancelID,
            ppErrors);

    }
    catch (HRESULT hrEx) { hr = hrEx; }
    catch (...) { hr = E_FAIL; }

    if (avpVQTs) {
        for (DWORD i = 0; i < dwCount; i++) {
            VariantClear(&avpVQTs[i].vDataValue);
        }
    }

    return hr;
#endif
}



//=========================================================================
// IOPCAsyncIO2::Refresh2                                         INTERFACE
// ----------------------
//    Forces a callback to IOPCDataCallback::OnDataChange for all
//    active items in the group.
//=========================================================================
STDMETHODIMP DaGroup::Refresh2(
    /* [in] */                    OPCDATASOURCE     dwSource,
    /* [in] */                    DWORD             dwTransactionID,
    /* [out] */                   DWORD* pdwCancelID)
{
    LOGFMTI("IOPCAsyncIO2::Refresh2");

    return Refresh2OrRefreshMaxAge(
        &dwSource,
        NULL,
        dwTransactionID,
        pdwCancelID);
}



//=========================================================================
// IOPCAsyncIO2::Cancel2                                          INTERFACE
// ---------------------
//    Request to cancel an outstanding transaction.
//=========================================================================
STDMETHODIMP DaGroup::Cancel2(
    /* [in] */                    DWORD             dwCancelID)
{
    HRESULT        hres;
    DaGenericGroup* pGGroup;

    hres = GetGenericGroup(&pGGroup);          // Check group state and get the pointer
    if (FAILED(hres)) {
        LOGFMTE("IOPCAsyncIO2::Cancel2: Internal error: No generic Group");
        return hres;
    }
    LOGFMTI("IOPCAsyncIO2::Cancel2 transaction with cancel ID %lu", dwCancelID);

    EnterCriticalSection(&pGGroup->m_AsyncThreadsCritSec);

    DataCallbackThread* pThreadToCancel;
    // Get the thread with the specified cancel ID
    hres = pGGroup->m_oaAsyncThread.GetElem(dwCancelID, (DaAsynchronousThread**)&pThreadToCancel);
    if (SUCCEEDED(hres)) {
        // OK, there is an outstanding transaction
        hres = pThreadToCancel->CreateCancelThread();
    }
    else {
        LOGFMTI("   It is 'too late' to cancel the transaction");
    }
    LeaveCriticalSection(&pGGroup->m_AsyncThreadsCritSec);
    ReleaseGenericGroup();
    return hres;
}



//=========================================================================
// IOPCAsyncIO2::SetEnable                                        INTERFACE
// -----------------------
//    Sets the callback enable value to FALSE or TRUE.
//=========================================================================
STDMETHODIMP DaGroup::SetEnable(
    /* [in] */                    BOOL              bEnable)
{
    DaGenericGroup* pGGroup;
    HRESULT        hres;

    hres = GetGenericGroup(&pGGroup);
    if (FAILED(hres)) {
        LOGFMTE("IOPCAsyncIO2::SetEnable: Internal error: No generic group.");
        return hres;
    }

    OPCWSTOAS(pGGroup->m_Name)
        LOGFMTI("IOPCAsyncIO2::SetEnable( %s ) for group %s", bEnable ? "TRUE" : "FALSE", OPCastr);
}

if (*m_vec.begin() == NULL) {                // There mus be a registered callback function
    LOGFMTE("SetEnable() failed with error: The client has not registered a callback through IConnectionPoint::Advise");
    hres = CONNECT_E_NOCONNECTION;
}
else {
    if (bEnable) {
        pGGroup->ResetKeepAliveCounter();
    }
    pGGroup->m_fCallbackEnable = bEnable;
}
ReleaseGenericGroup();
return hres;
}



//=========================================================================
// IOPCAsyncIO2::GetEnable                                        INTERFACE
// -----------------------
//    Retrieves the last callback enable value set with SetEnable().
//=========================================================================
STDMETHODIMP DaGroup::GetEnable(
    /* [out] */                   BOOL* pbEnable)
{
    DaGenericGroup* pGGroup;
    HRESULT        hres;

    hres = GetGenericGroup(&pGGroup);
    if (FAILED(hres)) {
        LOGFMTE("IOPCAsyncIO2::GetEnable: Internal error: No generic group.");
        return hres;
    }

    USES_CONVERSION;
    LOGFMTI("IOPCAsyncIO2::GetEnable for group %s", W2A(pGGroup->m_Name));

    if (*m_vec.begin() == NULL) {                // There mus be a registered callback function
        LOGFMTE("GetEnable() failed with error: The client has not registered a callback through IConnectionPoint::Advise");
        hres = CONNECT_E_NOCONNECTION;
    }
    else {
        *pbEnable = pGGroup->m_fCallbackEnable;
    }
    ReleaseGenericGroup();
    return hres;
}



///////////////////////////////////////////////////////////////////////////
///////////////////////////// IConnectionPoint ////////////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// IConnectionPoint::Advise                                       INTERFACE
// ------------------------
//    Overriden ATL implementation of IConnectionPoint::Advise.
//    The function checks if there is an existing IAdviseSink callback.
//=========================================================================
STDMETHODIMP DaGroup::Advise(IUnknown* pUnkSink, DWORD* pdwCookie)
{
    {
        IID iid;
        IOPCGroupConnectionPointImpl::GetConnectionInterface(&iid);
        if (IsEqualIID(iid, IID_IOPCDataCallback)) {
            LOGFMTI("IConnectionPoint::Advise, IID_IOPCDataCallback");
        }
        else {
            LOGFMTI("IConnectionPoint::Advise, Unknown IID");
        }
    }

    DaGenericGroup* pGGroup;
    HRESULT        hres;

    hres = GetGenericGroup(&pGGroup);
    if (SUCCEEDED(hres)) {

        EnterCriticalSection(&pGGroup->m_CallbackCritSec);

        if (pGGroup->m_DataTimeCallback ||
            pGGroup->m_DataCallback ||
            pGGroup->m_DataTimeCallbackDisp) {

            hres = CONNECT_E_ADVISELIMIT;
            LOGFMTE("Error: There is an existing IAdviseSink connection");
        }
        else {

            Lock();                                // Lock the connection point list

            hres = IOPCGroupConnectionPointImpl::Advise(pUnkSink, pdwCookie);

            if (SUCCEEDED(hres)) {
                // Register the callback interface in the global interface table
                hres = core_generic_main.m_pGIT->RegisterInterfaceInGlobal(pUnkSink, IID_IOPCDataCallback, &m_dwCookieGITDataCb);
                if (FAILED(hres)) {
                    IOPCGroupConnectionPointImpl::Unadvise(*pdwCookie);
                    pUnkSink->Release();             // Note :   register increments the refcount
                }                                   //          even if the function failed     
            }
            Unlock();                              // Unlock the connection point list

            if (SUCCEEDED(hres)) {               // Advise succeeded
                AddRef();                           // Prevents the release of the COM object if all interfaces will be released.
            }
        }

        LeaveCriticalSection(&pGGroup->m_CallbackCritSec);
        ReleaseGenericGroup();
    }
    return hres;
}



//=========================================================================
// IConnectionPoint::Unadvise                                     INTERFACE
// --------------------------
//    Overriden ATL implementation of IConnectionPoint::Unadvise.
//    The function clears the last read value of all generic
//    items. For this reason the values of all group items are sent
//    again if the next callback function will be registered.
//=========================================================================
STDMETHODIMP DaGroup::Unadvise(DWORD dwCookie)
{
    LOGFMTI("IConnectionPoint::Unadvise");

    Lock();                                      // Lock the connection point list

    HRESULT hresGIT = S_OK;
    hresGIT = core_generic_main.m_pGIT->RevokeInterfaceFromGlobal(m_dwCookieGITDataCb);

    HRESULT hres = IOPCGroupConnectionPointImpl::Unadvise(dwCookie);

    Unlock();                                    // Unlock the connection point list

    if (FAILED(hres)) {
        return hres;
    }

    Release();                                   // Unadvise succeeded
    // Permits the release of the COM object if all interfaces will be released.
    if (FAILED(hresGIT)) {
        return hresGIT;
    }

    DaGenericGroup* pGGroup;

    hres = GetGenericGroup(&pGGroup);
    if (SUCCEEDED(hres)) {
        pGGroup->ResetLastReadOfAllGenericItems();
        ReleaseGenericGroup();
    }
    return hres;
}



//=========================================================================
// Invoke IOPCDataCallback::OnDataChange with TransactionID = 0
//=========================================================================
HRESULT DaGroup::FireOnDataChange(DWORD dwNumOfItems, OPCITEMSTATE* pItemStates, HRESULT* errors)
{
    HRESULT           hres = S_OK;
    HRESULT           hrMasterError = S_OK;
    HRESULT           hrMasterQuality = S_OK;
    BOOL              fLocked = FALSE;
    IOPCDataCallback* pCallback = NULL;

    CFixOutArray< OPCHANDLE >  fxaHandles;
    CFixOutArray< VARIANT >    fxaVal;
    CFixOutArray< WORD >       fxaQual;
    CFixOutArray< FILETIME >   fxaTStamps;

    // Pointers to the arrays above
    OPCHANDLE* phClientItems = NULL;
    VARIANT* pvValues = NULL;
    WORD* pwQualities = NULL;
    FILETIME* pftTimeStamps = NULL;

    LOGFMTI("FireOnDataChange %d Items", dwNumOfItems);

    try {
        // Initialze all arrays passed via the registered callback to the client
        fxaHandles.Init(dwNumOfItems, &phClientItems);
        fxaVal.Init(dwNumOfItems, &pvValues);
        fxaQual.Init(dwNumOfItems, &pwQualities);
        fxaTStamps.Init(dwNumOfItems, &pftTimeStamps);

        // Sets the values in the arrays with the values from pItemStates
        DataCallbackThread::SetCallbackResultsFromItemStates(
            // IN
            dwNumOfItems,
            pItemStates,

            // IN / OUT
            &hrMasterError,
            errors,

            // OUT
            &hrMasterQuality,
            phClientItems,
            pvValues,
            pwQualities,
            pftTimeStamps);

        // Invoke the callback
        Lock();                                   // Lock the connection point list
        fLocked = TRUE;

        hres = GetCallbackInterface(&pCallback);
        if (SUCCEEDED(hres)) {

            pCallback->OnDataChange(
                0,                      // TransactionID is always 0 for exception based data changes
                // Client Group Handle
                m_pGroup->m_hClientGroupHandle,
                hrMasterQuality,        // Master Quality
                hrMasterError,          // Master Error
                dwNumOfItems,           // Number of Items
                phClientItems,          // List of Client Handles
                pvValues,               // List of Variant Values
                pwQualities,            // List of Quality Values
                pftTimeStamps,          // List of Time Stamps
                errors);              // List of Errors

            pCallback->Release();                  // All is done with this interface

            DaGenericGroup* pGGroup;
            if (SUCCEEDED(GetGenericGroup(&pGGroup))) {
                pGGroup->ResetKeepAliveCounter();
                ReleaseGenericGroup();
            }
        }
        Unlock();                                 // Unlock the connection point list
        fLocked = FALSE;
    }
    catch (HRESULT hresEx) {
        if (pCallback) {
            pCallback->Release();
        }
        if (fLocked) {
            Unlock();                              // Unlock the connection point list
        }
        hres = hresEx;
    }

    // Cleanup all arrays
    fxaHandles.Cleanup();
    fxaVal.Cleanup();
    fxaQual.Cleanup();
    fxaTStamps.Cleanup();

    return hres;
}



///////////////////////////////////////////////////////////////////////////
//////////////////////////////// Implementation ///////////////////////////
///////////////////////////////////////////////////////////////////////////

//=========================================================================
// GetCallbackInterface                                            INTERNAL
// --------------------
//    Gets a pointer to the IOPCDataCallback interface from the connection
//    point list or from the Global Interface Table.
//        
//    The GetCallbackInterface() method calls IUnknown::AddRef() on the
//    pointer obtained in the ppCallback parameter. It is the caller's
//    responsibility to call Release() on this pointer.
//=========================================================================
HRESULT DaGroup::GetCallbackInterface(IOPCDataCallback** ppCallback)
{
    HRESULT hres = CONNECT_E_NOCONNECTION;       // There is no registered callback function

    *ppCallback = NULL;

    IUnknown** pp = m_vec.begin();               // There can be only one registered data callback sink.
    if (*pp) {
        hres = core_generic_main.m_pGIT->GetInterfaceFromGlobal(m_dwCookieGITDataCb, IID_IOPCDataCallback, (LPVOID*)ppCallback);
    }
    return hres;
}
//DOM-IGNORE-END
