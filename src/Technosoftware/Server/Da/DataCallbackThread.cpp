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
#include <process.h>
#include "DataCallbackThread.h"
#include "Logger.h"

//-----------------------------------------------------------------------
// CODE
//-----------------------------------------------------------------------

//=======================================================================
// Constructor
//=======================================================================
DataCallbackThread::DataCallbackThread(DaGenericGroup* pParent)
{
    _ASSERTE(pParent);                         // Must not be NULL

    // Attach to the group so that it cannot be deleted until
    // all the threads running over it are killed
    pParent->Attach();

    m_pGServer = pParent->m_pServer;
    _ASSERTE(m_pGServer);

    // Initialize members with default values
    m_hThread = 0;
    m_dwCount = 0;                    // Number of items
    m_dwTransactionID = 0;                    // Transaction ID
    m_dwSource = OPC_DS_DEVICE;        // Default data source
    m_fReadTransaction = TRUE;                 // The activated transaction type (read or refresh)

    m_pItemStates = NULL;                 // Array which the Device Item States
    m_ppDItems = NULL;                 // Array with Device Items to be read
    m_pfPhyval = NULL;                 // Array which marks items added with their physical value
    m_pVQTsToWrite = NULL;                 // Array with Values, Qualities and TimeStamps to write
    m_pdwMaxAge = NULL;                 // Array with Max Age definitions
    m_ppTmpBufferForItemsToReadFromDevice = NULL;// Array with temporary used Device Item pointers

    m_pParent = pParent;              // A generic group
    m_pCancelThread = NULL;
}



//=======================================================================
// Destructor
//=======================================================================
DataCallbackThread::~DataCallbackThread()
{
    // Release external memory allocated from the heap
    if (m_pfPhyval) {
        delete[] m_pfPhyval;
    }
    DWORD i;

    if (m_pItemStates) {
        for (i = 0; i < m_dwCount; i++) {
            VariantClear(&m_pItemStates[i].vDataValue);
        }
        delete[] m_pItemStates;
    }
    if (m_ppDItems) {
        for (i = 0; i < m_dwCount; i++) {
            if (m_ppDItems[i]) {
                m_ppDItems[i]->Detach();            // Release the items
            }
        }
        delete[] m_ppDItems;
    }
    if (m_pVQTsToWrite) {
        for (i = 0; i < m_dwCount; i++) {
            VariantClear(&m_pVQTsToWrite[i].vDataValue);
        }
        delete[] m_pVQTsToWrite;
    }
    if (m_pdwMaxAge) {
        delete[] m_pdwMaxAge;
    }
    if (m_ppTmpBufferForItemsToReadFromDevice) {
        delete[] m_ppTmpBufferForItemsToReadFromDevice;
    }

    // Release internal memory allocated from the COM memory manager
    m_fxaHandles.Cleanup();
    m_fxaVal.Cleanup();
    m_fxaQual.Cleanup();
    m_fxaTStamps.Cleanup();
    m_fxaErrors.Cleanup();

    m_pParent->Detach();                          // Now it possible to delete the parent
}



//-----------------------------------------------------------------------
// OPERATIONS
//-----------------------------------------------------------------------

//=========================================================================
// Creates a thread to cancel an asynchronous transaction
//=========================================================================
HRESULT DataCallbackThread::CreateCancelThread()
{
    m_pCancelThread = new CCancelThread;
    if (m_pCancelThread == NULL) {
        return E_OUTOFMEMORY;
    }
    return m_pCancelThread->Create(this);
}



//=========================================================================
// Starts an asynchronous Read or Refresh transaction.
//=========================================================================
HRESULT DataCallbackThread::CreateCustomReadOrRefresh(
    /* [in] */                    DWORD             dwCount,          // Number of items to handle
    /* [in] */                    DWORD             dwTransactionID,  // Transaction ID
    /* [size_is][in] */           DaDeviceItem    ** ppDItems,         // Array with device item ptrs  
    /* [size_is][in] */           OPCITEMSTATE   *  pItemStates,      // Client handles + requested data types
    /* [size_is][in] */           DWORD          *  pdwMaxAge,        // Array with Max Age definitions, may be a NULL pointer
    /* [size_is][in] */           BOOL           *  pfPhyval,         // Array which marks items added with their physical value, may be a NULL pointer
    /* [out] */                   DWORD          *  pdwCancelID)
{
    _ASSERTE(dwCount);                         // Items must be specified
    _ASSERTE(ppDItems);
    _ASSERTE(pItemStates);
    _ASSERTE(pdwCancelID);
    // Note : pdwMaxAge may be NULL
    // Note : pfPhyval may be NULL

    HRESULT  hr = S_OK;
    BOOL     fLocked = FALSE;

    try {
        // Initialze all arrays passed via the registered callback to the client
        m_fxaHandles.Init(dwCount, &m_phClientItems);
        m_fxaVal.Init(dwCount, &m_pvValues);
        m_fxaQual.Init(dwCount, &m_pwQualities);
        m_fxaTStamps.Init(dwCount, &m_pftTimeStamps);
        m_fxaErrors.Init(dwCount, &m_pErrors);

        m_dwCount = dwCount;              // Number of items
        m_dwTransactionID = dwTransactionID;      // Number of items
        m_ppDItems = ppDItems;             // Device Items to be read
        m_pItemStates = pItemStates;          // Client handles + requested data types
        m_pdwMaxAge = pdwMaxAge;            // Max Age definitions
        m_pfPhyval = pfPhyval;             // Array which marks items added with their physical value

        if (pdwMaxAge) {
            m_ppTmpBufferForItemsToReadFromDevice = new DaDeviceItem*[dwCount];
            _OPC_CHECK_PTR(m_ppTmpBufferForItemsToReadFromDevice);

            HRESULT hr = CoFileTimeNow(&m_ftNow);
            _OPC_CHECK_HR(hr);
        }
        // Lock the list before create the thread.
        // This way remove requests are prevented
        // before the thread was added to the list.
        EnterCriticalSection(&m_pParent->m_AsyncThreadsCritSec);
        fLocked = TRUE;
        // The array index also used as cancel ID.
        m_dwCancelID = *pdwCancelID = m_pParent->m_oaAsyncThread.New();

        hr = m_pParent->m_oaAsyncThread.PutElem(m_dwCancelID, (DaAsynchronousThread *)this);
        _OPC_CHECK_HR(hr);

        unsigned uCreationFlags = 0;
        unsigned uThreadAddr;

        m_hThread = (HANDLE)_beginthreadex(
            NULL,                // No thread security attributes
            0,                   // Default stack size  
                                 // Pointer to thread function 
            DataCallbackReadThreadHandler,
            this,                // Pass class to new thread 
            uCreationFlags,      // Into ready state 
            &uThreadAddr);      // Address of the thread


        if (m_hThread == 0) {                     // Cannot create the thread
                                                  // Remove from list                        
            m_pParent->m_oaAsyncThread.PutElem(m_dwCancelID, NULL);
            throw (HRESULT)HRESULT_FROM_WIN32(GetLastError());
        }
    }
    catch (HRESULT hrEx) {
        m_ppDItems = NULL;                     // Prevents cleanup of arrays with memory allocated
        m_pItemStates = NULL;                     // outside of the class and passed via parameter list
        m_pdwMaxAge = NULL;                     // if the function fails.
        m_pfPhyval = NULL;                     // Cleanup of arrays with internal allocated memory
        hr = hrEx;                                // is done in the destructor
    }

    if (fLocked) {
        LeaveCriticalSection(&m_pParent->m_AsyncThreadsCritSec);
    }
    return hr;
}



//=========================================================================
// Starts an asynchronous Write transaction.
//=========================================================================
HRESULT DataCallbackThread::CreateCustomWriteVQT(
    DWORD          dwCount,          // Number of items to handle
    DWORD          dwTransactionID,  // Transaction ID
    DaDeviceItem**  ppDItems,         // Array with device item ptrs  
    OPCITEMSTATE*  pItemStates,      // Client handles + requested data types
    OPCITEMVQT*    pItemVQTs,        // The Values, Qualities and TimeStamps to write
    BOOL*          pfPhyval,         // Array which marks items added with their physical value
    // [out]
    DWORD*         pdwCancelID)
{
    _ASSERTE(dwCount);                         // Items must be specified
    _ASSERTE(ppDItems);
    _ASSERTE(pItemStates);
    _ASSERTE(pItemVQTs);
    _ASSERTE(pdwCancelID);
    // Note : pfPhyval may be NULL

    // Initialze all arrays passed via the registered callback to the client
    try {
        m_fxaHandles.Init(dwCount, &m_phClientItems);
        m_fxaErrors.Init(dwCount, &m_pErrors);
    }
    catch (HRESULT hresEx) {
        return hresEx;                            // Cleanup is done in the destructor
    }

    m_dwCount = dwCount;              // Number of items
    m_dwTransactionID = dwTransactionID;      // Number of items
    m_ppDItems = ppDItems;             // Device Items to be read
    m_pItemStates = pItemStates;          // Client handles + requested data types
    m_pfPhyval = pfPhyval;             // Array which marks items added with their physical value
    m_pVQTsToWrite = pItemVQTs;            // Values, Qualities and TimeStamps to write

    unsigned uCreationFlags = 0;
    unsigned uThreadAddr;

    EnterCriticalSection(&m_pParent->m_AsyncThreadsCritSec);
    // Lock the list before create the thread.
    // This way remove requests are prevented
    // before the thread was added to the list.

    m_dwCancelID = *pdwCancelID = m_pParent->m_oaAsyncThread.New();
    // Array index also used as cancel ID
    HRESULT hr = m_pParent->m_oaAsyncThread.PutElem(m_dwCancelID, (DaAsynchronousThread *)this);
    if (FAILED(hr)) {
        LeaveCriticalSection(&m_pParent->m_AsyncThreadsCritSec);
        return hr;
    }

    m_hThread = (HANDLE)_beginthreadex(
        NULL,                // No thread security attributes
        0,                   // Default stack size  
                             // Pointer to thread function 
        DataCallbackWriteThreadHandler,
        this,                // Pass class to new thread 
        uCreationFlags,      // Into ready state 
        &uThreadAddr);      // Address of the thread


    if (m_hThread == 0) {                        // Cannot create the thread
                                                 // Reove from list                        
        m_pParent->m_oaAsyncThread.PutElem(m_dwCancelID, NULL);
        m_ppDItems = NULL;                    // Prevents cleanup in destructor
        m_pItemStates = NULL;
        m_pfPhyval = NULL;
        m_pVQTsToWrite = NULL;

        hr = HRESULT_FROM_WIN32(GetLastError());
    }
    LeaveCriticalSection(&m_pParent->m_AsyncThreadsCritSec);
    return hr;
}



//=========================================================================
// Sets the values in the arrays with the values from pItemStates. A
// data value is copied only if the severity code of the corresponding
// error code is of type 'Success'.
//=========================================================================
void DataCallbackThread::SetCallbackResultsFromItemStates(
    // IN
    DWORD          dwCount,
    OPCITEMSTATE*  pItemStates,

    // IN / OUT
    HRESULT*       phrMasterError,
    HRESULT*       errors,

    // OUT
    HRESULT*       phrMasterQuality,
    OPCHANDLE*     phClientItems,
    VARIANT*       pvValues,
    WORD*          pwQualities,
    FILETIME*      pftTimeStamps)
{
    HRESULT  hres;
    DWORD    i;

    *phrMasterQuality = S_OK;
    for (i = 0; i < dwCount; i++, pItemStates++) {

        phClientItems[i] = pItemStates->hClient;
        pwQualities[i] = pItemStates->wQuality;
        pftTimeStamps[i] = pItemStates->ftTimeStamp;

        if (SUCCEEDED(errors[i])) {            // Copy value only if source pItemStates->vDataValue is valid
            hres = VariantCopy(&pvValues[i], &pItemStates->vDataValue);
            if (FAILED(hres)) {
                errors[i] = hres;
                *phrMasterError = S_FALSE;          // Not all item errors are S_OK
            }
        }
        else {
            *phrMasterError = S_FALSE;             // Not all item errors are S_OK
        }
        if ((pItemStates->wQuality & OPC_QUALITY_MASK) != OPC_QUALITY_GOOD) {
            *phrMasterQuality = S_FALSE;           // Not all Qualities are good
        }
    }
}



//-----------------------------------------------------------------------
// IMPLEMENTATION
//-----------------------------------------------------------------------

//=========================================================================
// Kill this instance and the thread running over it
//=========================================================================
void DataCallbackThread::DoKill(void)
{
    CloseHandle(m_hThread);                    // Thread handle must be closed explicitly
                                                 // if _endthreadex() is used.
    delete this;                                 // Kill the class object
    _endthreadex(0);                             // Kill the thread
}



//=========================================================================
// Checks if the cancel flag is set.
// If the flag is set then the instance and the thread running over it
// are killed -> the function never returns.
//
// Parameters:
//    fPreventCancel    If this flag is TRUE then the transaction can be
//                      no longer canceled after executed this function.
//                      (The thread will be removed from the array
//                      of callback threads).
//=========================================================================
void DataCallbackThread::CheckCancelRequestAndKillFlag(BOOL fPreventCancel /* = FALSE */)
{
    EnterCriticalSection(&m_pParent->m_AsyncThreadsCritSec);
    // Prevents installation of a Cancel Thread
// Check if the group has the kill flag set
    if (m_pParent->Killed()) {
        // Remove the thread from the array
        // of data callback threads
        m_pParent->m_oaAsyncThread.PutElem(m_dwCancelID, NULL);
        LeaveCriticalSection(&m_pParent->m_AsyncThreadsCritSec);
        DoKill();                                 // Kill the thread and and the class
        _ASSERTE(0);                            // Should never get here
    }
    // Check if there is a cancel request
    if (m_pCancelThread) {
        m_pParent->m_oaAsyncThread.PutElem(m_dwCancelID, NULL);
        LeaveCriticalSection(&m_pParent->m_AsyncThreadsCritSec);
        // There is a cancel request for the current transaction
        m_pCancelThread->InvokeCancelCallback();  // Invoke the registered cancel complete callback

                                                  // Wait until the cancel thread has terminated
        WaitForSingleObject(m_pCancelThread->m_hThread, INFINITE);

        DoKill();                                 // Kill the thread and and the class
        _ASSERTE(0);                            // Should never get here
    }
    else if (fPreventCancel) {
        // After this it is no longer possible to cancel the transaction.
        m_pParent->m_oaAsyncThread.PutElem(m_dwCancelID, NULL);
    }
    LeaveCriticalSection(&m_pParent->m_AsyncThreadsCritSec);
}


//-----------------------------------------------------------------------
// HANDLER THREADS
//-----------------------------------------------------------------------

//=========================================================================
// Thread handling asynchronous Read and Refresh transactions
// ----------------------------------------------------------
//    The thread handles a single request and then terminates.
//=========================================================================
unsigned __stdcall DataCallbackReadThreadHandler(void* pCreator)
{
    DataCallbackThread* pThrd = static_cast<DataCallbackThread *>(pCreator);
    _ASSERTE(pThrd);

    // Read the data from the device or cache, returns only S_OK or S_FALSE.
    HRESULT hrMasterError = S_OK;
    HRESULT hrMasterQuality = S_OK;

    if (pThrd->m_pdwMaxAge) {                    // Initiated by IOPCAsyncIO3::ReadMaxAge()
                                                 // or IOPCAsyncIO3::RefreshMaxAge()
        _ASSERTE(pThrd->m_ppTmpBufferForItemsToReadFromDevice);
        DWORD i;
        // Initialize the Requested Data Types
        for (i = 0; i < pThrd->m_dwCount; i++) {
            V_VT(&pThrd->m_pvValues[i]) = V_VT(&pThrd->m_pItemStates[i].vDataValue);
        }

        hrMasterError = pThrd->m_pParent->m_pServer->InternalReadMaxAge(
            &pThrd->m_ftNow,
            pThrd->m_dwCount,
            pThrd->m_ppDItems,
            pThrd->m_ppTmpBufferForItemsToReadFromDevice,
            pThrd->m_pdwMaxAge,
            pThrd->m_pvValues,
            pThrd->m_pwQualities,
            pThrd->m_pftTimeStamps,
            pThrd->m_pErrors,
            pThrd->m_pfPhyval);

        pThrd->CheckCancelRequestAndKillFlag();   // Maybe never returns from this function      

                                                  // Initialize the Client Handle array and the Master Quality
        for (i = 0; i < pThrd->m_dwCount; i++) {

            pThrd->m_phClientItems[i] = pThrd->m_pItemStates[i].hClient;

            if ((pThrd->m_pwQualities[i] & OPC_QUALITY_MASK) != OPC_QUALITY_GOOD) {
                hrMasterQuality = S_FALSE;          // Not all Qualities are good
            }
        }

    }
    else {                                       // Initiated by IOPCAsyncIO2::Read()
                                                 // or IOPCAsyncIO2::Refresh2()
        hrMasterError = pThrd->m_pParent->InternalRead(
            pThrd->m_dwSource,
            pThrd->m_dwCount,
            pThrd->m_ppDItems,
            pThrd->m_pItemStates,
            pThrd->m_pErrors,
            pThrd->m_pfPhyval);

        pThrd->CheckCancelRequestAndKillFlag();   // Maybe never returns from this function

        // Initialize the result arrays with the values from the OPC item states
        pThrd->SetCallbackResultsFromItemStates(
            // IN
            pThrd->m_dwCount,
            pThrd->m_pItemStates,

            // IN / OUT
            &hrMasterError,
            pThrd->m_pErrors,

            // OUT
            &hrMasterQuality,
            pThrd->m_phClientItems,
            pThrd->m_pvValues,
            pThrd->m_pwQualities,
            pThrd->m_pftTimeStamps);
    }

    pThrd->CheckCancelRequestAndKillFlag(TRUE);// The last check of the cancel flag.
                                                 // After that it is no longer possible to cancel the transaction.
                                                 // Maybe never returns from this function.

    // Get the COM object
    CComObject<DaGroup>* pCOMGroup = NULL;

    pThrd->m_pGServer->CriticalSectionCOMGroupList.BeginReading();     // Lock reading COM list

    HRESULT hres = pThrd->m_pGServer->m_COMGroupList.GetElem(
        pThrd->m_pParent->m_hServerGroupHandle,
        &pCOMGroup);

    if (SUCCEEDED(hres)) {

        // Invoke the callback
        pCOMGroup->Lock();                        // Lock the connection point list

        IOPCDataCallback* pCallback;

        hres = pCOMGroup->GetCallbackInterface(&pCallback);
        if (SUCCEEDED(hres)) {

            pThrd->m_fReadTransaction ?

                pCallback->OnReadComplete(
                    pThrd->m_dwTransactionID,        // Transaction ID
                                                     // Client Group Handle
                    pThrd->m_pParent->m_hClientGroupHandle,
                    hrMasterQuality,                 // Master Quality
                    hrMasterError,                   // Master Error
                    pThrd->m_dwCount,                // Number of elements in the following arrays
                    pThrd->m_phClientItems,          // List of client item handles
                    pThrd->m_pvValues,               // List of Variant Values
                    pThrd->m_pwQualities,            // List of Quality Values
                    pThrd->m_pftTimeStamps,          // List of Time Stamps
                    pThrd->m_pErrors)               // List of Errors
                :
                pCallback->OnDataChange(
                    pThrd->m_dwTransactionID,        // Transaction ID
                                                     // Client Group Handle
                    pThrd->m_pParent->m_hClientGroupHandle,
                    hrMasterQuality,                 // Master Quality
                    hrMasterError,                   // Master Error
                    pThrd->m_dwCount,                // Number of elements in the following arrays
                    pThrd->m_phClientItems,          // List of client item handles
                    pThrd->m_pvValues,               // List of Variant Values
                    pThrd->m_pwQualities,            // List of Quality Values
                    pThrd->m_pftTimeStamps,          // List of Time Stamps
                    pThrd->m_pErrors);              // List of Errors

            pCallback->Release();                  // All is done with this interface

            DaGenericGroup* pGGroup;
            if (SUCCEEDED(pCOMGroup->GetGenericGroup(&pGGroup))) {
                pGGroup->ResetKeepAliveCounter();
                pCOMGroup->ReleaseGenericGroup();
            }
        }

        pCOMGroup->Unlock();                      // Unlock the connection point list
    }

    pThrd->m_pGServer->CriticalSectionCOMGroupList.EndReading();   // Unlock reading COM list

    pThrd->DoKill();                             // Kill the thread and and the class
    return 0;                                    // Should never get here
}



//=========================================================================
// Thread handling asynchronous Write transactions
// -----------------------------------------------
//    The thread handles a single request and then terminates.
//=========================================================================
unsigned __stdcall DataCallbackWriteThreadHandler(void* pCreator)
{
    DataCallbackThread* pThrd = static_cast<DataCallbackThread *>(pCreator);
    _ASSERTE(pThrd);

    pThrd->CheckCancelRequestAndKillFlag(TRUE);// Check of the cancel flag.
                                                 // After that it is no longer possible to cancel the transaction.
                                                 // Maybe never returns from this function.

    // Write the data to the device
    HRESULT hrMasterError = pThrd->m_pParent->InternalWriteVQT(
        pThrd->m_dwCount,
        pThrd->m_ppDItems,
        pThrd->m_pVQTsToWrite,
        pThrd->m_pErrors,
        pThrd->m_pfPhyval
    );

    // Get the COM object
    CComObject<DaGroup>* pCOMGroup = NULL;

    pThrd->m_pGServer->CriticalSectionCOMGroupList.BeginReading();     // Lock reading COM list

    HRESULT hres;
    hres = pThrd->m_pGServer->m_COMGroupList.GetElem(
        pThrd->m_pParent->m_hServerGroupHandle,
        &pCOMGroup);

    if (SUCCEEDED(hres)) {

        // Invoke the callback
        pCOMGroup->Lock();                        // Lock the connection point list

        IOPCDataCallback* pCallback;

        hres = pCOMGroup->GetCallbackInterface(&pCallback);
        if (SUCCEEDED(hres)) {

            // Initialize the result arrays with the values from the OPC item states
            for (DWORD i = 0; i < pThrd->m_dwCount; i++) {
                pThrd->m_phClientItems[i] = pThrd->m_pItemStates[i].hClient;
                // Initialize because olny type is set
                VariantInit(&pThrd->m_pItemStates[i].vDataValue);
            }

            pCallback->OnWriteComplete(
                pThrd->m_dwTransactionID,        // Transaction ID
                                                 // Client Group Handle
                pThrd->m_pParent->m_hClientGroupHandle,
                hrMasterError,                   // Master Error, S_OK or S_FALSE from InternalWriteVQT
                pThrd->m_dwCount,                // Number of elements in the following arrays
                pThrd->m_phClientItems,          // List of client item handles
                pThrd->m_pErrors);              // List of Errors

            pCallback->Release();                  // All is done with this interface

            DaGenericGroup* pGGroup;
            if (SUCCEEDED(pCOMGroup->GetGenericGroup(&pGGroup))) {
                pGGroup->ResetKeepAliveCounter();
                pCOMGroup->ReleaseGenericGroup();
            }
        }

        pCOMGroup->Unlock();                      // Unlock the connection point list
    }

    pThrd->m_pGServer->CriticalSectionCOMGroupList.EndReading();   // Unlock reading COM list

    pThrd->DoKill();                             // Kill the thread and and the class
    return 0;                                    // Should never get here
}



///////////////////////////////////////////////////////////////////////////
// DataCallbackThread::CCancelThread
///////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////
//=========================================================================
// Constructor
//=========================================================================
DataCallbackThread::CCancelThread::CCancelThread()
{
    m_pThreadToCancel = NULL;                // The owner
    m_hOperationCanceledEvent = NULL;
}



//=========================================================================
// Destructor
//=========================================================================
DataCallbackThread::CCancelThread::~CCancelThread()
{
    if (m_hOperationCanceledEvent) {
        CloseHandle(m_hOperationCanceledEvent);
    }
}



//=========================================================================
// Initialization
// --------------
//    Creates the Cancel Thread.
//=========================================================================
HRESULT DataCallbackThread::CCancelThread::Create(DataCallbackThread* pThreadToCancel)
{
    _ASSERTE(pThreadToCancel);

    m_pThreadToCancel = pThreadToCancel;

    m_hOperationCanceledEvent = CreateEvent(
        NULL,                // Handle cannot be inherited
        TRUE,                // Manually reset requested
        FALSE,               // Initial state is nonsignaled
        NULL);              // No event object name 

    if (m_hOperationCanceledEvent == NULL) {
        return E_FAIL;                            // Cannot create event
    }

    unsigned uCreationFlags = 0;
    unsigned uThreadAddr;

    m_hThread = (HANDLE)_beginthreadex(
        NULL,                // No thread security attributes
        0,                   // Default stack size  
        CancelThreadHandler, // Pointer to thread function 
        this,                // Pass class to new thread
        uCreationFlags,      // Into ready state 
        &uThreadAddr);      // Address of the thread

    if (m_hThread == 0) {                        // Cannot create the thread
        return HRESULT_FROM_WIN32(GetLastError());
    }
    return S_OK;
}



//=========================================================================
// Cancel Thread Handler
// ---------------------
//    Waits until the thread of the transaction has terminated or
//    successfully canceled.
//    If the transaction successfully canceled then the registered 
//    cancel callback will be invoked.
//    After that the thread terminates and deletes the class.
//=========================================================================
unsigned __stdcall CancelThreadHandler(void* pCreator)
{
    DataCallbackThread::CCancelThread* pThrd = static_cast<DataCallbackThread::CCancelThread *>(pCreator);
    _ASSERTE(pThrd);

    const DataCallbackThread* pThreadToCancel = pThrd->m_pThreadToCancel;
    _ASSERTE(pThreadToCancel);

    // Waits until the thread of the transaction has successfully canceled.
    WaitForSingleObject(pThrd->m_hOperationCanceledEvent, INFINITE);


    // Invoke the cancel complete callback
    LOGFMTI("Transaction successfully canceled");

    // Get the COM object
    CComObject<DaGroup>* pCOMGroup = NULL;

    // Lock reading COM list
    pThreadToCancel->m_pGServer->CriticalSectionCOMGroupList.BeginReading();

    HRESULT hres = pThreadToCancel->m_pGServer->m_COMGroupList.GetElem(
        pThreadToCancel->m_pParent->m_hServerGroupHandle,
        &pCOMGroup);

    if (SUCCEEDED(hres)) {
        // Invoke the callback
        pCOMGroup->Lock();               // Lock the connection point list

        IOPCDataCallback* pCallback;

        hres = pCOMGroup->GetCallbackInterface(&pCallback);
        if (SUCCEEDED(hres)) {

            pCallback->OnCancelComplete(
                // Transaction ID
                pThreadToCancel->m_dwTransactionID,
                // Group Handle
                pThreadToCancel->m_pParent->m_hClientGroupHandle);

            pCallback->Release();         // All is done with this interface
        }

        pCOMGroup->Unlock();             // Unlock the connection point list
    }
    // Unlock reading COM list
    pThreadToCancel->m_pGServer->CriticalSectionCOMGroupList.EndReading();

    // Kill the thread and and the class
    CloseHandle(pThrd->m_hThread);             // Thread handle must be closed explicitly
                                                 // if _endthreadex() is used.
    delete pThrd;                                // Kill the class object
    _endthreadex(0);                             // Kill the thread
    return 0;                                    // Should never get here
}
//DOM-IGNORE-END
