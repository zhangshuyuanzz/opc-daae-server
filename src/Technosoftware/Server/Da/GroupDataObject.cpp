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
#include "UtilityFuncs.h"
#include "VariantPack.h"
#include "DaGenericGroup.h"

//=================================================================================
// DaGenericGroup::SendDataStream
// -----------------------------
// Builds the stream for OnDataChange and async read/write/refresh operations.
//
// Create the Stream which consists of:
//  GroupHeader
//  ItemHeaders
//  Variants and Related Data.
// returns;
//    S_OK              succeeded
//    S_FALSE           stream sent but with error code in hrStatus
//                      (sent only the group header)
//    E_xxx             appropriate error code
//=================================================================================
HRESULT  DaGenericGroup::SendDataStream(
    BOOL           WithTime,
    DWORD          NumItems,
    OPCITEMSTATE  *pItemValues,
    DWORD          tid)
{
    DWORD             HdrSize, TotalSize;
    DWORD             n;
    char             *gp;
    HGLOBAL           gh;
    OPCGROUPHEADER   *GrpPtr;
    OPCITEMHEADER1   *ItemHdr1;   // With Time
    OPCITEMHEADER2   *ItemHdr2;   // Without Time
    char             *DataPtr;
    DWORD             vsize;

    HRESULT  hrStatus = S_OK;
    HRESULT  res = S_OK;

    if (NumItems == 0) {
        return S_OK;
    }

    // Allocate the Data Stream
    if (WithTime) {
        HdrSize = sizeof(OPCGROUPHEADER) + NumItems * sizeof(OPCITEMHEADER1);   // with time
    }
    else {
        HdrSize = sizeof(OPCGROUPHEADER) + NumItems * sizeof(OPCITEMHEADER2);   // without time
    }

    TotalSize = HdrSize;
    for (n = 0; n < NumItems; n++) {
        long lsize = SizePackedVariant(&pItemValues[n].vDataValue);           // to calculate VARIANT size
        if (lsize == -1) {                                                      // Version 1.3, AM
            hrStatus = E_FAIL;
            TotalSize = HdrSize;                                                 // Send only the Group Header
            NumItems = 0;
            break;
        }
        TotalSize += lsize;
    }

    gh = GlobalAlloc(GMEM_FIXED + GMEM_SHARE, TotalSize);
    if (gh == NULL) {                                                          // If there is not enough memory try to allocate
        gh = GlobalAlloc(GMEM_FIXED + GMEM_SHARE, HdrSize);                   // memory only for the group header.
        if (gh == NULL) {
            return E_OUTOFMEMORY;                                                // Also not enough memory for group header.
        }
        hrStatus = E_OUTOFMEMORY;                                              // Send only the Group Header
        TotalSize = HdrSize;
        NumItems = 0;
    }
    gp = (char*)GlobalLock(gh);

    GrpPtr = (OPCGROUPHEADER *)gp;
    ItemHdr1 = (OPCITEMHEADER1 *)(gp + sizeof(OPCGROUPHEADER));   //With Time
    ItemHdr2 = (OPCITEMHEADER2 *)ItemHdr1;                        //Without Time
    DataPtr = gp + HdrSize;

    // Fill in the Group header
    GrpPtr->dwSize = TotalSize;
    GrpPtr->dwItemCount = NumItems;
    GrpPtr->hClientGroup = m_hClientGroupHandle;
    GrpPtr->hrStatus = hrStatus;
    GrpPtr->dwTransactionID = tid;


    if (SUCCEEDED(hrStatus)) {
        // Fill in the Item Information
        for (n = 0; n < NumItems; n++) {
            // Get ItemHeader info and ItemValue
            if (WithTime) {
                ItemHdr1->hClient = pItemValues[n].hClient;
                ItemHdr1->wQuality = pItemValues[n].wQuality;
                ItemHdr1->ftTimeStampItem = pItemValues[n].ftTimeStamp;
                ItemHdr1->dwValueOffset = (DWORD)(DataPtr - gp);
                ItemHdr1->wReserved = 0;

            }
            else {  // without time
                ItemHdr2->hClient = pItemValues[n].hClient;
                ItemHdr2->wQuality = pItemValues[n].wQuality;
                ItemHdr2->dwValueOffset = (DWORD)(DataPtr - gp);
                ItemHdr2->wReserved = 0;
            }

            if ((pItemValues[n].wQuality & OPC_QUALITY_MASK) != OPC_QUALITY_GOOD) {
                GrpPtr->hrStatus = S_FALSE;            // One or more items has a quality status of BAD or UNCERTAIN.
            }

            vsize = CopyPackVariant(DataPtr, &pItemValues[n].vDataValue);
            if (vsize == -1) {                       // Version 1.3, AM
                GrpPtr->dwSize = HdrSize;
                GrpPtr->dwItemCount = 0;
                GrpPtr->hrStatus = E_FAIL;
                res = S_FALSE;     // stream sent, but with errror code
                break;
            }
            DataPtr += vsize;
            ItemHdr1++;
            ItemHdr2++;
        } //for
    } // succeeded
    else {
        res = S_FALSE;
    }

    // Invoke the callback
    STGMEDIUM stm;
    FORMATETC fmt;
    stm.tymed = TYMED_HGLOBAL;
    stm.hGlobal = gh;
    stm.pUnkForRelease = NULL;

    fmt.ptd = NULL;
    fmt.dwAspect = DVASPECT_CONTENT;
    fmt.lindex = -1;
    fmt.tymed = TYMED_HGLOBAL;


    EnterCriticalSection(&m_CallbackCritSec);

    if (WithTime) {                  // with time
        fmt.cfFormat = m_StreamDataTime;

        // make sure client did not unadvise in the mean time
        if (m_DataTimeCallback) {
            // do the callback

            IAdviseSink* pSink;

            res = core_generic_main.m_pGIT->GetInterfaceFromGlobal((DWORD)m_DataTimeCallback, IID_IAdviseSink, (LPVOID*)&pSink);
            if (SUCCEEDED(res)) {
                pSink->OnDataChange(&fmt, &stm);
                pSink->Release();
            }
        }
        else {
            res = E_FAIL;
        }
    }
    else {                           // without time
        fmt.cfFormat = m_StreamData;

        // make sure client did not unadvise in the mean time
        if (m_DataCallback) {
            // do the callback

            IAdviseSink* pSink;

            res = core_generic_main.m_pGIT->GetInterfaceFromGlobal(m_DataCallback, IID_IAdviseSink, (LPVOID*)&pSink);
            if (SUCCEEDED(res)) {
                pSink->OnDataChange(&fmt, &stm);
                pSink->Release();
            }
        }
        else {
            res = E_FAIL;
        }
    }
    LeaveCriticalSection(&m_CallbackCritSec);

    GlobalUnlock(gh);
    GlobalFree(gh);                // free the buffer area

    return res;
}



//=================================================================================
// DaGenericGroup::SendWriteStream
// ------------------------------
// This function is used to build the stream for OnWriteComplete
//
// Create the Stream which consists of:
//  GroupHeader
//  ItemHeaders
//=================================================================================
HRESULT   DaGenericGroup::SendWriteStream(
    DWORD                NumItems,
    OPCITEMHEADERWRITE * pItemInfo,
    HRESULT              Status,
    DWORD                tid)
{
    DWORD    i, HdrSize, TotalSize;
    char   * gp;
    HGLOBAL  gh;
    OPCGROUPHEADERWRITE * GrpPtr;
    OPCITEMHEADERWRITE  * ItemHdr1;

    HRESULT res;

    if (NumItems == 0) {
        return S_OK;
    }

    // Allocate the Data Stream
    HdrSize = sizeof(OPCGROUPHEADERWRITE) + NumItems * sizeof(OPCITEMHEADERWRITE);
    TotalSize = HdrSize;

    if ((gh = GlobalAlloc(GMEM_FIXED, TotalSize)) == NULL) {
        return E_OUTOFMEMORY;          // no memory, ignore the request
    }
    gp = (char*)GlobalLock(gh);
    GrpPtr = (OPCGROUPHEADERWRITE*)gp;
    ItemHdr1 = (OPCITEMHEADERWRITE*)(gp + sizeof(OPCGROUPHEADERWRITE));

    // Fill in the Group header
    GrpPtr->dwItemCount = NumItems;
    GrpPtr->hClientGroup = m_hClientGroupHandle;
    GrpPtr->hrStatus = Status;
    GrpPtr->dwTransactionID = tid;

    // Fill in the Item Information
    for (i = 0; i < NumItems; i++) {
        ItemHdr1[i] = pItemInfo[i];
    }

    // And invoke the callback
    STGMEDIUM stm;
    FORMATETC fe;
    fe.cfFormat = m_StreamWrite;
    fe.ptd = NULL;
    fe.dwAspect = DVASPECT_CONTENT;
    fe.lindex = -1;
    fe.tymed = TYMED_HGLOBAL;
    stm.tymed = TYMED_HGLOBAL;
    stm.hGlobal = gh;
    stm.pUnkForRelease = NULL;

    res = S_OK;

    EnterCriticalSection(&m_CallbackCritSec);
    // make sure client did not unadvise 
    if (m_WriteCallback) {
        // do the callback

        IAdviseSink* pSink;

        res = core_generic_main.m_pGIT->GetInterfaceFromGlobal(m_WriteCallback, IID_IAdviseSink, (LPVOID*)&pSink);
        if (SUCCEEDED(res)) {
            pSink->OnDataChange(&fe, &stm);
            pSink->Release();
        }
    }
    else {
        res = E_FAIL;
    }
    LeaveCriticalSection(&m_CallbackCritSec);

    GlobalUnlock(gh);
    GlobalFree(gh);

    return res;
}


//=================================================================================
//=================================================================================
HRESULT DaGenericGroup::SendDataStreamDisp(
    BOOL           WithTime,
    long           NumItems,
    OPCITEMSTATE  *pItemValues,
    DWORD          tid)
{

    unsigned int            firstErrArg;
    DISPPARAMS              *dp;

    VARIANTARG              *args;
    long                    ch;
    VARIANT                 valvt;
    long                    qual;
    FILETIME                tim;
    DATE                    datum;
    long                    i;
    HRESULT                 res;
    OPCITEMSTATE            *pIS;


    if (NumItems == 0) {
        return S_OK;
    }


    dp = (DISPPARAMS *)pIMalloc->Alloc(sizeof(DISPPARAMS));
    if (dp == NULL) {
        res = E_OUTOFMEMORY;
        goto SendDataDispExit2;
    }
    dp->rgdispidNamedArgs = NULL;
    dp->cNamedArgs = 0;
    dp->cArgs = 7;

    args = (VARIANTARG *)pIMalloc->Alloc(sizeof(VARIANTARG) * dp->cArgs);
    if (args == NULL) {
        res = E_OUTOFMEMORY;
        goto SendDataDispExit3;
    }
    for (i = 0; i < (long)dp->cArgs; i++) {
        VariantInit(&args[i]);
    }

    // Call Automation procedure (Visual Basic or Delphi)
    //
    //   OnDataChange(   GroupHandle as Long, 
    //                   Count as Long,
    //                   ClientHandles() as Long;
    //                   Values() as Variant,
    //                   Qualities() as Long,
    //                   GroupTime as Date,
    //                   TimeStamps() as Date  )
    //

    // Timestamps of the items -------------------- PARAM 6
    res = VariantInitArrayOf(&(args[0]),
        NumItems,
        VT_DATE);
    if (FAILED(res)) {
        goto SendDataDispExit4;
    }

    // Number of items sent ----------------------- PARAM 5
    VariantInit(&(args[1]));
    CoFileTimeNow(&tim);
    FileTimeToDATE(&tim, datum);
    V_DATE(&args[1]) = datum;

    // Qualities of the items --------------------- PARAM 4
    res = VariantInitArrayOf(&(args[2]),
        NumItems,
        VT_I4);
    if (FAILED(res)) {
        goto SendDataDispExit5;
    }

    // Values of the items ------------------------ PARAM 3
    res = VariantInitArrayOf(&(args[3]),
        NumItems,
        VT_VARIANT);
    if (FAILED(res)) {
        goto SendDataDispExit6;
    }

    // Client Handles of the items ---------------- PARAM 2
    res = VariantInitArrayOf(&(args[4]),
        NumItems,
        VT_I4);
    if (FAILED(res)) {
        goto SendDataDispExit7;
    }

    // Number of items sent ----------------------- PARAM 1
    VariantInit(&args[5]);
    V_VT(&args[5]) = VT_I4;
    V_I4(&args[5]) = NumItems;

    // Group Handle ------------------------------- PARAM 0
    VariantInit(&args[6]);
    V_VT(&args[6]) = VT_I4;
    V_I4(&args[6]) = m_hClientGroupHandle;

    dp->rgvarg = args;

    VariantInit(&valvt);
    i = 0;
    while (i < NumItems) {
        pIS = &(pItemValues[i]);

        FileTimeToDATE(&(pIS->ftTimeStamp), datum);
        res = SafeArrayPutElement(V_ARRAY(&args[0]), &i, &datum);
        if (FAILED(res)) {
            goto SendDataDispExit9;
        }

        qual = pIS->wQuality;
        res = SafeArrayPutElement(V_ARRAY(&args[2]), &i, &qual);
        if (FAILED(res)) {
            goto SendDataDispExit9;
        }

        res = VariantCopy(&valvt, &(pIS->vDataValue));
        if (FAILED(res)) {
            goto SendDataDispExit9;
        }
        res = SafeArrayPutElement(V_ARRAY(&args[3]), &i, &valvt);
        if (FAILED(res)) {
            goto SendDataDispExit9;
        }

        ch = pIS->hClient;
        res = SafeArrayPutElement(V_ARRAY(&args[4]), &i, &ch);
        if (FAILED(res)) {
            goto SendDataDispExit9;
        }


        i++;
    }
    VariantClear(&valvt);

    //                                              CALL

    EnterCriticalSection(&m_CallbackCritSec);
    // make sure client did not unadvise 
    if (m_DataTimeCallbackDisp) {
        // do the callback

        IDispatch* pDisp;

        res = core_generic_main.m_pGIT->GetInterfaceFromGlobal(m_DataTimeCallbackDisp, IID_IDispatch, (LPVOID*)&pDisp);
        if (SUCCEEDED(res)) {
            res = pDisp->Invoke(m_DataTimeCallbackMethodID,
                IID_NULL,
                m_dwLCID,
                DISPATCH_METHOD,
                dp,
                NULL,
                NULL,
                &firstErrArg
            );
            pDisp->Release();
        }

        if (FAILED(res)) {
            LeaveCriticalSection(&m_CallbackCritSec);
            goto SendDataDispExit8;
        }
    }
    else {
        res = E_FAIL;
    }
    LeaveCriticalSection(&m_CallbackCritSec);

    // free the params (arrays and not)
    for (i = 0; i < (long)dp->cArgs; i++) {
        VariantClear(&args[i]);
    }
    pIMalloc->Free(args);
    pIMalloc->Free(dp);

    return res;


SendDataDispExit9:
    VariantClear(&valvt);

SendDataDispExit8:
    VariantUninitArrayOf(&(args[4]));

SendDataDispExit7:
    VariantUninitArrayOf(&(args[3]));

SendDataDispExit6:
    VariantUninitArrayOf(&(args[2]));

SendDataDispExit5:
    VariantUninitArrayOf(&(args[0]));

SendDataDispExit4:
    pIMalloc->Free(args);

SendDataDispExit3:
    pIMalloc->Free(dp);

SendDataDispExit2:
    return res;


}




//=================================================================================
//=================================================================================
HRESULT   DaGenericGroup::SendWriteStreamDisp(
    long                 NumItems,
    long                 *pClientHandles,
    long                 *pError,
    HRESULT              Status,
    DWORD                tid)    // Transaction ID
{

    unsigned int            firstErrArg;
    DISPPARAMS              *dp;

    VARIANTARG              *args;
    long                    ch;
    long                    i;
    HRESULT                 res;


    if (NumItems == 0) {
        return S_OK;
    }

    dp = (DISPPARAMS *)pIMalloc->Alloc(sizeof(DISPPARAMS));
    if (dp == NULL) {
        res = E_OUTOFMEMORY;
        goto SendWriteDispExit2;
    }
    dp->rgdispidNamedArgs = NULL;
    dp->cNamedArgs = 0;
    dp->cArgs = 4;

    args = (VARIANTARG *)pIMalloc->Alloc(sizeof(VARIANTARG) * dp->cArgs);
    if (args == NULL) {
        res = E_OUTOFMEMORY;
        goto SendWriteDispExit3;
    }
    for (i = 0; i < (long)dp->cArgs; i++) {
        VariantInit(&args[i]);
    }


    // Call Automation procedure (Visual Basic or Delphi)
    //
    //   OnAsyncWriteComplete( GroupHandle as Long, 
    //                         Count as Long,
    //                         ClientHandles() as Long;
    //                         ItemResults() as Long )
    //

    // Item Results ------------------------------- PARAM 3
    res = VariantInitArrayOf(&(args[0]),
        NumItems,
        VT_I4);
    if (FAILED(res)) {
        goto SendWriteDispExit4;
    }

    // Client Handles ----------------------------- PARAM 2
    res = VariantInitArrayOf(&(args[1]),
        NumItems,
        VT_I4);
    if (FAILED(res)) {
        goto SendWriteDispExit5;
    }

    // Number of items sent ----------------------- PARAM 1
    VariantInit(&args[2]);
    V_VT(&args[2]) = VT_I4;
    V_I4(&args[2]) = NumItems;

    // Group Handle ------------------------------- PARAM 0
    VariantInit(&args[3]);
    V_VT(&args[3]) = VT_I4;
    V_I4(&args[3]) = m_hClientGroupHandle;

    dp->rgvarg = args;

    i = 0;
    while (i < NumItems) {

        ch = pError[i];
        res = SafeArrayPutElement(V_ARRAY(&args[0]), &i, &ch);
        if (FAILED(res)) {
            goto SendWriteDispExit6;
        }

        ch = pClientHandles[i];
        res = SafeArrayPutElement(V_ARRAY(&args[1]), &i, &ch);
        if (FAILED(res)) {
            goto SendWriteDispExit6;
        }

        i++;
    }

    //                                              CALL

    EnterCriticalSection(&m_CallbackCritSec);
    // make sure client did not unadvise 
    if (m_WriteCallbackDisp) {
        // do the callback

        IDispatch* pDisp;

        res = core_generic_main.m_pGIT->GetInterfaceFromGlobal((DWORD)m_WriteCallbackDisp, IID_IDispatch, (LPVOID*)&pDisp);
        if (SUCCEEDED(res)) {
            res = pDisp->Invoke(m_WriteCallbackMethodID,
                IID_NULL,
                m_dwLCID,
                DISPATCH_METHOD,
                dp,
                NULL,
                NULL,
                &firstErrArg
            );
            pDisp->Release();
        }
        if (FAILED(res)) {
            LeaveCriticalSection(&m_CallbackCritSec);
            goto SendWriteDispExit6;
        }
    }
    else {
        res = E_FAIL;
    }
    LeaveCriticalSection(&m_CallbackCritSec);

    // free the params (arrays and not)
    for (i = 0; i < (long)dp->cArgs; i++) {
        VariantClear(&args[i]);
    }
    pIMalloc->Free(args);
    pIMalloc->Free(dp);

    return res;



SendWriteDispExit6:
    VariantClear(&(args[3]));
    VariantClear(&(args[2]));
    VariantUninitArrayOf(&(args[1]));

SendWriteDispExit5:
    VariantUninitArrayOf(&(args[0]));

SendWriteDispExit4:
    pIMalloc->Free(args);

SendWriteDispExit3:
    pIMalloc->Free(dp);

SendWriteDispExit2:

    return res;
}



//=========================================================================
// Collect and send Data of Update thread to custom client
// The return value may be used by the server to check
// if the system is overloaded ( errors may return
//
//    returns  E_FAIL if group ok but could not send
//             S_OK   if group to kill or group ok and successfully sent
//=========================================================================
HRESULT DaGenericGroup::UpdateToClient(BOOL custom, BOOL WithTime, BOOL DataCallbackOnly)
{
    long           i, TotGroupItems;
    long           TotItemsToRead, TotItemsToTransmit;
    DaDeviceItem   **ppDItems, *pDItem;
    DaGenericItem  **ppGItems, *pGItem;
    HRESULT        *pErr, res;
    OPCITEMSTATE   *pItemStates;
    DWORD          AccessRight;

    // while building arrays don't allow add and delete of items to group
    EnterCriticalSection(&m_ItemsCritSec);

    // number of items in the group
    res = S_OK;
    TotGroupItems = m_oaItems.TotElem();
    if (TotGroupItems == 0) {
        LeaveCriticalSection(&m_ItemsCritSec);
        goto UpdateToClient0;
    }
    res = E_OUTOFMEMORY;
    ppGItems = new DaGenericItem*[TotGroupItems];// create generic item array object
    if (!ppGItems) {
        LeaveCriticalSection(&m_ItemsCritSec);
        goto UpdateToClient0;
    }

    ppDItems = new DaDeviceItem*[TotGroupItems]; // create device item array object
    if (!ppDItems) {
        LeaveCriticalSection(&m_ItemsCritSec);
        goto UpdateToClient1;
    }

    // Initialize the arrays for generic and Device Items
    TotItemsToRead = 0;
    res = m_oaItems.First(&i);
    while (SUCCEEDED(res)) {
        m_oaItems.GetElem(i, &pGItem);
        if (pGItem && pGItem->get_Active()) {     // item must be existent and active

            if (pGItem->AttachDeviceItem(&pDItem) >= 0) {
                // Item to be handled
                // Device item exist and has not set the killed flag
                res = pGItem->get_AccessRights(&AccessRight);
                if (SUCCEEDED(res) &&
                    ((AccessRight & OPC_READABLE) != 0)) {

                    ppGItems[TotItemsToRead] = pGItem;
                    ppDItems[TotItemsToRead] = pDItem;
                    pGItem->Attach();

                    TotItemsToRead++;
                }
                else {
                    pDItem->Detach();
                }
            }
        } // item is existent and active

        res = m_oaItems.Next(i, &i);
    }
    LeaveCriticalSection(&m_ItemsCritSec);     // finished building item arrays

    if (TotItemsToRead == 0) {
        res = S_FALSE;                            // There are no Device Items to read
        goto UpdateToClient2;
    }

    res = E_OUTOFMEMORY;                         // New default error code
    // Generate all required arrays with size TotItemsToRead
    pErr = (HRESULT *)pIMalloc->Alloc(TotItemsToRead * sizeof(HRESULT));
    if (!pErr)        goto UpdateToClient2;

    pItemStates = (OPCITEMSTATE *)pIMalloc->Alloc(TotItemsToRead * sizeof(OPCITEMSTATE));
    if (!pItemStates) goto UpdateToClient3;

    for (i = 0; i < TotItemsToRead; i++) {
        pErr[i] = S_OK;
        pGItem = ppGItems[i];
        pItemStates[i].hClient = pGItem->get_ClientHandle();
        VariantInit(&pItemStates[i].vDataValue);
        V_VT(&pItemStates[i].vDataValue) = pGItem->get_RequestedDataType();

    }

    // read current values of the items to be handled
    res = InternalRead(OPC_DS_CACHE,           // perform the read
        TotItemsToRead,
        ppDItems,
        pItemStates,
        pErr
    );
    // Note : res is always S_OK or S_FALSE

    if (Killed() == TRUE) {
        res = S_OK;                               // update aborted due to group being deleted
        goto UpdateToClient5;
    }

    //
    //
    // Transfer only value of changed items to the client
    //
    BOOL fItemValueChanged;
    TotItemsToTransmit = 0;
    for (i = 0; i < TotItemsToRead; i++) {

        if (FAILED(pErr[i])) {
            continue;
        }

        fItemValueChanged = FALSE;

        res = ppGItems[i]->CompareLastRead(m_PercentDeadband,
            pItemStates[i].vDataValue,    // The New Value
            pItemStates[i].wQuality,      // The New Quality
            fItemValueChanged);          // The Result we want

        if (FAILED(res)) {
            continue;
        }

        if (fItemValueChanged) {                  // Item Value has changed

              // Copy new value/quality to last read value/quality (even if sending doesn't work)
            res = ppGItems[i]->UpdateLastRead(pItemStates[i].vDataValue, pItemStates[i].wQuality);
            if (FAILED(res)) {
                continue;
            }
            // Only items to transmit are stored in the array.
            // Move the item data in the array.
            if (TotItemsToTransmit != i) {
                pItemStates[TotItemsToTransmit].hClient = pItemStates[i].hClient;
                pItemStates[TotItemsToTransmit].ftTimeStamp = pItemStates[i].ftTimeStamp;
                pItemStates[TotItemsToTransmit].wQuality = pItemStates[i].wQuality;
                pItemStates[TotItemsToTransmit].wReserved = pItemStates[i].wReserved;
                VariantClear(&pItemStates[TotItemsToTransmit].vDataValue);
                VariantCopy(&(pItemStates[TotItemsToTransmit].vDataValue), &(pItemStates[i].vDataValue));
                pErr[TotItemsToTransmit] = pErr[i];
            }
            TotItemsToTransmit++;                  // keep this item

        } // Item Value has changed
    } // Handle all readable items

    if (TotItemsToTransmit) {                   // There are items with changed values -> Transmit.

        if (m_fCallbackEnable) {
            CComObject<DaGroup>*  pCOMGroup;

            m_pServer->CriticalSectionCOMGroupList.BeginReading();       // lock reading

            res = m_pServer->m_COMGroupList.GetElem(m_hServerGroupHandle, &pCOMGroup);
            if (SUCCEEDED(res)) {
                res = pCOMGroup->FireOnDataChange(TotItemsToTransmit, pItemStates, pErr);
            }

            m_pServer->CriticalSectionCOMGroupList.EndReading();     // unlock reading

        }
        if (DataCallbackOnly == FALSE) {          // Also handle IAdviseSink callbacks.

            if (custom) {
                // now pack and send the read values to the requesting client
                res = SendDataStream(WithTime,
                    TotItemsToTransmit,
                    pItemStates,
                    0);          // Transaction Id
            }
            else {
                res = SendDataStreamDisp(WithTime,
                    TotItemsToTransmit,
                    pItemStates,
                    0);           // Transaction Id
            }
        }
        // At least sent one value successfully
        if (SUCCEEDED(res)) {
            res = CoFileTimeNow(&m_pServer->m_LastUpdateTime);
        }
    }

UpdateToClient5:
    for (i = 0; i < TotItemsToRead; i++) {           // release the item values
        VariantClear(&pItemStates[i].vDataValue);
    }

    pIMalloc->Free(pItemStates);

UpdateToClient3:
    pIMalloc->Free(pErr);                      // release error array

UpdateToClient2:
    for (i = 0; i < TotItemsToRead; i++) {           // release the attached items
        _ASSERTE(ppGItems[i]);
        _ASSERTE(ppDItems[i]);
        ppGItems[i]->Detach();
        ppDItems[i]->Detach();
    }
    delete[] ppDItems;                          // free the device item array

UpdateToClient1:
    delete[] ppGItems;                          // free the generic item array

UpdateToClient0:
    return res;
}



//=========================================================================
// UpdateNotify
//
//=========================================================================
HRESULT DaGenericGroup::UpdateNotify(void)
{
    BOOL           WithTime = FALSE;
    BOOL           Custom = FALSE;
    BOOL           DataCallbackOnly = FALSE;

    if ((m_DataCallback == NULL) &&
        (m_DataTimeCallback == NULL) &&
        (m_DataTimeCallbackDisp == NULL) &&
        (m_fCallbackEnable == FALSE)) {
        // try next time
        return S_OK;
    }

    if (m_DataTimeCallback != NULL) {
        WithTime = TRUE;
        Custom = TRUE;
    }
    else if (m_DataCallback != NULL) {
        WithTime = FALSE;
        Custom = TRUE;
    }
    else if (m_DataTimeCallbackDisp != NULL) {
        WithTime = TRUE;
        Custom = FALSE;
    }
    else {
        // There are no IAdviseSink connections
        // Check if there are registered callback functions
        CComObject<DaGroup>* pCOMGroup = NULL;

        m_pServer->CriticalSectionCOMGroupList.BeginReading();          // lock reading
        m_pServer->m_COMGroupList.GetElem(m_hServerGroupHandle, &pCOMGroup);
        if (pCOMGroup == NULL) {
            m_pServer->CriticalSectionCOMGroupList.EndReading();     // unlock reading
            return S_OK;
        }
        IUnknown** pp = pCOMGroup->m_vec.begin(); // Check if there is a registered callback function

        m_pServer->CriticalSectionCOMGroupList.EndReading();        // unlock reading
        if (*pp == NULL) {
            return S_OK;                           // There is no registered callback
        }
        DataCallbackOnly = TRUE;
    }

    return UpdateToClient(Custom, WithTime, DataCallbackOnly);
}


//DOM-IGNORE-END
