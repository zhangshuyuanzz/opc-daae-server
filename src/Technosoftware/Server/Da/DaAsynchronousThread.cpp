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
#include <process.h>
#include "UtilityFuncs.h"   
#include "DaGenericGroup.h"

static DWORD sCurrentTransactionID = 0;

//=========================================================================
// Constructor
//=========================================================================
DaAsynchronousThread::DaAsynchronousThread( DaGenericGroup *pGGroup )
{
   m_pParent   = pGGroup ;

            // Attach to the group so that it cannot be deleted until 
            // all the threads running over it are killed
   m_pParent->Attach();

   m_ToCancel = FALSE ;

            // Number of itmes to be handled
   m_NumItems = 0;

            // set to invalid handle
   m_ThreadHandle = 0;

            //
            // Pointer to memory areas.
            // Must be released in destructor.
            //

            // Pointers to memory allocated outside of class
            // ---------------------------------------------

            // Array with items to be handled
   m_ppDItems = NULL;

            // Answer Item Header Array  ( Custom Interface )
   m_pItemHdr = NULL;
   
            // Answer client handle Array ( Automation interface )
   m_pClientHnd = NULL;                            // use delete
   
            // Ptr to Array with read item values
   m_pItemStates = NULL;                           // use delete

            // Array with active state of all generic items
   m_pfItemActiveState = NULL;                     // use new/delete

            // Array to mark items added with their physical value (CALL-R)
   m_pfPhyval = NULL;
}



//=========================================================================
// Destructor
//=========================================================================
DaAsynchronousThread::~DaAsynchronousThread( void )
{
   DWORD i;

   // Release External Allocated Memory
   // ---------------------------------
   
      // release the read items!
   if (m_ppDItems) {
      for (i=0; i<m_NumItems; i++) {
         if (m_ppDItems[i] != NULL) {
            m_ppDItems[i]->Detach();     // release the items
         }
      }
      // free the item array
      delete [] m_ppDItems;
   }
   
   
            // Answer Item Header Array  ( Custom Interface )
   if (m_pItemHdr) {
      delete [] m_pItemHdr;
      m_pItemHdr = NULL;
   }
   
      // Answer client handle Array ( Automation interface )
   if (m_pClientHnd) {
      delete [] m_pClientHnd;
      m_pClientHnd = NULL;
   }
   
      // Ptr to Array with read item values
   if (m_pItemStates) {
      for (i=0; i<m_NumItems; i++) {
         VariantClear( &m_pItemStates[i].vDataValue );            
      }
      delete [] m_pItemStates;
      m_pItemStates = NULL;
   }
   
      // release array with active state of all generic items
   if (m_pfItemActiveState) {
      delete [] m_pfItemActiveState;
      m_pfItemActiveState = NULL;
   }
   
      // release array which marks items added with their physical value
   if (m_pfPhyval) {
      delete [] m_pfPhyval;
      m_pfPhyval = NULL;
   }

   
   // Release Internal Allocated Memory
   // ---------------------------------
   
      // Array with values, qualities and timestamps to write
   if (m_avpVQTsToWrite) {
      for (i=0; i<m_NumItems; i++) {
         VariantClear( &m_avpVQTsToWrite[i].vDataValue );
      }
      m_avpVQTsToWrite.Free();
   }
   
   m_ThreadHandle = 0;
   m_NumItems = 0;
}



//=========================================================================
// Init the class instance and start the handler thread for READ and
// REFRESH through the CUSTOM interface
//=========================================================================
HRESULT DaAsynchronousThread::CreateCustomRead( 
                   BOOL           WithTime,
                   OPCDATASOURCE  Source,             // CACHE or DEVICE
                   DWORD          NumItems,           // number of items to handle
                   DaDeviceItem  **ppDItems,           // array with generic item ptrs  
                   OPCITEMSTATE  *pItemStates,        // client handles + requested data type
                   BOOL          *pfItemActiveState,  // array with active state of all generic items
                   BOOL          *pfPhyval,           // array which marks items added with their physical value
                   // [out]
                   DWORD         *pdwTransactionID )
{
      unsigned    dwCreationFlags = 0 ;

                                                      
   m_pfItemActiveState = pfItemActiveState;           // array with active state of all generic items

   m_pfPhyval     = pfPhyval;                         // array which marks items added with their physical value

   m_WithTime     = WithTime;

   m_DataSource   = Source;
   m_NumItems     = NumItems;
      // Device Items to be read
   m_ppDItems     = ppDItems;
      // Item States 
   m_pItemStates  = pItemStates;

   sCurrentTransactionID++;

   *pdwTransactionID = m_TransactionID = sCurrentTransactionID;

   m_ThreadHandle = _beginthreadex(
                              NULL,                   // no thread security attributes
                              0,                      // default stack size  
                              Thread_Read_Handler,    // pointer to thread function 
                              this,                   // pass class to new thread 
                              dwCreationFlags,        // into ready state 
                              &m_ThreadId  );         // pointer to returned thread identifier 

   if (m_ThreadHandle == 0) {
      DoDelete();                                     // Cannot create thread, delete advise class
      return HRESULT_FROM_WIN32( GetLastError() );    // because not deleted by thread handler.

   }
   else {
      return S_OK;
   }
}



//=========================================================================
// Init the class instance and start the handler thread for READ, REFRESH
// through the AUTOMATION Interface
//
// Thread Type is ignored in this implementation!
//=========================================================================
HRESULT DaAsynchronousThread::CreateAutomationRead( 
                   OPCDATASOURCE  Source,             // CACHE or DEVICE
                   DWORD          NumItems,           // number of items to handle
                   DaDeviceItem  **ppDItems,           // array with generic item ptrs  
                   OPCITEMSTATE  *pItemStates,        // client handles + requested data type
                   BOOL          *pfPhyval,           // array which marks items added with their physical value
                   // [out]
                   long          *pdwTransactionID )
{
   m_pfPhyval     = pfPhyval;                         // array which marks items added with their physical value
   m_DataSource   = Source ;
   m_NumItems     = NumItems ;
   m_ppDItems     = ppDItems ;
   m_pItemStates  = pItemStates ;
   m_WithTime     = TRUE ;                

   sCurrentTransactionID++;

   *pdwTransactionID = m_TransactionID = sCurrentTransactionID;

   m_ThreadHandle = _beginthreadex(
                              NULL,                   // no thread security attributes
                              0,                      // default stack size  
                              Thread_ReadAut_Handler, // pointer to thread function 
                              this,                   // pass class to new thread 
                              0,                      // into ready state 
                              &m_ThreadId  );         // pointer to returned thread identifier 

   if (m_ThreadHandle == 0) {
      DoDelete();                                     // Cannot create thread, delete advise class
      return HRESULT_FROM_WIN32( GetLastError() );    // because not deleted by thread handler.

   }
   else {
      return S_OK;
   }
}




//=========================================================================
// Init the class instance and start the handler thread for WRITE
// through the CUSTOM Interface
//=========================================================================
HRESULT DaAsynchronousThread::CreateCustomWrite( 
                   DWORD          NumItems,           // number of items to handle
                   DaDeviceItem  **ppDItems,           // array with generic item ptrs
                   OPCITEMHEADERWRITE *pItemHdr,      // Answer Item Header Array
                   VARIANT       *pItemValues,        // array with item values
                   BOOL          *pfPhyval,           // array which marks items added with their physical value
                   // [out]
                   DWORD         *pdwTransactionID )
{
   HRESULT  hr = S_OK;

   try {
      m_pfPhyval  = pfPhyval;                   // Array which marks items added with their physical value
      m_NumItems  = NumItems;
      m_ppDItems  = ppDItems;
      m_pItemHdr  = pItemHdr;

      // Map the Values to Write into OPCITEMVQT
      if (!m_avpVQTsToWrite.Allocate( NumItems )) throw E_OUTOFMEMORY;

      memset( m_avpVQTsToWrite, 0, sizeof (OPCITEMVQT) * NumItems );
      for (DWORD i = 0; i < NumItems; i++) {
                                                // Deep copy
         hr = VariantCopy( &m_avpVQTsToWrite[i].vDataValue, &pItemValues[i] );
         _OPC_CHECK_HR( hr );
      }

      sCurrentTransactionID++;

      *pdwTransactionID = m_TransactionID = sCurrentTransactionID;

      m_ThreadHandle = _beginthreadex(
                        NULL,                   // no thread security attributes
                        0,                      // default stack size  
                        Thread_Write_Handler,   // pointer to thread function
                        this,                   // pass class to new thread 
                        0,                      // into ready state 
                        &m_ThreadId  );         // pointer to returned thread identifier 

      if (m_ThreadHandle == 0) {
         hr = HRESULT_FROM_WIN32( GetLastError() );
         throw hr;
      }
      hr = S_OK;
   }
   catch (HRESULT hrEx) { hr = hrEx; }
   catch (...) { hr = E_FAIL; }

   if (FAILED( hr )) {
      DoDelete();                               // Delete advise class because not deleted by 
   }                                            // thread handler (because not started).
   return hr;
}



//=========================================================================
// Init the class instance and start the handler thread for WRITE
// through the AUTOMATION Interface
//=========================================================================
HRESULT DaAsynchronousThread::CreateAutomationWrite( 
                   DWORD          NumItems,           // number of items to handle
                   DaDeviceItem  **ppDItems,           // array with generic item ptrs
                   long          *pClientHandles,     // Answer client handle Array
                   VARIANT       *pItemValues,        // array with item values
                   BOOL          *pfPhyval,           // array which marks items added with their physical value
                   // [out]
                   long          *pdwTransactionID )
{
         DWORD    i;
         HRESULT  hres;

   m_pfPhyval     = pfPhyval;                         // array which marks items added with their physical value
   m_NumItems     = NumItems;
   m_ppDItems     = ppDItems;
   m_pClientHnd   = pClientHandles;
  
   // Map the Values to Write into OPCITEMVQT
   if (!m_avpVQTsToWrite.Allocate( NumItems )) {
      hres = E_OUTOFMEMORY;
      goto CreateAutoWriteExit1;
   }
  
   memset( m_avpVQTsToWrite, 0, sizeof (OPCITEMVQT) * NumItems );
   for (i = 0; i < NumItems; i++) {
                                                // The item values provided by the user are copied outside of this function.
                                                // Shallow copy
      m_avpVQTsToWrite[i].vDataValue = pItemValues[i];
   }

   sCurrentTransactionID++;

   *pdwTransactionID = m_TransactionID = sCurrentTransactionID;

   m_ThreadHandle = _beginthreadex(
                              NULL,                   // no thread security attributes
                              0,                      // default stack size  
                              Thread_WriteAut_Handler,// pointer to thread function
                              this,                   // pass class to new thread 
                              0,                      // into ready state 
                              &m_ThreadId  );         // pointer to returned thread identifier 

   if (m_ThreadHandle == 0) { 
      hres = HRESULT_FROM_WIN32( GetLastError() );
      goto CreateAutoWriteExit1;
   }
   return S_OK;

CreateAutoWriteExit1:
   DoDelete();                                        // Delete advise class because not deleted by 
                                                      // thread handler (because not started).
   return hres;
}


    
//=========================================================================
// Kill this instance and the thread running over it
//=========================================================================
void DaAsynchronousThread::DoKill( void )
{
                                 // remove from list in group so that no AsyncIO.Cancel
                                 // can be done on it.
   RemoveAdviseFromGroup();

   HANDLE hThread = (HANDLE)m_ThreadHandle;

   m_pParent->Detach();
   delete this;
                                 // Version 1.3, AM
   CloseHandle( hThread );       //    Thread handle must be closed explicitly
                                 //    if _endthreadex() is used.

   _endthreadex(0);              // thread kill requested
}




//=========================================================================
// Delete the instance.
// called only if CreateXXX functions failed.
//=========================================================================
void DaAsynchronousThread::DoDelete( void )
{
                                 // remove from list in group so that no AsyncIO.Cancel
                                 // can be done on it.
   RemoveAdviseFromGroup();
   delete this;                  // All cleanup stuff is done by destructor.
}




//=========================================================================
// Cancel the Asynch, Read, Write and Refresh Thread and object
//=========================================================================
HRESULT DaAsynchronousThread::Cancel(void)
{
   m_ToCancel = TRUE;
   return S_OK;
}




//=========================================================================
// take away given thread object from list in group
//    returns  E_FAIL if thread has been canceled or group killed!
//             S_OK   else.
//=========================================================================
void DaAsynchronousThread::RemoveAdviseFromGroup( void )
{
         long              i;
         HRESULT           res;
         DaAsynchronousThread     *pCurAdv;

            
      // protect thread object list while searching and removing
   EnterCriticalSection( &m_pParent->m_AsyncThreadsCritSec );

   res = m_pParent->m_oaAsyncThread.First( &i );

   while (SUCCEEDED(res)) {
      m_pParent->m_oaAsyncThread.GetElem( i, &pCurAdv );
      if ( pCurAdv == this ) {
            // found the pointer to this object ...
            // delete it from the group
         m_pParent->m_oaAsyncThread.PutElem( i, NULL );
         goto RemoveAdviseFromGroupExitWhile;
      }
      res = m_pParent->m_oaAsyncThread.Next( i, &i );
   }

      // here we could raise an exception because 
      // the object should always be found in the 
      // group list

RemoveAdviseFromGroupExitWhile:
   LeaveCriticalSection( &m_pParent->m_AsyncThreadsCritSec );
}




///////////////////////////////////////////////////////////////////////////
//////////////////   Handler Threads  /////////////////////////////////////
///////////////////////////////////////////////////////////////////////////
  
//=========================================================================
// Thread handling Async READ and Refresh throough CUSTOM Interface
// ----------------------------------------------------------------
// The thread handles a single request and then terminates.
// This thread is also created from the permanent update thread
//=========================================================================
unsigned int __stdcall Thread_Read_Handler( void *Par ) 
{
         DaAsynchronousThread    *Adv;
         DWORD             i;
         HRESULT          *pErr;
         HRESULT           res;
         DaGenericGroup    *group;               // Parent DaGenericGroup
         DWORD             dwNumItem;
         DaDeviceItem     **ppDItems;
         OPCITEMSTATE     *pItemStates;
         BOOL              fQualityOutOfService;
         DWORD             dwNumOfDItems = 0;   // Number of successfully attached DeviceItems

   Adv         = (DaAsynchronousThread *)Par;
   group       = Adv->m_pParent;
   dwNumItem   = Adv->m_NumItems;
   ppDItems    = Adv->m_ppDItems;
   pItemStates = Adv->m_pItemStates;

      // create errors array for InternalRead
   pErr = new HRESULT [dwNumItem];
   if (pErr == NULL) { 
      res = E_OUTOFMEMORY;
      goto ThreadCustReadExit0;
   }

   for (i=0; i<dwNumItem; i++) {
      if (ppDItems[i] != NULL) {

         pErr[i] = ppDItems[i]->HasReadAccess() ? S_OK : OPC_E_BADRIGHTS;
         if (FAILED( pErr[i] )) {
               // this item cannot be read
            ppDItems[i]->Detach();
            ppDItems[i] = NULL;
         }
         else {
            dwNumOfDItems++;
         }
      }
      else {
         pErr[i] = E_FAIL;
      }
   }
         // success
   if (dwNumOfDItems) {
      res = group->InternalRead(                // perform the read
                        Adv->m_DataSource,
                        dwNumItem,
                        ppDItems,
                        pItemStates,
                        pErr
                        );
   }
   else {
         // There are no Device Items to read
      res = S_FALSE;
   }
   // Note : res is always S_OK or S_FALSE
                                                // Handle the Quality

      // if read from cache and group inactive then
      // all items must return quality OPC_QUALITY_OUT_OF_SERVICE
   if( ( Adv->m_DataSource == OPC_DS_CACHE ) && ( group->GetActiveState() == FALSE ) ) {
      fQualityOutOfService = TRUE;
   } else {
      fQualityOutOfService = FALSE;
   }

   for( i=0 ; i < dwNumItem ; i++ ) {
      if( SUCCEEDED( pErr[i]) ) {
                                                // Set Quality to 'Bad: Out of Service' if
                                                // Group or Item is inactive.
         if (fQualityOutOfService ||                                       // Group is inactive 
            ( (Adv->m_DataSource == OPC_DS_CACHE) &&                       // Item is inactive
              (Adv->m_pfItemActiveState[i] == FALSE) ) ) {

            pItemStates[i].wQuality =                                      // Set Quality to 'Bad: Out of Service'
                        OPC_QUALITY_BAD |                                  // new quality is 'bad'
                        OPC_QUALITY_OUT_OF_SERVICE |                       // new substatus is 'out of service'
                        (pItemStates[i].wQuality & OPC_LIMIT_MASK);        // left the limit unchanged
         }
      }
   }

      // if the group is not about to be killed
      // and the advise thread was not canceled from outside
      // send the data to the source ...
   if ( (group->Killed() == TRUE) || (Adv->m_ToCancel == TRUE) ) {
         // send nothing
      goto ThreadCustReadExit1;
   }

      // now pack and send the read values to the requesting client
   group->SendDataStream(  Adv->m_WithTime,        // may be TRUE or FALSE
                           dwNumItem,
                           Adv->m_pItemStates,
                           Adv->m_TransactionID    // Transaction Id
                        );     

ThreadCustReadExit1:
     // clean up internals
   delete [] pErr;

ThreadCustReadExit0:
      // release the read items!

      // kill thread and object
   Adv->DoKill();
      // should never get here
   return 0; 
}




//=========================================================================
// Thread handling Async READ and Refresh through AUTOMATION Interface
// -------------------------------------------------------------------
// The thread handles a single request and then terminates.
// This thread is also created from the permanent update thread
//=========================================================================
unsigned int __stdcall Thread_ReadAut_Handler( void *Par ) 
{
         DaAsynchronousThread  *pAdv;
         DWORD           i;
         long            *pErr;
         DaGenericGroup   *group;
         DWORD            numItems;
         HRESULT          res;
         OPCITEMSTATE    *pItemStates;
         DaDeviceItem    **ppDItems;
         DWORD            dwNumOfDItems = 0; // Number of successfully attached DeviceItems

   pAdv        = (DaAsynchronousThread*) Par;
   group       = pAdv->m_pParent;
   numItems  = pAdv->m_NumItems;
   pItemStates = pAdv->m_pItemStates;
   ppDItems    = pAdv->m_ppDItems;

      // create errors array
   pErr = new HRESULT[ numItems ];
   if( pErr == NULL ) {   
      res = E_OUTOFMEMORY;
      goto ThreadAutReadExit0;
   }
   i = 0;
   while ( i < numItems ) {
      pErr[i] = S_OK;
      i ++;
   }

   for (i=0; i<numItems; i++) {
      if ( ppDItems[i] != NULL ) {

         pErr[i] = ppDItems[i]->HasReadAccess() ? S_OK : OPC_E_BADRIGHTS;
         if (FAILED( pErr[i] )) {
               // this item cannot be read
            ppDItems[i]->Detach();
            ppDItems[i] = NULL;
         }
         else {
            dwNumOfDItems++;
         }

      } else {
         pErr[i] = E_FAIL;
      }
   }

   if (dwNumOfDItems) {
      res = group->InternalRead(                // perform the read
                        pAdv->m_DataSource, 
                        numItems, 
                        ppDItems, 
                        pItemStates,
                        pErr
                        );
   }
   else {
         // There are no Device Items to read
      res = S_FALSE;
   }
   // Note : res is always S_OK or S_FALSE

      // if the group is not about to be killed
      // and the advise thread was not canceled from outside
      // send the data to the source ...
   if ( (group->Killed() == TRUE) || (pAdv->m_ToCancel == TRUE) ) {
         // send nothing
      goto ThreadAutReadExit1;
   }

      // now pack and send the read values to the requesting client
   group->SendDataStreamDisp(  pAdv->m_WithTime,        // may be TRUE or FALSE
                               numItems,
                               pItemStates,
                               pAdv->m_TransactionID    // Transaction Id
                            ) ;     



ThreadAutReadExit1:
     // clean up internals
   delete [] pErr;

ThreadAutReadExit0:
      // kill thread and object
   pAdv->DoKill();
      // should never get here
   return 0;                                                 // should never get here!
}






//=========================================================================
// Thread handling Async WRITE requested through the CUSTOM Interface
// ------------------------------------------------------------------
// The thread handles a single request and then terminates.
//=========================================================================
unsigned int  __stdcall Thread_Write_Handler( void *Par ) 
{
         DaAsynchronousThread *Adv;
         HRESULT       *pErr;
         HRESULT        res;
         DWORD          i;
         DaGenericGroup *group;               // Parent DaGenericGroup
         DWORD          NumItems;
         DaDeviceItem  **ppDItems;
         DWORD          dwNumOfDItems = 0;   // Number of successfully attached DeviceItems

   Adv      = (DaAsynchronousThread*) Par;
   group    = Adv->m_pParent ;
   NumItems = Adv->m_NumItems;
   ppDItems = Adv->m_ppDItems;

   pErr = new HRESULT [ NumItems ];
   if( pErr == NULL ) {                      // success
      goto ThreadCustWriteExit0;
   }
   for (i=0; i<NumItems; i++) {
      pErr[i] = S_OK;
   }

   for (i=0; i<NumItems; i++) {
      if ( ppDItems[i] != NULL ) {

         pErr[i] = ppDItems[i]->HasWriteAccess() ? S_OK : OPC_E_BADRIGHTS;
         if (FAILED( pErr[i] )) {
               // this item cannot be read
            ppDItems[i]->Detach();
            ppDItems[i] = NULL;
         }
         else {
            dwNumOfDItems++;
         }

      } else {
         pErr[i] = E_FAIL;
      }
   }

   if (dwNumOfDItems) {
         // write the item values into the cache
      res = group->InternalWriteVQT(
                        NumItems,                    
                        ppDItems, 
                        Adv->m_avpVQTsToWrite,
                        pErr
                        );
   }
   else {
         // There are no Device Items to write.
      res = S_FALSE;
   }

      // check if group killed or thread canceled
   if ( (group->Killed() == TRUE) || (Adv->m_ToCancel == TRUE) ) {
         // send nothing
      goto ThreadCustWriteExit1;
   }

   for (i=0; i<NumItems; i++) {
      Adv->m_pItemHdr[i].dwError = pErr[i];
   }

      // Advise the client
   res = group->SendWriteStream( NumItems, 
                                 Adv->m_pItemHdr,
                                 res,
                                 Adv->m_TransactionID );

ThreadCustWriteExit1:
     // clean up internals
   delete [] pErr;

ThreadCustWriteExit0:
      // kill thread and object
   Adv->DoKill();
   return 0;                                             // should never get here
}



//=========================================================================
// Thread handling Async WRITE requested through the AUTOMATION Interface
// ----------------------------------------------------------------------
// The thread handles a single request and then terminates.
//
// 
//=========================================================================
unsigned int  __stdcall Thread_WriteAut_Handler( void *Par ) 
{
         DaAsynchronousThread *pAdv;
         HRESULT       *pErr;
         HRESULT        res;
         DWORD          i;
         DaGenericGroup *group;     
         DWORD          NumItems;
         DaDeviceItem  **ppDItems;
         DWORD          dwNumOfDItems = 0;   // Number of successfully attached DeviceItems

   pAdv     = (DaAsynchronousThread*) Par ;                   
   group    = pAdv->m_pParent;
   NumItems = pAdv->m_NumItems;
   ppDItems = pAdv->m_ppDItems;


   pErr = new HRESULT [ NumItems ];
   if( pErr == NULL ) {                      // success
      goto ThreadAutoWriteExit0;
   }
   i = 0;
   while ( i < NumItems ) {
      pErr[i] = S_OK;
      i ++;
   }

   for (i=0; i<NumItems; i++) {
      if ( ppDItems[i] != NULL ) {

            // ... inverse operation is done when thread ended with DeviceItem.Detach() !
         pErr[i] = ppDItems[i]->HasWriteAccess() ? S_OK : OPC_E_BADRIGHTS;
         if (FAILED( pErr[i] )) {
               // this item cannot be read
            ppDItems[i]->Detach();
            ppDItems[i] = NULL;
         }
         else {
            dwNumOfDItems++;
         }

      } else {
         pErr[i] = E_FAIL;
      }
   }

   if (dwNumOfDItems) {
         // write the item values into the cache
      res = group->InternalWriteVQT(
                        NumItems,
                        ppDItems,
                        pAdv->m_avpVQTsToWrite,
                        pErr
                        );
   }
   else {
         // There are no Device Items to write.
      res = S_FALSE;
   }

      // check if group killed or thread canceled
   if ( (group->Killed() == TRUE) || (pAdv->m_ToCancel == TRUE) ) {
         // send nothing
      goto ThreadAutoWriteExit1;
   }


      // Advise the client
   res = group->SendWriteStreamDisp(   NumItems, 
                                       pAdv->m_pClientHnd,
                                       pErr,
                                       res,
                                       pAdv->m_TransactionID );

ThreadAutoWriteExit1:
     // clean up internals
   delete [] pErr;

ThreadAutoWriteExit0:
      // kill thread and object
   pAdv->DoKill();
      // should never get here
   return 0;                                             
}


//DOM-IGNORE-END
