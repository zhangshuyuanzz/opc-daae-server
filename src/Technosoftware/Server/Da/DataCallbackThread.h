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

#if !defined(DATACALLBACKTHREAD_H)
#define DATACALLBACKTHREAD_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "DaGenericGroup.h"
#include "FixOutArray.h"

//-----------------------------------------------------------------------
// CLASS
//-----------------------------------------------------------------------
class DataCallbackThread
{
   friend unsigned __stdcall DataCallbackReadThreadHandler( void* pCreator );
   friend unsigned __stdcall DataCallbackWriteThreadHandler( void* pCreator );
   friend unsigned __stdcall CancelThreadHandler( void* pCreator );

public:
   DataCallbackThread( DaGenericGroup* pParent );
   virtual ~DataCallbackThread();

// Operations
public:
   //
   // Note :
   //    The creator of the Data Callback Thread object must not use the created
   //    instance after CreateXXX calls because the object destroys itself.
   //    Also all by parameters provided arrays will be deleted.
   //

   inline HRESULT CreateCustomRead(
               DWORD          dwCount,          // Number of items to handle
               DWORD          dwTransactionID,  // Transaction ID
               DaDeviceItem**  ppDItems,         // Array with device item ptrs  
               OPCITEMSTATE*  pItemStates,      // Client handles + requested data types
               BOOL*          pfPhyval,         // Array which marks items added with their physical value
               // [out]       
               DWORD*         pdwCancelID )
               {
                  return CreateCustomReadOrRefresh(
                              dwCount, dwTransactionID, ppDItems,
                              pItemStates, NULL, pfPhyval, pdwCancelID );
               }

   inline HRESULT CreateCustomRefresh(
               OPCDATASOURCE  dwSource,
               DWORD          dwCount,          // Number of items to handle
               DWORD          dwTransactionID,  // Transaction ID
               DaDeviceItem**  ppDItems,         // Array with device item ptrs  
               OPCITEMSTATE*  pItemStates,      // Client handles + requested data types
               BOOL*          pfPhyval,         // Array which marks items added with their physical value
               // [out]       
               DWORD*         pdwCancelID )
               {
                  m_fReadTransaction= FALSE;    // Refresh transaction
                  m_dwSource        = dwSource; // Source specified by the user

                  return CreateCustomReadOrRefresh(
                              dwCount, dwTransactionID, ppDItems,
                              pItemStates, NULL, pfPhyval, pdwCancelID );
               }

   inline HRESULT CreateCustomReadMaxAge(
               DWORD          dwCount,          // Number of items to handle
               DWORD          dwTransactionID,  // Transaction ID
               DaDeviceItem**  ppDItems,         // Array with device item ptrs  
               OPCITEMSTATE*  pItemStates,      // Client handles + requested data types
               DWORD*         pdwMaxAge,        // Array with Max Age definitions
               BOOL*          pfPhyval,         // Array which marks items added with their physical value
               // [out]       
               DWORD*         pdwCancelID )
               {
                  return CreateCustomReadOrRefresh(
                              dwCount, dwTransactionID, ppDItems,
                              pItemStates, pdwMaxAge, pfPhyval, pdwCancelID );
               }

   inline HRESULT CreateCustomRefreshMaxAge(
               DWORD          dwCount,          // Number of items to handle
               DWORD          dwTransactionID,  // Transaction ID
               DaDeviceItem**  ppDItems,         // Array with device item ptrs  
               OPCITEMSTATE*  pItemStates,      // Client handles + requested data types
               DWORD*         pdwMaxAge,        // Array with Max Age definitions
               BOOL*          pfPhyval,         // Array which marks items added with their physical value
               // [out]       
               DWORD*         pdwCancelID )
               {
                  m_fReadTransaction= FALSE;    // Refresh transaction

                  return CreateCustomReadOrRefresh(
                              dwCount, dwTransactionID, ppDItems,
                              pItemStates, pdwMaxAge, pfPhyval, pdwCancelID );
               }

   HRESULT CreateCustomWriteVQT(
               DWORD          dwCount,          // Number of items to handle
               DWORD          dwTransactionID,  // Transaction ID
               DaDeviceItem**  ppDItems,         // Array with device item ptrs
               OPCITEMSTATE*  pItemStates,      // Client handles + requested data types
               OPCITEMVQT*    pItemVQTs,
               BOOL*          pfPhyval,         // Array which marks items added with their physical value
               // [out]
               DWORD*         pdwCancelID );

   HRESULT CreateCancelThread();
                                                
   static void SetCallbackResultsFromItemStates(// Sets the values in the arrays with the  
                                                // values from pItemStates
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
            FILETIME*      pftTimeStamps );

// Implementation
protected:
   HRESULT CreateCustomReadOrRefresh(
               DWORD          dwCount,          // Number of items to handle
               DWORD          dwTransactionID,  // Transaction ID
               DaDeviceItem**  ppDItems,         // Array with device item ptrs  
               OPCITEMSTATE*  pItemStates,      // Client handles + requested data types
               DWORD*         pdwMaxAge,        // Array with Max Age definitions, may be a NULL pointer
               BOOL*          pfPhyval,         // Array which marks items added with their physical value
               // [out]       
               DWORD*         pdwCancelID );

   class CCancelThread
   {
      friend unsigned __stdcall CancelThreadHandler( void* pCreator );
   
   // Construction
   public:
      CCancelThread();
      virtual ~CCancelThread();
   
      HRESULT  Create( DataCallbackThread* pThreadToCancel );
   
   // Attributes
   public:
      HANDLE               m_hThread;              // Thread handle

   // Operations
   public:
      inline void InvokeCancelCallback( void ) {
         _ASSERTE( m_hOperationCanceledEvent );
         SetEvent( m_hOperationCanceledEvent );
      }
   
   // Implementation
   protected:
      HANDLE               m_hOperationCanceledEvent;
      DataCallbackThread* m_pThreadToCancel;      // The onwer
   
   };

   //
   // Data members
   //
   DaGenericGroup*    m_pParent;                 // A generic group
   DaGenericServer*   m_pGServer;                // A generic server
   CCancelThread*    m_pCancelThread;           // Pointer to cancel thread if cancel request
   DWORD             m_dwCount;                 // Number of items
   DWORD             m_dwTransactionID;         // Transaction ID
   OPCDATASOURCE     m_dwSource;                // Data source
   BOOL              m_fReadTransaction;        // The activated transaction type (read or refresh)
   HANDLE            m_hThread;                 // Thread handle
   DWORD             m_dwCancelID;              // The cancel ID is the index in the
                                                // arry of data callback threads.
   FILETIME          m_ftNow;                   // Used by functions with Max Age parameters
   
      // Arrays with memory allocated from the heap outside of the
      // class and passed via parameter list of CreateXXX functions.

   OPCITEMSTATE*     m_pItemStates;             // Array with read item values
   DaDeviceItem**     m_ppDItems;                // Array with Device Items to be read
   BOOL*             m_pfPhyval;                // Array which marks items added with their physical value
   OPCITEMVQT*       m_pVQTsToWrite;            // Array with Values, Qualities and TimeStamps to write
   DWORD*            m_pdwMaxAge;               // Array with Max Age definitions

      // Arrays with internal allocated memory from the heap.
   DaDeviceItem**     m_ppTmpBufferForItemsToReadFromDevice;

      // Arrays with internal allocated memory from the COM memory manager.
      // This arrays are passed via the registered callback to the client.
      // Memory for this arrays is allocated in the CreateXXX function.
      // The DataCallbackXXXThreadHandler sets the values.

      // Used by Read/Refresh and Write transactions
   CFixOutArray< OPCHANDLE >  m_fxaHandles;
   CFixOutArray< HRESULT >    m_fxaErrors;
      // Used by Read/Refresh transactions only
   CFixOutArray< VARIANT >    m_fxaVal;
   CFixOutArray< WORD >       m_fxaQual;
   CFixOutArray< FILETIME >   m_fxaTStamps;

      // Pointers to the arrays above
   OPCHANDLE*        m_phClientItems;
   VARIANT*          m_pvValues;
   HRESULT*          m_pErrors;
   WORD*             m_pwQualities;
   FILETIME*         m_pftTimeStamps;

   //
   // Function members
   //
   void DoKill( void );
   void CheckCancelRequestAndKillFlag( BOOL fPreventCancel = FALSE );
};
//DOM-IGNORE-END

#endif // !defined(DATACALLBACKTHREAD_H)
