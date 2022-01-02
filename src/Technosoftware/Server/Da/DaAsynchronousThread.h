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

#ifndef __DaAsynchronousThread_H_
#define __DaAsynchronousThread_H_

//DOM-IGNORE-BEGIN

class DaGenericGroup;

#if (_ATL_VER < 0x0700)
   #include "UtilityDefs.h"                     // Contains the CAutoVectorPtr template
#endif

class DaAsynchronousThread  {
   friend unsigned int __stdcall Thread_Read_Handler(       void * Par ); 
   friend unsigned int __stdcall Thread_ReadAut_Handler(    void * Par ); 
   friend unsigned int __stdcall Thread_Write_Handler(      void * Par ); 
   friend unsigned int __stdcall Thread_WriteAut_Handler(   void * Par ); 

// Construction
public:
   DaAsynchronousThread(                        // Constructor
      DaGenericGroup *pGGroup );          // Parent GenericItem Object
      
   ~DaAsynchronousThread( void );               // Destructor

// Operations
public:
   //
   // Note :
   //    The creator of the Thread Avice object must not use the created
   //    instance after CreateXXX calls because the object destroys itself.
   //    Also all by parameters provided arrays will be deleted.
   //

   // Create thread for Read or Refresh through Custom Interface
   HRESULT CreateCustomRead(  BOOL           WithTime,            // OLE connection key
                              OPCDATASOURCE  Source,              // CACHE or DEVICE
                              DWORD          NumItems,            // number of items to handle
                              DaDeviceItem  **ppDItems,            // array with generic item ptrs  
                              OPCITEMSTATE  *pItemStates,         // OPCITEMSTATE result array
                              BOOL          *pfItemActiveState,   // array with active state of all generic items
                              BOOL          *pfPhyval,            // array which marks items added with their physical value
                              // [out]
                              DWORD         *pdwTransactionID );


   // Create thread for Read or Refresh through Automation Interface
   HRESULT CreateAutomationRead(
                              OPCDATASOURCE  Source,              // CACHE or DEVICE
                              DWORD          NumItems,            // number of items to handle
                              DaDeviceItem  **ppDItems,            // array with generic item ptrs  
                              OPCITEMSTATE  *pItemStates,         // client handles + requested data type
                              BOOL          *pfPhyval,            // array which marks items added with their physical value
                              // [out]
                              long          *pdwTransactionID );


   // Create thread for Write through Custom Interface
   HRESULT CreateCustomWrite( DWORD          NumItems,            // number of items to handle
                              DaDeviceItem  **ppDItems,            // array with generic item ptrs  
                              OPCITEMHEADERWRITE *pItemHdr,       // result array with client handle
                              VARIANT       *pItemValues,         // array with item handles
                              BOOL          *pfPhyval,            // array which marks items added with their physical value
                              // [out]
                              DWORD         *pdwTransactionID );

     
   // Create thread for Write through Automation Interface
   HRESULT CreateAutomationWrite(
                              DWORD          NumItems,            // number of items to handle
                              DaDeviceItem  **ppDItems,            // array with generic item ptrs  
                              long          *pClientHandles,      // result array with client handle
                              VARIANT       *pItemValues,         // array with item handles
                              BOOL          *pfPhyval,            // array which marks items added with their physical value
                              // [out]
                              long          *pdwTransactionID );


   HRESULT  Cancel( void );

   inline BOOL IsTransactionID( DWORD dwTransactionID )
                        { return (dwTransactionID == m_TransactionID); }

// Implementation
protected:

   //
   // Data members
   //

            // Thread handle. Also used as transaction id.
   uintptr_t      m_ThreadHandle;

            // Transaction id.
   DWORD          m_TransactionID;

            // used to synchronize access to this class instance
   BOOL           m_ToCancel;

            // parent class
   DaGenericGroup *m_pParent;

   unsigned       m_ThreadId;

            // defines stream format to client
   BOOL           m_WithTime;

            // Number of itmes to be handled
   DWORD          m_NumItems ;

            // Array with items to be handled
   DaDeviceItem  **m_ppDItems ;

            // Source:  CACHE or DEVICE
   OPCDATASOURCE  m_DataSource;

            // Answer Item Header Array  ( Custom Interface )
   OPCITEMHEADERWRITE * m_pItemHdr;

            // Answer client handle Array ( Automation interface )
   long          *m_pClientHnd;

            // Ptr to Array with read item values
   OPCITEMSTATE  *m_pItemStates ;

            // Array with values, qualities and timestamps to write
   CAutoVectorPtr<OPCITEMVQT>
                  m_avpVQTsToWrite;

            // Array with active state of all generic items
   BOOL          *m_pfItemActiveState;

            // Array to mark items added with their physical value (CALL-R)
   BOOL          *m_pfPhyval;

   //
   // Function members
   //
   void DoKill( void );
   void DoDelete( void );
   void RemoveAdviseFromGroup( void );
};
//DOM-IGNORE-END

#endif // __DaAsynchronousThread_H_
