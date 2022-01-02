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

//-------------------------------------------------------------------------
// INCLUDE
//-------------------------------------------------------------------------
#include <windows.h>
#include "ReadWriteLock.h"

//-------------------------------------------------------------------------
// CODE
//-------------------------------------------------------------------------


ReadWriteLock::ReadWriteLock( void )
{
   readMutex_		= NULL;
   writeMutex_		= NULL;
   readerEvent_		= NULL;
   counter_			= 0;
}


ReadWriteLock::~ReadWriteLock()
{
   if( readerEvent_ ) {  
	   CloseHandle(readerEvent_);
   }
   if( readMutex_ ) {
	   CloseHandle(readMutex_);
   }
   if( writeMutex_ ) {
	   CloseHandle(writeMutex_);
   }
}


BOOL ReadWriteLock::Initialize()
{
   readerEvent_=CreateEvent(NULL,			// pointer to security attributes
                            TRUE,				// flag for manual-reset event 
                            FALSE,				// flag for initial state 
                            NULL);			// pointer to event-object name 
   if( readerEvent_ == NULL ) {
      return FALSE;
   }
   readMutex_ = CreateEvent(NULL,			// pointer to security attributes
                         FALSE,					// flag for manual-reset event 
                         TRUE,					// flag for initial state 
                         NULL);				// pointer to event-object name 
   if( readMutex_ == NULL ){
      CloseHandle(readerEvent_);
      return FALSE;
   }

   writeMutex_ = CreateMutex(NULL,			// pointer to security attributes 
                                FALSE,			// flag for initial ownership 
                                NULL);		// pointer to mutex-object name 
   if( writeMutex_ == NULL ){
      CloseHandle(readerEvent_);
      CloseHandle(readMutex_);
      return FALSE;
   }
   counter_ = -1;
   return TRUE;
}


void ReadWriteLock::BeginReading()
{ 
   if (InterlockedIncrement(&counter_) == 0)
   {
      WaitForSingleObject(readMutex_, INFINITE);
      SetEvent(readerEvent_);
   }
   WaitForSingleObject(readerEvent_,INFINITE);
}


void ReadWriteLock::BeginWriting()
{ 
   WaitForSingleObject(writeMutex_,INFINITE);
   WaitForSingleObject(readMutex_, INFINITE);
}


void ReadWriteLock::EndReading()
{
   if (InterlockedDecrement(&counter_) < 0)
   {
      ResetEvent(readerEvent_);
      SetEvent(readMutex_);
   }
}


void ReadWriteLock::EndWriting()
{
   SetEvent(readMutex_);
   ReleaseMutex(writeMutex_);
}


//DOM-IGNORE-END
