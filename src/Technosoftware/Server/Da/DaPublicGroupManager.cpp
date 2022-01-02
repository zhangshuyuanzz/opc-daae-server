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
#include "DaPublicGroupManager.h"
#include "DaPublicGroup.h"
#include "UtilityFuncs.h"

//====================================================================
// Constructor
//====================================================================
DaPublicGroupManager::DaPublicGroupManager( ) 
{
   m_TotGroups = 0;
   InitializeCriticalSection( &m_CritSec );
}


//====================================================================
// Destructor
//====================================================================
DaPublicGroupManager::~DaPublicGroupManager() 
{
      long           i, ni;
      HRESULT        res;
      DaPublicGroup   *pg;

      // delete all public groups even thought
      // they are accesses (which should never happen as
      // long as the public group handler is deleted
      // when the module is unloading
   EnterCriticalSection(&m_CritSec);

   ni = m_array.Size();
   i = 0;
   while ( i < ni ) {
      res = m_array.GetElem( i, &pg );
      if ( (SUCCEEDED(res)) && (pg != NULL) ) {
         delete pg;
      }
      i ++;
   }         

   LeaveCriticalSection(&m_CritSec);
   DeleteCriticalSection( &m_CritSec );
}



//====================================================================
// returns the handle of the added Public Group
// check for duplicate names!
//====================================================================
HRESULT DaPublicGroupManager::AddGroup( DaPublicGroup *pGroup, long *hPGHandle ) 
{
      HRESULT  res;
      long     sh;
         
   EnterCriticalSection(&m_CritSec);

   if ( pGroup->m_pOwner != NULL ) {
         // public group is already owned by some Handler
      res = E_FAIL;
      goto PGHAddGroupExit0;
   }

   res = FindGroup( pGroup->m_Name, &sh );
   if ( SUCCEEDED(res) ) {
         // there already is a public group with this name!
      res = OPC_E_DUPLICATENAME;
      goto PGHAddGroupExit0;
   }

   *hPGHandle = m_array.New();
   res = m_array.PutElem( *hPGHandle, pGroup);
   if ( FAILED(res) ) {
         // cannot add the group
      pGroup->m_pOwner = NULL;
      goto PGHAddGroupExit0;
   }
   m_TotGroups ++;
   pGroup->m_pOwner = this;
   pGroup->m_Handle = *hPGHandle;

PGHAddGroupExit0:
   LeaveCriticalSection(&m_CritSec);
   return res;
};




//====================================================================
// removes the Public Group with the given handle
// this means also deleting it!
//====================================================================
HRESULT DaPublicGroupManager::RemoveGroup( long hPGHandle ) 
{
      DaPublicGroup *pg;
      HRESULT res;

   res = S_OK;
   EnterCriticalSection(&m_CritSec);

   m_array.GetElem( hPGHandle, &pg );
   if( ( FAILED(res) ) || ( pg->Killed() == TRUE ) ) {
      res = E_INVALIDARG;
      goto PGHRemoveGroupExit1;
   }
   if ( pg->m_RefCount == 0 ) {
               // no thread is accessing the group so we can delete it
      m_array.PutElem( hPGHandle, NULL);
      delete pg;
   } else {           // kill the group as soon as possible
      pg->Kill( FALSE );
   }
   m_TotGroups --;

PGHRemoveGroupExit1:
   LeaveCriticalSection(&m_CritSec);
   return res;
};


//====================================================================
// returns the group with the given handle
// and increments ref count
//====================================================================
HRESULT DaPublicGroupManager::GetGroup( long hPGHandle, DaPublicGroup **pGroup )
{
      long res;

   EnterCriticalSection(&m_CritSec);

   res = m_array.GetElem( hPGHandle, pGroup );
   if( ( FAILED(res) ) || ( (*pGroup)->Killed() == TRUE ) ) {
      res = E_INVALIDARG;
      goto PGHGetGroupExit1;
   }
   (*pGroup)->Attach() ;
   res = S_OK;

PGHGetGroupExit1:
   LeaveCriticalSection(&m_CritSec);
   return res;
};


//====================================================================
// decrements the ref count on the group 
//====================================================================
HRESULT DaPublicGroupManager::ReleaseGroup( long hPGHandle )
{
      long           res;
      DaPublicGroup   *pGroup;

   EnterCriticalSection(&m_CritSec);

   res = m_array.GetElem( hPGHandle, &pGroup );
   if( FAILED(res) ) {
      res = E_INVALIDARG;
      goto PGHReleaseGroupExit1;
   }

   pGroup->Detach();

   res = S_OK;

PGHReleaseGroupExit1:
   LeaveCriticalSection(&m_CritSec);
   return res;
}




//====================================================================
// returns the group with the given name
//====================================================================
HRESULT DaPublicGroupManager::FindGroup( BSTR Name, long *hPGHandle )
{
      long           res;

   EnterCriticalSection(&m_CritSec);
   res = FindGroupNoLock( Name, hPGHandle );
   LeaveCriticalSection(&m_CritSec);
   return res;
};


//====================================================================
// returns the group with the given name
//====================================================================
HRESULT DaPublicGroupManager::FindGroupNoLock( BSTR Name, long *hPGHandle )
{
      long           res;
      DaPublicGroup   *pg;
      long           i, ni;

      // search all the m_array
   ni = m_array.Size();
   i = 0;
   while ( i < ni ) {
      res = m_array.GetElem( i, &pg );
      if( (SUCCEEDED(res)) 
            &&  ( pg->Killed() == FALSE ) 
            &&  ( wcscmp( pg->m_Name, Name) == 0 ) ) {
         *hPGHandle = i;
         res = S_OK;
         goto PGHFindGroupExit0;
      }
      i ++;
   }         

   res = E_FAIL;

PGHFindGroupExit0:

   return res;
};



//====================================================================
// returns the first group available
//====================================================================
HRESULT DaPublicGroupManager::First( long *hPGHandle )
{
      long           res;

   EnterCriticalSection(&m_CritSec);
   res = m_array.First( hPGHandle );
   LeaveCriticalSection(&m_CritSec);
   return res;
};


//====================================================================
// starting from the next
//====================================================================
HRESULT DaPublicGroupManager::Next( long hStartHandle, long *hNextHandle )
{
      long           res;

   if ( hStartHandle < 0 ) {
      return E_INVALIDARG;
   }

   EnterCriticalSection(&m_CritSec);
   res = m_array.Next( hStartHandle, hNextHandle );
   LeaveCriticalSection(&m_CritSec);
   return res;
};



//====================================================================
// the number of public groups
//====================================================================
HRESULT DaPublicGroupManager::TotGroups( long *nrGroups )
{
   *nrGroups = m_TotGroups;
   return S_OK;
}
//DOM-IGNORE-END