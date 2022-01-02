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
#include "DaPublicGroup.h"
#include "DaPublicGroupManager.h"
#include "UtilityFuncs.h"
#include "Logger.h"

/********************************************************************/
//////////////////////////////////////////////////////////////////////
// PUBLIC GROUP 
//////////////////////////////////////////////////////////////////////
/********************************************************************/


HRESULT DaPublicGroup::CommonInit(
            LPWSTR      Name,
            float       PercentDeadband,
            long        LCID,
            BOOL        InitialActive,
            long        *pInitialTimeBias,
            long        InitialUpdateRate
         )
{
   if (PercentDeadband < 0.0 || PercentDeadband > 100.0) {
      LOGFMTE( "CommonInit() failed with invalid argument(s): Percent Deadband out of range (0...100)" );
      return  E_INVALIDARG;
   }

   m_Name = WSTRClone( Name, NULL );
   if (!m_Name) {
      return E_OUTOFMEMORY;
   }

   m_PercentDeadband = PercentDeadband;
   m_LCID              = LCID;
   m_InitialActive     = InitialActive;
   m_InitialUpdateRate = InitialUpdateRate;

   if (pInitialTimeBias == NULL) {              // use default Time Bias

      TIME_ZONE_INFORMATION tzi;
      if (GetTimeZoneInformation( &tzi ) == TIME_ZONE_ID_INVALID) {
         return HRESULT_FROM_WIN32( GetLastError() );
      }
      else {
         m_InitialTimeBias = tzi.Bias;
      }
   }
   else {                                       // use specified Time Bias
      m_InitialTimeBias = *pInitialTimeBias;
   }
   return S_OK;
}


//====================================================================
// DaPublicGroup constructor
//====================================================================
DaPublicGroup::DaPublicGroup( void )
{
   m_Created      = FALSE;
   m_pOwner       = NULL;
   m_Name         = NULL;
   m_TotItems     = 0;
   m_AllocItems   = 0;
   m_ExtItemDefs  = NULL;
   m_ItemDefs     = NULL;
   m_ToKill       = FALSE;                      // used for lazy deletion of public group
   m_RefCount     = 0;
   InitializeCriticalSection( &m_CritSec );
}


//====================================================================
// DaPublicGroup creator
//====================================================================
HRESULT DaPublicGroup::Create(
                        LPWSTR      Name,
                        float       PercentDeadband,
                        long        LCID,
                        BOOL        InitialActive,
                        long        *pInitialTimeBias,
                        long        InitialUpdateRate,
                        DWORD       NumItems,
                        OPCITEMDEF  *ItemDefs,
                        ITEMDEFEXT  *ExtItemDefs /*= NULL */
                  )
{
      DWORD i;
      HRESULT res;

   if (m_Created == TRUE) {                     // instance already initialized   
      return E_FAIL;
   }
   res = CommonInit(
            Name,
            PercentDeadband,
            LCID,
            InitialActive,
            pInitialTimeBias,
            InitialUpdateRate
          );

   if (FAILED( res ))
      return res;

   m_TotItems = NumItems;
   m_AllocItems = NumItems;

   if ( NumItems > 0 ) {
         // copy item definitions locally!
      m_ItemDefs = new OPCITEMDEF[ NumItems ];
      if ( m_ItemDefs == NULL ) {
         res = E_OUTOFMEMORY;
         goto PublicGroupCreateWithItemsExit0;
      }

         // make a deep copy of the item definitions!
      i = 0;
      while ( i < NumItems ) {
         res = CopyOPCITEMDEF( &(m_ItemDefs[i]), &(ItemDefs[i]) );
         if ( FAILED(res) ) {
            goto PublicGroupCreateWithItemsExit1;
         }
         i ++;
      }

         // copy additional item definitions locally!
      if (ExtItemDefs) {      // only if exist
         m_ExtItemDefs = new ITEMDEFEXT[ NumItems ];
         if ( m_ExtItemDefs == NULL ) {
            res = E_OUTOFMEMORY;
            goto PublicGroupCreateWithItemsExit1;
         }
         // make a copy of the additional item definitions!
         memcpy( m_ExtItemDefs, ExtItemDefs, NumItems * sizeof (ITEMDEFEXT) );
      }

   } else {
      m_ItemDefs     = NULL;
      m_ExtItemDefs  = NULL;
   }

   m_Created = TRUE;
   return S_OK;

PublicGroupCreateWithItemsExit1:
   i --;
   while ( i >= 0 ) {
      FreeOPCITEMDEF( &(m_ItemDefs[i]) );
      i --;
   }
   delete [] m_ItemDefs;
   m_ItemDefs     = NULL;
   m_ExtItemDefs  = NULL;

PublicGroupCreateWithItemsExit0:
   return res;
}


//====================================================================
// DaPublicGroup constructor with no items
//====================================================================
HRESULT DaPublicGroup::Create( 
                        LPWSTR      Name,
                        float       PercentDeadband,
                        long        LCID,
                        BOOL        InitialActive,
                        long        *pInitialTimeBias,
                        long        InitialUpdateRate
                  )
{
   HRESULT res;

   if (m_Created == TRUE) {                     // instance already initialized
      return E_FAIL;
   }
   res = CommonInit(
            Name,
            PercentDeadband,
            LCID,
            InitialActive,
            pInitialTimeBias,
            InitialUpdateRate
          );

   if (FAILED( res ))
      return res;

   m_Created = TRUE;
   return S_OK;
}


//====================================================================
// DaPublicGroup destructor
//====================================================================
DaPublicGroup::~DaPublicGroup()
{
   if (m_pOwner) {
      m_pOwner->m_array.PutElem( m_Handle, NULL );
   }

   if (m_ItemDefs) {
      for (DWORD i=0; i < m_TotItems; i++) {
         FreeOPCITEMDEF( &m_ItemDefs[i] );
      }
      delete [] m_ItemDefs;
   }

   if (m_ExtItemDefs) {
      delete [] m_ExtItemDefs;
   }

   if (m_Name) {
      delete m_Name;
   }

   DeleteCriticalSection( &m_CritSec );
}


//====================================================================
// Add an item def to the group
//====================================================================
HRESULT DaPublicGroup::AddItem( 
         LPWSTR      ItemID,
         LPWSTR      InitialAccessPath,
         BOOL        InitialActive,
         OPCHANDLE   InitialClientHandle,
         DWORD       InitialBlobSize,
         BYTE       *pInitialBlob,
         VARTYPE     InitialRequestedDataType
      )
{
      DWORD       tot;
      DWORD       alloc;
      OPCITEMDEF  *itemDefinitions;
      DWORD       i;
      HRESULT     res;

   _ASSERTE( m_Created == TRUE );

   EnterCriticalSection( &m_CritSec );

   if ( m_pOwner != NULL ) {
         // group already added to a handler!
      res = E_FAIL;
      goto CopyOPCITEMDEFExit2;
   }

   if ( (InitialBlobSize < 0) || (ItemID == NULL) || (wcscmp(ItemID,L"") == 0) ) {
      res = E_INVALIDARG;
      goto CopyOPCITEMDEFExit2;
   }

   tot = m_TotItems;
      // if there is no more place in the actually allocated array
      // then create a greater one 
   if ( tot + 1 > m_AllocItems ) {
         // calc new allloc size
      if ( m_AllocItems < 4 ) {
         alloc = 4;
      } else {
         alloc = m_AllocItems * 2;
      }
         // allocate new array
      itemDefinitions = new OPCITEMDEF[ alloc ];
      if ( itemDefinitions == NULL ) {
         res = E_OUTOFMEMORY;
         goto CopyOPCITEMDEFExit0;
      }
      memset( itemDefinitions, 0, alloc * sizeof( OPCITEMDEF ) );

         // copy old to new array
      i = 0;
      while ( i < m_AllocItems ) {
         res = CopyOPCITEMDEF( &(itemDefinitions[i]), &(m_ItemDefs[i]) );
         if ( res != S_OK ) {
            goto CopyOPCITEMDEFExit1;
         }
         i ++;
      }

         // free old array
      i = 0;
      while ( i < m_AllocItems ) {
         FreeOPCITEMDEF( &(m_ItemDefs[i]) );
         i ++;
      }
      delete m_ItemDefs;

      m_ItemDefs = itemDefinitions;
      m_AllocItems = alloc;
   }

      // add the new item
   m_ItemDefs[ tot ].szItemID = WSTRClone( ItemID, NULL  );
   if ( m_ItemDefs[ tot ].szItemID == NULL ) {
      res = E_OUTOFMEMORY;
      goto CopyOPCITEMDEFExit2;
   }
   m_ItemDefs[ tot ].szAccessPath = WSTRClone( InitialAccessPath, NULL );
   if ( m_ItemDefs[ tot ].szAccessPath == NULL ) {
      res = E_OUTOFMEMORY;
      goto CopyOPCITEMDEFExit3;
   }
   
   m_ItemDefs[ tot ].bActive = InitialActive;
   m_ItemDefs[ tot ].hClient = InitialClientHandle;      // Version 1.3, AM

   m_ItemDefs[ tot ].dwBlobSize = InitialBlobSize;
   if ( pInitialBlob != NULL ) {
      m_ItemDefs[ tot ].pBlob = new BYTE[ InitialBlobSize ];
      if (  m_ItemDefs[ tot ].pBlob == NULL ) {
         res = E_OUTOFMEMORY;
         goto CopyOPCITEMDEFExit4;
      }
      memcpy( m_ItemDefs[ tot ].pBlob, pInitialBlob, InitialBlobSize );
   } else {
      m_ItemDefs[ tot ].pBlob = NULL;
   }

   m_ItemDefs[ tot ].vtRequestedDataType = InitialRequestedDataType;

   m_TotItems = tot + 1;

   LeaveCriticalSection( &m_CritSec );

   return S_OK;

CopyOPCITEMDEFExit4:
   WSTRFree( m_ItemDefs[ tot ].szAccessPath, NULL );

CopyOPCITEMDEFExit3:
   WSTRClone( m_ItemDefs[ tot ].szItemID, NULL);

CopyOPCITEMDEFExit2:
   goto CopyOPCITEMDEFExit0;

CopyOPCITEMDEFExit1:
   while ( i > 0 ) {
      i --;
      FreeOPCITEMDEF( &(itemDefinitions[i]) );
   }
CopyOPCITEMDEFExit0:
   LeaveCriticalSection( &m_CritSec );
   return res;
}




//====================================================================
//  Attach to class by incrementing the RefCount
//  Returns the current RefCount
//  or -1 if the ToKill flag is set
//====================================================================
int  DaPublicGroup::Attach( void ) 
{
      long t;

   EnterCriticalSection( &m_CritSec );

   if( m_ToKill ) {
      LeaveCriticalSection( &m_CritSec );
      return -1 ;
   }

   t = ++m_RefCount;

   LeaveCriticalSection( &m_CritSec );

   return t;

}




//====================================================================
//  Detach from class by decrementing the RefCount.
//  Kill the class instance if request pending.
//  Returns the current RefCount or -1 if it was killed.
//====================================================================
int  DaPublicGroup::Detach( void )
{
      long t;

   EnterCriticalSection( &m_CritSec );

   if( m_RefCount > 0 ) {
      --m_RefCount ;
   }

   if( ( m_RefCount == 0 )                      // not referenced
         &&  ( m_ToKill == TRUE ) ){                  // kill request
      LeaveCriticalSection( &m_CritSec );
      delete this;                              // remove from memory
      return -1 ;
   }

   t = m_RefCount ;

   LeaveCriticalSection( &m_CritSec );

   return t ;
}


//=====================================================================================
//  Kill the class instance.
//  If the RefCount is > 0 then only the kill request flag is set.
//  Returns the current RefCount or -1 if it was killed.
//=====================================================================================
int  DaPublicGroup::Kill( BOOL WithDetach )
{
   EnterCriticalSection( &m_CritSec );

   if( WithDetach ) {
      if( m_RefCount > 0 )    
         --m_RefCount ;
   }

   if( m_RefCount ) {            // still referenced
      m_ToKill = TRUE ;          // set kill request flag
   } else {
      LeaveCriticalSection( &m_CritSec );
      delete this ;
      return -1 ;
   }  
   LeaveCriticalSection( &m_CritSec );

   return m_RefCount ;
}


//=====================================================================================
// Tells whether Kill() was called on this instance
//=====================================================================================
BOOL DaPublicGroup::Killed( void )
{
      BOOL b;

   EnterCriticalSection( &m_CritSec );

   b = m_ToKill;

   LeaveCriticalSection( &m_CritSec );

   return b;
}

//DOM-IGNORE-END
