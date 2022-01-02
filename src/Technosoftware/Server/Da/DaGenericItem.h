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

#ifndef __GENERICITEM_H_
#define __GENERICITEM_H_

//DOM-IGNORE-BEGIN

#include "DaDeviceItem.h"


/////////////////////////////////////////////////////////////////
// Item Definition Extension
// Structure to define additional data for a generic item.
/////////////////////////////////////////////////////////////////
typedef struct tagITEMDEFEXT {

   BOOL  m_fPhyvalItem;
   ///
   ///

} ITEMDEFEXT;




class DaGenericGroup;

class DaGenericItem {
public:

      // ===============================================================
      //             Management Functions
      // ===============================================================

      //--------------------------------------------------------------
      // Constructor
      //--------------------------------------------------------------
   DaGenericItem( void );

      //--------------------------------------------------------------
      // Create
      // ------
      // Initializes an item
      //--------------------------------------------------------------
   HRESULT Create(   BOOL           Active,
                     OPCHANDLE      Client,
                     VARTYPE        ReqDataType,
                     DaGenericGroup  *pGroup,
                     DaDeviceItem    *pDeviceItem,
                     long           *pServerHandle,
                     BOOL            fPhyValItem );
      
      //--------------------------------------------------------------
      // Create a Clone
      // --------------
      // Initializes an item copying from given item
      //--------------------------------------------------------------
   HRESULT Create(   DaGenericGroup  *pGroup,
                     DaGenericItem   *pCloned,
                     long           *pServerHandle );

      //--------------------------------------------------------------
      //  Destructor
      //--------------------------------------------------------------
   ~DaGenericItem();


      //--------------------------------------------------------------
      //  Attach to class by incrementing the RefCount
      //  Returns the current RefCount
      //--------------------------------------------------------------
   int  Attach( void );

      //--------------------------------------------------------------
      //  Detach from class by decrementing the RefCount.
      //  Kill the class instance if request pending.
      //  Returns the current RefCount or -1 if it was killed.
      //--------------------------------------------------------------
   int  Detach( void );


      //--------------------------------------------------------------
      //  Kill the class instance.
      //  If the RefCount is > 0 then only the kill request flag is set.
      //  Returns the current RefCount or -1 if it was killed.
      //--------------------------------------------------------------
   int  Kill( void );

      //--------------------------------------------------------------
      // Tells whether the class instance was killed or not
      //--------------------------------------------------------------
   BOOL Killed( void );



      //--------------------------------------------------------------
      //  Attach to connected DeviceItem class by incrementing the RefCount
      //  Returns the current RefCount and sets the class pointer variable.
      //--------------------------------------------------------------
   int  AttachDeviceItem( DaDeviceItem **DItem ) ;
   

      //--------------------------------------------------------------
      //  Detach from class by decrementing the RefCount.
      //  Kill the class instance if request pending.
      //  Returns the current RefCount or -1 if it was killed.
      //--------------------------------------------------------------
   int  DetachDeviceItem( void ) ;


   
   
      // ===============================================================
      // function to get/change members and members 
      // ===============================================================

            //  Get/Set additional data of a generic item
   HRESULT get_ExtItemDef( ITEMDEFEXT * pExtItemDef );
   HRESULT set_ExtItemDef( ITEMDEFEXT * pExtItemDef );

            //  Get/Set Active Flag
   BOOL    get_Active( void ) ;
   HRESULT set_Active( BOOL Active ) ;

            // Get/Set the data type requestd by the client
   VARTYPE get_RequestedDataType( void ) ;
   HRESULT set_RequestedDataType( VARTYPE RequestedDataType ) ;

            // Get/Set the client handle
   unsigned long get_ClientHandle( void ) ;
   void          set_ClientHandle( unsigned long ClientHandle ) ;


      // ===============================================================
      // function to get/change members and members of the
      // connected DaDeviceItem class
      // ===============================================================
   HRESULT  get_ItemIDPtr( LPWSTR *ItemID ) ;
   HRESULT  get_ItemIDCopy( LPWSTR *ItemID ) ;
   HRESULT  set_ItemID( LPWSTR ItemID ) ;

   HRESULT  get_AccessPath( LPWSTR *AccessPath );
   HRESULT  set_AccessPath( LPWSTR AccessPath );
   
   HRESULT  get_Blob( BYTE **pBlob, DWORD *BlobSize );
   HRESULT  set_Blob( BYTE *pBlob, DWORD BlobSize );
   
   VARTYPE  get_CanonicalDataType( void ) ;
   
   HRESULT  get_AccessRights( DWORD *DaAccessRights ) ;
   
   HRESULT  get_EUData( OPCEUTYPE *EUType, VARIANT *EUInfo ) ;

                  // Get all Item Attributes
   HRESULT  get_Attr( OPCITEMATTRIBUTES * pAttr, OPCHANDLE hServer );

                  // Modify the Active Counter of the attached DeviceItem.
   HRESULT  AttachActiveCountOfDeviceItem( void );
   HRESULT  DetachActiveCountOfDeviceItem( void );

                  // Functions to handle the members m_LastReadValue and m_LastReadQuality
   HRESULT  CompareLastRead( float fltPercentDeadband,
                             const VARIANT& vCompValue, WORD wCompQuality,
                             BOOL& fChanged );
   HRESULT  UpdateLastRead( VARIANT vValue, WORD wQuality );
   void     ResetLastRead( void );


//=============================  Member Variables  ==================================
protected:
                  // tells if instance was successfully created
                  // ( is FALSE if Create was not called or failed )
   BOOL           m_Created;

                  // these values are used by the group
                  // for synchronisation and lazy removal
   BOOL           m_ToKill;

                  // Last value/quality read by the client.
                  // This is used by the update thread to find if the item has changed and
                  // must be transmitted to the client.
   VARIANT        m_LastReadValue;
   WORD           m_LastReadQuality;

                  // the group owning the generic item
   DaGenericGroup  *m_pGroup;

                  // the server handle assigned to the item in the group
   OPCHANDLE      m_ServerHandle;

                  // Defines if the group is active for periodic update of the client
   BOOL           m_Active ;


                  // Number of connected clients
   long           m_RefCount;

                  // Defines additional item definitions
                  // e.g for CALL-R items
   ITEMDEFEXT     m_ExtItemDef;

                  // Client Handle 
                  // is returned to client to help identify the item
                  // ( specially in async functions )
   unsigned long  m_ClientHandle ;

   //SM should be    // access path: the access path associated with this instance
   //SM should be LPWSTR         m_AccessPath;

                  // Requested data type
   VARTYPE        m_RequestedDataType;

                  // Connection to the Device Item instance
   DaDeviceItem *  m_DeviceItem;

   CRITICAL_SECTION m_CritSec;

private:
};
//DOM-IGNORE-END


#endif // __GENERICITEM_H_




