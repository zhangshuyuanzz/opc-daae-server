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

#ifndef __DEVICEITEM_H_
#define __DEVICEITEM_H_


#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

class DaBaseServer;


class DaDeviceItem  {
public:
      //--------------------------------------------------------------
      // constructor
      //--------------------------------------------------------------
   DaDeviceItem();

      //--------------------------------------------------------------
      // initializer
      //--------------------------------------------------------------
   HRESULT Create(   LPWSTR      szItemID,
                     DWORD       dwAccessRights,
                     LPVARIANT   pvValue,
                     BOOL        fActive        = TRUE,
                     DWORD       dwBlobSize     = 0,
                     LPBYTE      pBlob          = NULL,
                     LPWSTR      szAccessPath   = NULL,
                     OPCEUTYPE   eEUType        = OPC_NOENUM,
                     LPVARIANT   pvEUInfo       = NULL
                  );

      //--------------------------------------------------------------
      // destructor
      //--------------------------------------------------------------
   virtual ~DaDeviceItem();
   
      //--------------------------------------------------------------
      //  Attach to class by incrementing the RefCount
      //  Returns the current RefCount
      //--------------------------------------------------------------
   virtual int  Attach( void );

      //--------------------------------------------------------------
      //  Detach from class by decrementing the RefCount.
      //  Kill the class instance if request pending.
      //  Returns the current RefCount or -1 if it was killed.
      //--------------------------------------------------------------
   virtual int  Detach( void );

      //--------------------------------------------------------------
      //  Kill the class instance.
      //  If the RefCount is > 0 then only the kill request flag is set.
      //  Returns the current RefCount or -1 if it was killed.
      //  If WithDeatch == TRUE then RefCound is first decremented.
      //--------------------------------------------------------------
   virtual int  Kill( BOOL WithDetach );


      //--------------------------------------------------------------
      // Tells whether Kill() was called on this instance
      //--------------------------------------------------------------
   virtual BOOL  Killed( void );



      //--------------------------------------------------------------
      // function to change value of members which require memory
      // allocation and deallocation
      //--------------------------------------------------------------
   virtual DWORD   get_RefCount(void);
   virtual BOOL    get_Active(void);
   virtual HRESULT set_Active( BOOL Active );
   virtual VARTYPE get_CanonicalDataType( void );
   virtual HRESULT set_CanonicalDataType( VARTYPE CanonicalDataType );

   virtual HRESULT get_AccessPath( LPWSTR *AccessPath );
   virtual HRESULT set_AccessPath( LPWSTR AccessPath );
   virtual HRESULT get_Blob( BYTE **pBlob, DWORD *BlobSize );
   virtual HRESULT set_Blob( BYTE *pBlob, DWORD BlobSize );
   virtual HRESULT get_ItemIDPtr( LPWSTR *ItemID );
   virtual HRESULT get_ItemIDCopy( LPWSTR *ItemID, IMalloc *pIMalloc = NULL );
   virtual HRESULT set_ItemID( LPWSTR ItemID );
   virtual HRESULT get_AccessRights( DWORD *pAccessRights );
   virtual HRESULT set_AccessRights( DWORD DaAccessRights );
   virtual HRESULT get_EUData( OPCEUTYPE *pEUType, VARIANT *pEUInfo );
   virtual HRESULT set_EUData( OPCEUTYPE EUType, VARIANT *pEUInfo );
   virtual HRESULT get_AnalogEURange( double* pdAnalogEURange );
   virtual HRESULT get_OPCITEMRESULT( BOOL blobUpdate, OPCITEMRESULT* pItemResult );

      //--------------------------------------------------------------
      // Write current value, quality and time stamp
      //--------------------------------------------------------------
   virtual HRESULT set_ItemValue(   LPVARIANT pvValue,
                                    WORD wQuality = OPC_QUALITY_GOOD | OPC_LIMIT_OK,
                                    LPFILETIME pftTimeStamp = NULL );
   virtual HRESULT set_ItemQuality( WORD wQuality,
                                    LPFILETIME pftTimeStamp = NULL );
   virtual HRESULT set_ItemVQT(     OPCITEMVQT* pItemVQT );


      //--------------------------------------------------------------
      // Read current value, quality and time stamp
      // with requested data type in Value.vt
      //--------------------------------------------------------------
   virtual HRESULT get_ItemValue(   LPVARIANT pvValue,
                                    LPWORD pwQuality,
                                    LPFILETIME pftTimeStamp );

      //--------------------------------------------------------------
      // Converts the specified variant value from one type to
      // the canonical data type.
      //--------------------------------------------------------------
   virtual HRESULT ChangeValueToCanonicalType( LPVARIANT pvValue );


      //--------------------------------------------------------------
      // Compares the current TimeStamp with the specified value
      //--------------------------------------------------------------
   virtual BOOL IsTimeStampOlderThan( LPFILETIME pftTimeStamp );


      //--------------------------------------------------------------
      // Returns the value of a standard item property.
      //--------------------------------------------------------------
   virtual HRESULT get_PropertyValue(  DaBaseServer* const pServerHandler,
                                       DWORD dwPropID,
                                       LPVARIANT pvPropData );

   inline BOOL  HasReadAccess()  const { return (m_AccessRights & OPC_READABLE) ? TRUE : FALSE; }
   inline BOOL  HasWriteAccess() const { return (m_AccessRights & OPC_WRITEABLE) ? TRUE : FALSE; }

      //--------------------------------------------------------------
      // Item Deadband
      //--------------------------------------------------------------
   virtual HRESULT SetItemDeadband( FLOAT fltPercentDeadband );
   virtual HRESULT GetItemDeadband( FLOAT* pfltPercentDeadband );
   virtual HRESULT ClearItemDeadband();

public:
      //--------------------------------------------------------------
      // to protect members of this class from multi thread access
      //--------------------------------------------------------------
   CRITICAL_SECTION  m_CritSec;

      //--------------------------------------------------------------
      // Used to protect all item attributes of this item.
      // Required since device item attributes can be changed by
      // Call-R functions.
      //
      // Access protection used by functions
      //    IOPCItemMgt::CreateEnumerator          (OPC)
      //    IOPCItemMgt::get_OPCITEMRESULT         (OPC)
      //    ICallrItemConfig::GetCallrItemDefs     (Call-R)
      //    ICallrItemConfig::ChangeCallrItemDefs  (Call-R)
      //
      //--------------------------------------------------------------
   CRITICAL_SECTION  m_CritSecAllAttrs;

protected:
               // zero terminated string that uniquely
               // identifies the item (UNICODE!) 
   LPWSTR      m_ItemID;

               // recommandation to the server on 'how to get the data' 
               //    ex. through which COM port 
   LPWSTR      m_AccessPath;

               // tells whether the cache for this item should be refreshed.
               // Group specific handling is controlled by the Active Flag
               // in the GenericItem class.
               // See also 'Active Count Handling' below.
   BOOL        m_Active ;

               // Number of references from GenericItems
   long        m_RefCount ;

               // Item should be killed as soon as it not referenced any more
   BOOL        m_ToKill ;

               // item access rights RO,WO,RW :   
               // OPCACCESSRIGTHS enum :   OPC_READABLE   (==1)
               //                          OPC_WRITEABLE  (==2)
               // high word available for vendor specific use.
   DWORD       m_AccessRights; 

               // Item Value Cache
   VARIANT     m_Value ;                  // data type and current value
   WORD        m_Quality ;                // OPC quality flag
   FILETIME    m_TimeStamp ;              // time when the item cache was written

               // the blob is a (zero terminated?) string 
               //    provided by the client or by the server 
               //    that should or could help the server 
               //    to speed up access to the item 
               //    "last time I looked for this Tag
               //    I found it here"
   DWORD       m_BlobSize;
   BYTE      * m_pBlob;

               // EUInfo (optional)
   OPCEUTYPE   m_EUType;
   VARIANT     m_EUInfo;
               // Contains the range of the EU Info if EU Type is Analog. For this reason
               // the range must not be calcuated at every update cycle.
   double      m_dAnalogEURange;

               // The PercentDeadband value of this item
   FLOAT       m_fltPercentDeadband;

      //--------------------------------------------------------------
      // Active Count Handling.
      //    Counts how many GenericItems with active state of an
      //    Active Group are attached to this item.
      //    This counter tells whether the cache for this item should
      //    be refreshed. It is recommended to use this counter in
      //    combination with the 'Active' flag. The Active flag can
      //    be handled from the User Specific Code and is not
      //    handled in the Generic Part in any way.
      //
      //    Toolkit Users :
      //    Don't call the function members AttachActiveCount()/
      //    DetachActiveCount() from User Specific Code. Instead use
      //    the function set_Active() to modify the active state.
      //    If you would be notified if the value of active counter
      //    changes overload the functions in your DeviceItem
      //    derived class.
      //--------------------------------------------------------------
   public:
      DWORD    get_ActiveCount( void );

      virtual  HRESULT AttachActiveCount( void );
      virtual  HRESULT DetachActiveCount( void );

   private:
               // Counts how many GenericItems with active state
               // are attached to this item.
      DWORD    m_dwActiveCount;

};

#endif // __DEVICEITEM_H_
