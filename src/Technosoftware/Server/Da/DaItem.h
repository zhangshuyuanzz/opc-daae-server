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

#ifndef __OPCITEM_H_
#define __OPCITEM_H_

//DOM-IGNORE-BEGIN

#include "DaGenericItem.h"     


class DaGenericGroup;     


/////////////////////////////////////////////////////////////////////////////
//                                                          DaItem
class ATL_NO_VTABLE DaItem : 
   public CComObjectRootEx<CComMultiThreadModel>,
   public CComCoClass<     DaItem, &CLSID_OPCItem>,
   public IDispatchImpl<   IOPCItemDisp, 
                           &IID_IOPCItemDisp, 
                           &LIBID_OPCSDKLib>
{
public:
   DaItem();
   ~DaItem(); 

DECLARE_GET_CONTROLLING_UNKNOWN()

BEGIN_COM_MAP(DaItem)
   COM_INTERFACE_ENTRY(IOPCItemDisp)
   COM_INTERFACE_ENTRY(IDispatch)
   COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
END_COM_MAP()

//+++++++++++++++++++++++++++++++++++++
//+ IOPCItemDisp
//+++++++++++++++++++++++++++++++++++++

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_AccessPath(
        /*[out, retval]*/ BSTR * pAccessPath
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_AccessRights(
        /*[out, retval]*/ long * pAccessRights
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_ActiveStatus(
//        /*[out, retval]*/ boolean * pActiveStatus
        /*[out, retval]*/ VARIANT_BOOL * pActiveStatus
      );

      ////[ propput ]
      HRESULT STDMETHODCALLTYPE put_ActiveStatus(
//        /*[in]*/ boolean ActiveStatus
        /*[in]*/ VARIANT_BOOL ActiveStatus
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_Blob(
        /*[out, retval]*/ VARIANT * pBlob
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_ClientHandle(
        /*[out, retval]*/ long * phClient
      );

      ////[ propput ]
      HRESULT STDMETHODCALLTYPE put_ClientHandle(
        /*[in]*/ long Client
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_ItemID(
        /*[out, retval]*/ BSTR * pItemID
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_ServerHandle(
        /*[out, retval]*/ long * pServer
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_RequestedDataType(
        /*[out, retval]*/ short * pRequestedDataType
      );

      ////[ propput ]
      HRESULT STDMETHODCALLTYPE put_RequestedDataType(
        /*[in]*/ short RequestedDataType
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_Value(
        /*[out, retval]*/ VARIANT * pValue
      );

      ////[ propput ]
      HRESULT STDMETHODCALLTYPE put_Value(
        /*[in]*/ VARIANT NewValue
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_Quality(
        /*[out, retval]*/ short * pQuality
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_Timestamp(
        /*[out, retval]*/ DATE * pTimeStamp
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_ReadError(
        /*[out, retval]*/ long * pError
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_WriteError(
        /*[out, retval]*/ long * pError
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_EUType(
        /*[out, retval]*/ short * pError
      );

      ////[ propget ]
      HRESULT STDMETHODCALLTYPE get_EUInfo(
        /*[out, retval]*/ VARIANT * pError
      );

      //
      // Methods
      //

      HRESULT STDMETHODCALLTYPE OPCRead(
        /*[in]           */ short     Source,
        /*[out]          */ VARIANT * pValue,
        /*[out, optional]*/ VARIANT * pQuality,
        /*[out, optional]*/ VARIANT * pTimestamp
      );

      HRESULT STDMETHODCALLTYPE OPCWrite(
        /*[in]*/ VARIANT NewValue
      );


//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+  ...
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



   HRESULT FinalConstruct()
   {
      return CoCreateFreeThreadedMarshaler(
         GetControllingUnknown(), &m_pUnkMarshaler.p);
   }

   void FinalRelease()
   {
      m_pUnkMarshaler.Release();
   }

   CComPtr<IUnknown> m_pUnkMarshaler;

// IOPCItem
public:


      // the group owning this item
   DaGenericGroup  *m_pGroup;

      // server handle
   long           m_ServerHandle;

      // the associated generic item in the case
      // the group object does no longer exist
      // but we have to "unnail" the generic item
   DaGenericItem   *m_pItem;

public:
      // initialize the COM item
   HRESULT Create( DaGenericGroup *pGroup, long ServerHandle, DaGenericItem *pItem );

      // get and lock the GenericItem associated with this group
   HRESULT GetGenericItem( DaGenericItem **ppItem );
      // unlock the GenericItem associated with this group
   HRESULT ReleaseGenericItem( );

      // specific methods
private:
   long           m_LastReadError;
   long           m_LastWriteError;
};
//DOM-IGNORE-END

#endif //__OPCITEM_H_
