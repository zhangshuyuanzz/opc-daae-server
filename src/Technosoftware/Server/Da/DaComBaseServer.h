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

#ifndef __DACOMBASESERVER_H_
#define __DACOMBASESERVER_H_

 //DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "DaServerInstanceHandle.h"   // type SRVINSTHANDLE
#include "DaBaseServer.h"
#include "OpcCommon.h"
#include "DaBrowse.h"

class DaGenericServer;


class DaComBaseServer :
    public CComObjectRootEx<CComMultiThreadModel>,
    public IDispatchImpl<   IOPCServerDisp,
    &IID_IOPCServerDisp,
    &LIBID_OPCSDKLib>,
    public IDispatchImpl<   IOPCServerPublicGroupsDisp,
    &IID_IOPCServerPublicGroupsDisp,
    &LIBID_OPCSDKLib>,
    public IDispatchImpl<   IOPCBrowseServerAddressSpaceDisp,
    &IID_IOPCBrowseServerAddressSpaceDisp,
    &LIBID_OPCSDKLib>,
    public IOPCServer,
    public IOPCServerPublicGroups,
    public DaBrowse,
    public IPersistFile,

    // OPC 2.0 Custom Interfaces
    public OpcCommon,
    public IOPCItemProperties,

    // OPC 3.0 Custom Interfaces
    public IOPCItemIO

{
public:


    // CONSTRUCTOR
    DaComBaseServer();

    // DESTRUCTOR
    virtual ~DaComBaseServer();

    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    //+   IOPCServerDisp
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

       ////[propget]
    STDMETHODIMP get_Count(
        /*[out, retval]*/ long* pCount
        );

    ////[ propget, restricted, id( DISPID_NEWENUM ) ]
    STDMETHODIMP get__NewEnum(
        /*[out, retval]*/ IUnknown** ppUnk
        );

    ////[ propget ]
    STDMETHODIMP get_StartTime(
        /*[out, retval]*/ DATE* pStartTime
        );

    ////[ propget ]
    STDMETHODIMP get_CurrentTime(
        /*[out, retval]*/ DATE* pCurrentTime
        );

    ////[ propget ]
    STDMETHODIMP get_LastUpdateTime(
        /*[out, retval]*/ DATE* pLastUpdateTime
        );

    ////[ propget ]
    STDMETHODIMP get_MajorVersion(
        /*[out, retval]*/ short* pMajorVersion
        );

    ////[ propget ]
    STDMETHODIMP get_MinorVersion(
        /*[out, retval]*/ short* pMinorVersion
        );

    ////[ propget ]
    STDMETHODIMP get_BuildNumber(
        /*[out, retval]*/ short* pBuildNumber
        );

    ////[ propget ]
    STDMETHODIMP get_VendorInfo(
        /*[out, retval]*/ BSTR* pVendorInfo
        );

    //
    // Methods
    //

    HRESULT STDMETHODCALLTYPE Item(
        /*[in]         */ VARIANT      Item,
        /*[out, retval]*/ IDispatch ** ppDisp
        );

    HRESULT STDMETHODCALLTYPE AddGroup(
        /*[in]          */ BSTR         Name,
        /*[in]          */ VARIANT_BOOL      Active,
        /*[in]          */ long         RequestedUpdateRate,
        /*[in]          */ long         ClientGroupHandle,
        /*[in]          */ float       * pPercentDeadband,
        /*[in]          */ long         LCID,
        /*[out]         */ long       * pServerGroupHandle,
        /*[out]         */ long       * pRevisedUpdateRate,
        /*[in, optional]*/ VARIANT    * pTimeBias,
        /*[out, retval] */ IDispatch ** ppDisp
        );

    HRESULT STDMETHODCALLTYPE GetErrorString(
        /*[in]         */ long   Error,
        /*[in]         */ long   Locale,
        /*[out, retval]*/ BSTR * ErrorString
        );

    HRESULT STDMETHODCALLTYPE GetGroupByName(
        /*[in]         */ BSTR         Name,
        /*[out, retval]*/ IDispatch ** ppDisp
        );


    HRESULT STDMETHODCALLTYPE RemoveGroup(
        /*[in]*/ long    ServerGroupHandle,
        /*[in]*/ VARIANT_BOOL Force
        );

    HRESULT STDMETHODCALLTYPE SaveConfig(
        /*[in]*/ BSTR FileName
        );

    HRESULT STDMETHODCALLTYPE LoadConfig(
        /*[in]*/ BSTR FileName
        );

    HRESULT STDMETHODCALLTYPE SetEnumeratorType(
        /*[in]*/ long  Scope,
        /*[in]*/ short Type
        );

    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    //+  IOPCServerPublicGroupsDisp
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    HRESULT STDMETHODCALLTYPE GetPublicGroupByName(
        /*[in]         */ BSTR         Name,
        /*[out, retval]*/ IDispatch ** ppDisp
        );


    HRESULT STDMETHODCALLTYPE RemovePublicGroup(
        /*[in]*/ long    ServerGroupHandle,
        /*[in]*/ VARIANT_BOOL Force
        );


    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    //+  IOPCBrowseServerAddressSpaceDisp
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

       //[ propget, restricted, id( DISPID_NEWENUM )   ]
    HRESULT STDMETHODCALLTYPE get__NewEnum_BROWSE(
        /*[out, retval]*/ IUnknown ** ppUnk
        );


    ////[ propget ]
    HRESULT STDMETHODCALLTYPE get_Organization(
        /*[out, retval]*/ long * pNameSpaceType
        );


    //
    // Methods
    //

    HRESULT STDMETHODCALLTYPE ChangeBrowsePosition(
        /*[in]*/ long BrowseDirection,
        /*[in]*/ BSTR Position
        );


    HRESULT STDMETHODCALLTYPE SetItemIDEnumerator(
        /*[in]*/          long         BrowseFilterType,
        /*[in]*/          BSTR         FilterCriteria,
        /*[in]*/          VARIANT      DataTypeFilter,
        /*[in]*/          long         AccessRightsFilter
        );


    HRESULT STDMETHODCALLTYPE GetItemIDString(
        /*[in]         */ BSTR   ItemDataID,
        /*[out, retval]*/ BSTR * ItemID
        );


    HRESULT STDMETHODCALLTYPE SetAccessPathEnumerator(
        /*[in]*/          BSTR         ItemID
        );



    // ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
    // [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCServer [[[[[[[[[[[[[[[[[[[[[[[[[[[[[
    // ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]

    HRESULT STDMETHODCALLTYPE AddGroup(
        /*[in, string]       */ LPCWSTR     szName,
        /*[in]               */ BOOL        bActive,
        /*[in]               */ DWORD       dwRequestedUpdateRate,
        /*[in]               */ OPCHANDLE   hClientGroup,
        /*[unique, in]       */ LONG      * pTimeBias,
        /*[unique, in]       */ float     * pPercentDeadband,
        /*[in]               */ DWORD       dwLCID,
        /*[out]              */ OPCHANDLE * phServerGroup,
        /*[out]              */ DWORD     * pRevisedUpdateRate,
        /*[in]               */ REFIID      riid,
        /*[out, iid_is(riid)]*/ LPUNKNOWN * ppUnk
        );

    HRESULT STDMETHODCALLTYPE GetErrorString(
        /*[in]         */ HRESULT  dwError,
        /*[in]         */ LCID     dwLocale,
        /*[out, string]*/ LPWSTR * ppString
        );

    HRESULT STDMETHODCALLTYPE GetGroupByName(
        /*[in, string]       */ LPCWSTR szName,
        /*[in]               */ REFIID riid,
        /*[out, iid_is(riid)]*/ LPUNKNOWN * ppUnk
        );


    STDMETHODIMP GetStatus(
        /*[out]*/ OPCSERVERSTATUS ** ppServerStatus
        );


    HRESULT STDMETHODCALLTYPE RemoveGroup(
        /*[in]*/ OPCHANDLE hServerGroup,
        /*[in]*/ BOOL      bForce
        );


    HRESULT STDMETHODCALLTYPE CreateGroupEnumerator(
        /*[in]             */   OPCENUMSCOPE dwScope,
        /*[in]             */   REFIID       riid,
        /*[out, iid_is(riid)]*/ LPUNKNOWN   *ppUnk
        );


    // ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
    // [[[[[[[[[[[[[[[[[[[[[[[[[ IOPCServerPublicGroups [[[[[[[[[[[[[[[[[[[[[[[
    // ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]


    HRESULT STDMETHODCALLTYPE GetPublicGroupByName(
        /*[in, string]       */ LPCWSTR     szName,
        /*[in]               */ REFIID      riid,
        /*[out, iid_is(riid)]*/ LPUNKNOWN * ppUnk
        );

    HRESULT STDMETHODCALLTYPE RemovePublicGroup(
        /*[in]*/ OPCHANDLE hServerGroup,
        /*[in]*/ BOOL      bForce
        );


    // ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
    // [[[[[[[[[[[[[[[[[[[[[[ IOPCBrowseServerAddressSpace [[[[[[[[[[[[[[[[[[[[
    // ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
    //
    // all functions of this interface are provided by class DaBrowse.
    //


    // ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
    // [[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCPersistFile [[[[[[[[[[[[[[[[[[[[[[[[[[[[
    // ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]


    STDMETHODIMP IsDirty(void);

    STDMETHODIMP Load(
        /* [in] */                 LPCOLESTR         pszFileName,
        /* [in] */                 DWORD             dwMode
        );

    STDMETHODIMP Save(
        /* [in, unique] */         LPCOLESTR         pszFileName,
        /* [in]         */         BOOL              fRemember
        );

    STDMETHODIMP SaveCompleted(
        /* [in, unique] */         LPCOLESTR         pszFileName
        );

    STDMETHODIMP GetCurFile(
        /* [out] */                LPOLESTR       *  ppszFileName
        );

    STDMETHODIMP GetClassID(
        /* [out] */                CLSID          *  pClassID
        );


    // ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
    // [[[[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCCommon [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
    // ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
    //
    // all functions of this interface are provided by class OpcCommon.
    //


    // ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
    // [[[[[[[[[[[[[[[[[[[[[[[[[ IOPCItemProperties [[[[[[[[[[[[[[[[[[[[[[[[[[[
    // ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]

    STDMETHODIMP QueryAvailableProperties(
        /* [in] */                             LPWSTR      szItemID,
        /* [out] */                            DWORD    *  pdwCount,
        /* [size_is][size_is][out] */          DWORD    ** ppPropertyIDs,
        /* [size_is][size_is][out] */          LPWSTR   ** ppDescriptions,
        /* [size_is][size_is][out] */          VARTYPE  ** ppvtDataTypes);

    STDMETHODIMP GetItemProperties(
        /* [in] */                             LPWSTR      szItemID,
        /* [in] */                             DWORD       dwCount,
        /* [size_is][in] */                    DWORD    *  pdwPropertyIDs,
        /* [size_is][size_is][out] */          VARIANT  ** ppvData,
        /* [size_is][size_is][out] */          HRESULT  ** ppErrors);

    STDMETHODIMP LookupItemIDs(
        /* [in] */                             LPWSTR      szItemID,
        /* [in] */                             DWORD       dwCount,
        /* [size_is][in] */                    DWORD    *  pdwPropertyIDs,
        /* [size_is][size_is][string][out] */  LPWSTR   ** ppszNewItemIDs,
        /* [size_is][size_is][out] */          HRESULT  ** ppErrors);


    //=========================================================================
    // OPC 3.0 Custom Interfaces
    //=========================================================================

    ///////////////////////////////////////////////////////////////////////////
    //////////////////////////// IOPCBrowse ///////////////////////////////////
    ///////////////////////////////////////////////////////////////////////////
    //
    // all functions of this interface are provided by class DaBrowse.
    //



    ///////////////////////////////////////////////////////////////////////////
    //////////////////////////// IOPCItemIO ///////////////////////////////////
    ///////////////////////////////////////////////////////////////////////////

    STDMETHODIMP Read(
        /* [in] */                    DWORD             dwCount,
        /* [size_is][in] */           LPCWSTR        *  pszItemIDs,
        /* [size_is][in] */           DWORD          *  pdwMaxAge,
        /* [size_is][size_is][out] */ VARIANT        ** ppvValues,
        /* [size_is][size_is][out] */ WORD           ** ppwQualities,
        /* [size_is][size_is][out] */ FILETIME       ** ppftTimeStamps,
        /* [size_is][size_is][out] */ HRESULT        ** ppErrors
        );

    STDMETHODIMP WriteVQT(
        /* [in] */                    DWORD             dwCount,
        /* [size_is][in] */           LPCWSTR        *  pszItemIDs,
        /* [size_is][in] */           OPCITEMVQT     *  pItemVQT,
        /* [size_is][size_is][out] */ HRESULT        ** ppErrors
        );

    //======================================================================

 //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 //+  ...
 //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

public:

    /**
     * @fn  HRESULT FinalConstructBase();
     *
     * @brief   Final construct base.
     *
     * @return  A hResult.
     */

    HRESULT  FinalConstructBase();

    /**
     * @fn  void FinalReleaseBase();
     *
     * @brief   Final release base.
     */

    void     FinalReleaseBase();

    /**
     * @fn  virtual void FireShutdownRequest(LPCWSTR szReason) = 0;
     *
     * @brief   Fires a 'Shutdown Request' to the subscribed client
     *          Pure virtual function, must be implemented by derived classes.
     *
     * @param   reason      The reason.
     */

    virtual void FireShutdownRequest(LPCWSTR reason) = 0;


public:

    /** @brief    if TRUE no server can will created. */
    BOOL creationFailed_;

    // This pointer must be set by a COM Server Class inheriting from this class.
    // It contains all the data shared by all server instances
    // of that Server Class
    DaBaseServer *daBaseServer_;


    /** @brief   the generic server attached to this COM server. */
    DaGenericServer *daGenericServer_;

    /**
     * @brief   Server Instance Handle. Servers created by diffrent clients has diffrent handles.
     *          Note :   This handle is set in the DaComServer constructor.
     *                   Do never change this handle after creation.
     */

    SRVINSTHANDLE  serverInstanceHandle_;
};
//DOM-IGNORE-END

#endif // __DACOMBASESERVER_H_
