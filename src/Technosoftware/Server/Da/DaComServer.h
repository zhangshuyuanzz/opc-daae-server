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

#ifndef __OPCSERVER_H_
#define __OPCSERVER_H_

 //DOM-IGNORE-BEGIN

#include  <atlbase.h>

#include "DaComBaseServer.h"
#include "ShutdownImpl.h"

/**
 * @class   ATL_NO_VTABLE
 *
 * @brief   Implementation of DaComServer
 *          When a client connects (with CreateObject) to the Server
 *          then an instance of this class is created.
 *          The implementation of the interfaces is in the superclass
 *          DaComBaseServer. The reason for this migration is the
 *          possibility to create more servers in the same EXE
 *          ( i.e. ex. for sharing resources ) .
 *          Different COM server classes will have a different
 *          ServerClassHandler;
 *          this is set in the constructor of the class
 */

class ATL_NO_VTABLE DaComServer :
    public CComCoClass<DaComServer, &CLSID_OPCServer>,
    public DaComBaseServer,
    public IConnectionPointContainerImpl<DaComServer>,
    public IOPCShutdownConnectionPointImpl<DaComServer>
{
public:

    /**
     * @fn  DaComServer();
     *
     * @brief   Default constructor.
     */

    DaComServer();

    DECLARE_NO_REGISTRY()
    DECLARE_GET_CONTROLLING_UNKNOWN()


    BEGIN_COM_MAP(DaComServer)
        COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)

        COM_INTERFACE_ENTRY(IOPCServerDisp)
        COM_INTERFACE_ENTRY(IOPCServerPublicGroupsDisp)
        COM_INTERFACE_ENTRY(IOPCBrowseServerAddressSpaceDisp)

        COM_INTERFACE_ENTRY(IOPCServer)
        COM_INTERFACE_ENTRY(IOPCServerPublicGroups)
        COM_INTERFACE_ENTRY(IOPCBrowseServerAddressSpace)

        COM_INTERFACE_ENTRY(IPersistFile)

        // OPC 2.0 Custom Interfaces
        COM_INTERFACE_ENTRY(IOPCCommon)
        COM_INTERFACE_ENTRY(IOPCItemProperties)

        // OPC 3.0 Custom Interfaces
        COM_INTERFACE_ENTRY(IOPCBrowse)
        COM_INTERFACE_ENTRY(IOPCItemIO)

        // OPC 1.0A Automation Interfaces
        COM_INTERFACE_ENTRY2(IDispatch, IOPCServerDisp)

        // !!!!!!!!!!!!!!!! MAYBE BETTER NO AGGREGATION!
        COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, unknownMarshaler_.p)
    END_COM_MAP()

    BEGIN_CONNECTION_POINT_MAP(DaComServer)
        CONNECTION_POINT_ENTRY(IID_IOPCShutdown)
    END_CONNECTION_POINT_MAP()

    /**
     * @fn  HRESULT FinalConstruct()
     *
     * @brief   Final construct.
     *
     * @return  A hResult.
     */

    HRESULT FinalConstruct()
    {
        HRESULT hres = FinalConstructBase();
        if (SUCCEEDED(hres)) {
            hres = CoCreateFreeThreadedMarshaler(GetControllingUnknown(),
                &unknownMarshaler_.p);
        }
        return hres;
    }

    /**
     * @fn  void FinalRelease()
     *
     * @brief   Final release.
     */

    void FinalRelease()
    {
        FinalReleaseBase();
        unknownMarshaler_.Release();
    }

    CComPtr<IUnknown> unknownMarshaler_;

    /**
     * @fn  static HRESULT CreateInstanceInternal( DaComBaseServer * creator, DaBaseServer * daBaseServerToConnect, REFIID refIid, LPUNKNOWN * unknown);
     *
     * @brief   This Method is not used in OPC Servers. Only OpenControl CALL-R Servers can
     *          dynamically create server instances. Creates another Server Instance and connect it
     *          to the specified Server Handler.
     *          
     * @param [in,out]  creator                 Creator of the new Server Instance. Used to set the
     *                                          Server Instance Handle of the new created Server.
     * @param [in,out]  daBaseServerToConnect   Specifiy the Server Class Handler
     *                                          to connect the new created Server Instance.
     * @param   refIid                          The interface requested from the new created server.
     * @param [in,out]  unknown                 Where to return the interface.
     *
     * @return  A hResult.
     */

    static HRESULT CreateInstanceInternal(
        /* [in]          */ DaComBaseServer *   creator,
        /* [in]          */ DaBaseServer  *     daBaseServerToConnect,
        /* [in]          */ REFIID              refIid,
        /* [iid_is][out] */ LPUNKNOWN *         unknown);

    /**
     * @fn  inline void FireShutdownRequest(LPCWSTR szReason)
     *
     * @brief   Fire shutdown request.
     *
     * @param   reason      The reason.
     */

    inline void FireShutdownRequest(LPCWSTR reason)
    {
        IOPCShutdownConnectionPointImpl<DaComServer>::FireShutdownRequest(reason);
    }

    // Implementation
private:  
    /** @brief   Critical Section to protect 'creator_' and 'daBaseServerToConnect_' */
    static _ThreadModel::AutoCriticalSection    criticalSectionCreate_;

    /**
     * @brief   Creator of the server instance. NULL means by (D)COM (default). 
     *          Otherwise the Server instance is created with CreateInstanceInternal(...)
     */

    static DaComBaseServer   *                  creator_;

    /**
     * @brief   Server instances created with CreateInstanceInternal(...) are connected to  
     *          the Server Handler specified by this member.  
     */

    static DaBaseServer *                       daBaseServerToConnect_;
};
//DOM-IGNORE-END

#endif //__OPCSERVER_H_
