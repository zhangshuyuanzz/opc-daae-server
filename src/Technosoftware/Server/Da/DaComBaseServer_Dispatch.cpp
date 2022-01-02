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

 //-----------------------------------------------------------------------
 // Notes:
 // -  NewEnum is actually implemented only for the IOPCServerDisp 
 //    interface ... the IOPCBrowseServerAddressSpaceDisp.NewEnum maps 
 //    to the same method!!! 
 //    To avoid this problem the two Methods (or only one) must be 
 //    renamed ... this implies a change in the opcauto10a.h file.
 //-----------------------------------------------------------------------

#include "stdafx.h"
#include "CoreMain.h"
#include "DaComBaseServer.h"
#include "DaBaseServer.h"
#include "enumclass.h"

#include "DaGenericGroup.h"
#include "openarray.h"
#include "UtilityFuncs.h"
#include <stdio.h>                        // for sscanf()
#include "Logger.h"

// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
// [[[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCServerDisp [[[[[[[[[[[[[[[[[[[[[[[[[[[[
// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]

//=========================================================================
// IOPCServerDisp.get_Count
//
// Return the number of Groups defined within the server's address space.
//
// [ propget ]
//=========================================================================
STDMETHODIMP DaComBaseServer::get_Count(
	/*[out, retval]*/ long* pCount)
{
	long                    gCnt, idx;
	HRESULT                 res;
	DaGenericGroup           *group;

	LOGFMTI("IOPCServerDisp::get_Count");
	if (pCount == NULL) {
		LOGFMTE("get_Count() failed with invalid argument(s):  pCount is NULL");
		return E_INVALIDARG;
	}

	// don't allow deletion or creation of groups while
	// counting
	EnterCriticalSection(&daGenericServer_->m_GroupsCritSec);

	daGenericServer_->RefreshPublicGroups();
	// go through all groups!
	gCnt = 0;
	res = daGenericServer_->m_GroupList.First(&idx);
	while (SUCCEEDED(res)) {
		daGenericServer_->m_GroupList.GetElem(idx, &group);
		if (group->Killed() == FALSE) {
			gCnt++;
		}
		res = daGenericServer_->m_GroupList.Next(idx, &idx);
	}

	*pCount = gCnt;

	LeaveCriticalSection(&daGenericServer_->m_GroupsCritSec);

	return S_OK;
}



//=========================================================================
//    //[ propget, restricted, id( DISPID_NEWENUM ) ]
// IOPCServerDisp.get__NewEnum
//
// Returns an enumerator that supports IEnumVariant for the items 
// currently in the collection. Read-only property.
//
// [ propget ]
//=========================================================================
STDMETHODIMP DaComBaseServer::get__NewEnum(
	/*[out, retval]*/ IUnknown** ppUnk)
{
	IOpcEnumVariant   *temp;
	int               GroupCount;
	HRESULT           res;
	LPUNKNOWN         *GroupList;
	IID               riid;
	short             varEnumType;

	LOGFMTI("IOPCServerDisp::get__NewEnum");
	if (ppUnk == NULL) {
		LOGFMTE("get__NewEnum() failed with invalid argument(s):  ppUnk is NULL");
		return E_INVALIDARG;
	}
	*ppUnk = NULL;

	if (daGenericServer_->m_Enum_Type == OPC_GROUP_ENUM) {
		riid = IID_IEnumUnknown;
		varEnumType = VT_UNKNOWN;
	}
	else {
		riid = IID_IEnumString;
		varEnumType = VT_BSTR;
	}

	// avoid deletion of groups while creating the list
	EnterCriticalSection(&daGenericServer_->m_GroupsCritSec);

	// Get a snapshot of the group list 
	// Note: this does NOT do AddRefs to the groups
	res = daGenericServer_->GetGrpList(riid, daGenericServer_->m_Enum_Scope, &GroupList, &GroupCount);
	if (FAILED(res)) {
		goto ServerDispNewEnumExit0;
	}
	// Create the Enumerator using the snapshot
	// Note that the enumerator will AddRef the server 
	// and also all of the groups.
	temp = new IOpcEnumVariant(varEnumType, GroupCount, (void *)GroupList,
		(IOPCServerDisp *)this, pIMalloc
	);
	if (temp == NULL) {
		res = E_OUTOFMEMORY;
		goto ServerDispNewEnumExit1;
	}

	daGenericServer_->FreeGrpList(riid, GroupList, GroupCount);
	LeaveCriticalSection(&(daGenericServer_->m_GroupsCritSec));

	// Then QI for the interface ('temp') actually is the interface
	// but QI is the 'proper' way to get it.
	// Note QI will do an AddRef of the Enum which will also do
	// an AddRef of the 'parent' - i.e. the 'this' pointer passed above.
	res = temp->QueryInterface(IID_IEnumVARIANT, (LPVOID*)ppUnk);
	if (FAILED(res)) {
		delete temp;
	}
	return res;

ServerDispNewEnumExit1:
	daGenericServer_->FreeGrpList(riid, GroupList, GroupCount);

ServerDispNewEnumExit0:
	LeaveCriticalSection(&daGenericServer_->m_GroupsCritSec);
	return res;
}



//=========================================================================
// IOPCServerDisp.SetEnumeratorType
//
// Create various enumerators for the groups provided by the Server.
//
// [ propget ]
//=========================================================================
STDMETHODIMP DaComBaseServer::SetEnumeratorType(
	/*[in]*/ long  Scope,
	/*[in]*/ short Type)
{
	LOGFMTI("IOPCServerDisp::SetEnumeratorType: Scope=%d, Type=%d", Scope, Type);

	switch (Scope) {
	case OPC_ENUM_PRIVATE_CONNECTIONS:
	case OPC_ENUM_PUBLIC_CONNECTIONS:
	case OPC_ENUM_ALL_CONNECTIONS:
	case OPC_ENUM_PRIVATE:
	case OPC_ENUM_PUBLIC:
	case OPC_ENUM_ALL:
		daGenericServer_->m_Enum_Scope = (OPCENUMSCOPE)Scope;
		break;
	default:
		LOGFMTE("SetEnumeratorType() failed with invalid argument(s):  Scope is not one of OPC_ENUM_xxx constants");
		return E_INVALIDARG;
		break;
	}

	if (Type == OPC_GROUP_ENUM) {
		daGenericServer_->m_Enum_Type = Type;
	}
	else {
		// take name enumeration as default
		daGenericServer_->m_Enum_Type = OPC_GROUPNAME_ENUM;
	}
	return S_OK;
}





//=========================================================================
// IOPCServerDisp.get_StartTime
//
// [ propget ]
//=========================================================================
STDMETHODIMP DaComBaseServer::get_StartTime(
	/*[out, retval]*/ DATE* pStartTime)
{
	DATE     temp;
	HRESULT  res;

	LOGFMTI("IOPCServerDisp::get_StartTime");
	if (pStartTime == NULL) {
		LOGFMTE("get_StartTime() failed with invalid argument(s):  pStartTime is NULL");
		return E_INVALIDARG;
	}
	res = FileTimeToDATE(&(daGenericServer_->m_StartTime), temp);
	*pStartTime = temp;
	return res;
}



//=========================================================================
// IOPCServerDisp.get_CurrentTime
//
// [ propget ]
//=========================================================================
STDMETHODIMP DaComBaseServer::get_CurrentTime(
	/*[out, retval]*/ DATE* pCurrentTime)
{
	FILETIME    now;
	DATE        temp;
	HRESULT     res;

	LOGFMTI("IOPCServerDisp::get_CurrentTime");
	if (pCurrentTime == NULL) {
		LOGFMTE("get_CurrentTime() failed with invalid argument(s):   pCurrentTime is NULL");
		return E_INVALIDARG;
	}

	CoFileTimeNow(&now);
	res = FileTimeToDATE(&now, temp);
	*pCurrentTime = temp;
	return res;
}



//=========================================================================
// IOPCServerDisp.get_LastUpdateTime
//
// [ propget ]
//=========================================================================
STDMETHODIMP DaComBaseServer::get_LastUpdateTime(
	/*[out, retval]*/ DATE* pLastUpdateTime)
{
	DATE     temp;
	HRESULT  res;

	LOGFMTI("IOPCServerDisp::get_LastUpdateTime");
	if (pLastUpdateTime == NULL) {
		LOGFMTE("get_LastUpdateTime() failed with invalid argument(s):  pLastUpdateTime is NULL");
		return E_INVALIDARG;
	}

	res = FileTimeToDATE(&daGenericServer_->m_LastUpdateTime, temp);
	*pLastUpdateTime = temp;
	return res;
}



//=========================================================================
// IOPCServerDisp.get_MajorVersion
//
// [ propget ]
//=========================================================================
STDMETHODIMP DaComBaseServer::get_MajorVersion(
	/*[out, retval]*/ short* pMajorVersion)
{
	LOGFMTI("IOPCServerDisp::get_MajorVersion");
	if (pMajorVersion == NULL) {
		LOGFMTE("get_MajorVersion() failed with invalid argument(s):  pMajorVersion is NULL");
		return E_INVALIDARG;
	}

	*pMajorVersion = core_generic_main.m_VersionInfo.m_wMajor;
	return S_OK;
}



//=========================================================================
// IOPCServerDisp.get_MinorVersion
//
// [ propget ]
//=========================================================================
STDMETHODIMP DaComBaseServer::get_MinorVersion(
	/*[out, retval]*/ short* pMinorVersion)
{
	LOGFMTI("IOPCServerDisp::get_MinorVersion");
	if (pMinorVersion == NULL) {
		LOGFMTE("get_MinorVersion() failed with invalid argument(s):  pMinorVersion is NULL");
		return E_INVALIDARG;
	}

	*pMinorVersion = core_generic_main.m_VersionInfo.m_wMinor;
	return S_OK;
}



//=========================================================================
// IOPCServerDisp.get_BuildNumber
//
// [ propget ]
//=========================================================================
STDMETHODIMP DaComBaseServer::get_BuildNumber(
	/*[out, retval]*/ short* pBuildNumber)
{
	LOGFMTI("IOPCServerDisp::get_BuildNumber");
	if (pBuildNumber == NULL) {
		LOGFMTE("get_BuildNumber() failed with invalid argument(s):  pBuildNumber is NULL");
		return E_INVALIDARG;
	}

	*pBuildNumber = core_generic_main.m_VersionInfo.m_wBuild;
	return S_OK;
}



//=========================================================================
// IOPCServerDisp.get_VendorInfo
//
// [ propget ]
//=========================================================================
STDMETHODIMP DaComBaseServer::get_VendorInfo(
	/*[out, retval]*/ BSTR* pVendorInfo)
{
	LOGFMTI("IOPCServerDisp::get_VendorInfo");
	if (pVendorInfo == NULL) {
		LOGFMTE("get_VendorInfo() failed with invalid argument(s):  pVendorInfo is NULL");
		return E_INVALIDARG;
	}

	// Get the vendor name
	LPCTSTR pszVendor = core_generic_main.m_VersionInfo.GetValue(_T("CompanyName"));
	if (pszVendor == NULL) {
		_ASSERTE(0);                            // The company name string must exist in the Version Info
		return E_FAIL;
	}
	USES_CONVERSION;                             // Custom interface needs global memory for strings
	*pVendorInfo = SysAllocString(T2COLE(pszVendor));
	if (*pVendorInfo == NULL) {
		return E_OUTOFMEMORY;
	}
	return S_OK;
}



//=========================================================================
// IOPCServerDisp.Item
//
// Special method used by Automation to implement the FOR/NEXT feature of Basic.
// The term Item in this context is a VB term which represents an item in a collection.
// It has no relation to an OPCItem.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::Item(
	/*[in]         */ VARIANT      Item,
	/*[out, retval]*/ IDispatch ** ppDisp)
{
	long                    idx, nidx;
	long                    AutoIdx, i;
	HRESULT                 res;
	DaGenericGroup           *group;

	LOGFMTI("IOPCServerDisp.Item");
	if (ppDisp != NULL) {
		*ppDisp = NULL;
	}

	if ((ppDisp == NULL) || ((V_VT(&Item) != VT_I2) && (V_VT(&Item) != VT_I4) && (V_VT(&Item) != VT_BSTR))) {
		LOGFMTE("Item() failed with invalid argument(s):");
		if (ppDisp == NULL) {
			LOGFMTE("      ppDisp is NULL");
		}
		// only integer and string collection member ids are accepted
		if ((V_VT(&Item) != VT_I2) && (V_VT(&Item) != VT_I4) && (V_VT(&Item) != VT_BSTR)) {
			LOGFMTE("      Item type must be VT_I2, VT_I4 or VT_BSTR");
		}

		return E_INVALIDARG;
	}

	if ((V_VT(&Item) == VT_I2) || (V_VT(&Item) == VT_I4)) {
		AutoIdx = 0;
		if (V_VT(&Item) == VT_I4) {
			AutoIdx = V_I4(&Item);
		}
		else if (V_VT(&Item) == VT_I2) {
			AutoIdx = V_I2(&Item);
		}
		if (AutoIdx <= 0) {
			return E_INVALIDARG;
		}

		// search by position in array of groups
		EnterCriticalSection(&(daGenericServer_->m_GroupsCritSec));

		i = 1;
		res = daGenericServer_->m_GroupList.First(&idx);
		while ((SUCCEEDED(res)) && (i < AutoIdx)) {
			i++;
			res = daGenericServer_->m_GroupList.Next(idx, &nidx);
			idx = nidx;
		}

		if (FAILED(res)) {
			// invalid index
			res = E_FAIL;
			goto ServerGroupItemExit1;
		}

		daGenericServer_->m_GroupList.GetElem(idx, &group);

	}
	else { // Item type is VT_BSTR
	   // search by name in array of groups

		EnterCriticalSection(&(daGenericServer_->m_GroupsCritSec));

		// search private group
		res = daGenericServer_->SearchGroup(FALSE,
			V_BSTR(&Item),
			&group,
			&idx);
		if (FAILED(res)) {
			// search public group
			res = daGenericServer_->SearchGroup(TRUE,
				V_BSTR(&Item),
				&group,
				&idx);
		}

		if (FAILED(res)) {
			goto ServerGroupItemExit1;
		}

	}

	res = daGenericServer_->GetCOMGroup(group, IID_IDispatch, (LPUNKNOWN *)ppDisp);
	if (FAILED(res)) {
		goto ServerGroupItemExit1;
	}
	LeaveCriticalSection(&daGenericServer_->m_GroupsCritSec);
	return S_OK;

ServerGroupItemExit1:
	LeaveCriticalSection(&daGenericServer_->m_GroupsCritSec);
	return res;
}



//=========================================================================
// IOPCServerDisp.AddGroup
//
// Adds a private group
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::AddGroup(
	/*[in]          */ BSTR         Name,
	/*[in]          */ VARIANT_BOOL Active,
	/*[in]          */ long         RequestedUpdateRate,
	/*[in]          */ long         ClientGroupHandle,
	/*[in]          */ float      * pPercentDeadband,
	/*[in]          */ long         LCID,
	/*[out]         */ long       * pServerGroupHandle,
	/*[out]         */ long       * pRevisedUpdateRate,
	/*[in, optional]*/ VARIANT    * pTimeBias,
	/*[out, retval] */ IDispatch ** ppDisp)
{
	DaGenericGroup         *theGroup;
	HRESULT               res;
	long                  *pTempTimeBias;
	long                  ServerGroupHandle, sgh;
	LPWSTR                theName;
	int                   bIntActive;
	DaGenericGroup         *foundGroup;

	USES_CONVERSION;
	LOGFMTI("IOPCServerDisp.AddGroup %s", W2A(Name));

	if (pServerGroupHandle != NULL) {
		*pServerGroupHandle = 0;
	}
	if (pRevisedUpdateRate != NULL) {
		*pRevisedUpdateRate = 0;
	}
	if (ppDisp != NULL) {
		*ppDisp = NULL;
	}

	if ((pServerGroupHandle == NULL)
		|| (pRevisedUpdateRate == NULL)
		|| (ppDisp == NULL)) {

		LOGFMTE("AddGroup() failed with invalid argument(s):");
		if (pServerGroupHandle == NULL) {
			LOGFMTE("      pServerGroupHandle is NULL");
		}
		if (pRevisedUpdateRate == NULL) {
			LOGFMTE("      pRevisedUpdateRate is NULL");
		}
		if (ppDisp == NULL) {
			LOGFMTE("      ppDisp is NULL");
		}

		return E_INVALIDARG;
	}

	if (Active == VARIANT_FALSE) {
		bIntActive = FALSE;
	}
	else {
		bIntActive = TRUE;
	}

	// manipulation of the group lists must be protected 
	EnterCriticalSection(&(daGenericServer_->m_GroupsCritSec));

	if ((Name == NULL) || (wcscmp(Name, L"") == 0)) {
		// no group name specified 

		// create unique name ...
		ServerGroupHandle = daGenericServer_->m_GroupList.Size();
		res = daGenericServer_->GetGroupUniqueName(ServerGroupHandle, &theName, FALSE);
		if (FAILED(res)) {
			goto SDAddGroupExit00;
		}
	}
	else {
		// check there is not another one with the same name
		// between all private groups
		theName = WSTRClone(Name, NULL);
		if (theName == NULL) {
			res = E_OUTOFMEMORY;
			goto SDAddGroupExit00;
		}
		res = daGenericServer_->SearchGroup(FALSE, theName, &foundGroup, &sgh);
		if (SUCCEEDED(res)) {             // found another group with that name!
			res = OPC_E_DUPLICATENAME;
			goto SDAddGroupExit0;
		}
	}

	// get time bias in the needed format
	if ((pTimeBias != NULL) && (V_VT(pTimeBias) == VT_I4)) {
		pTempTimeBias = &V_I4(pTimeBias);
	}
	else {
		pTempTimeBias = NULL;
	}

	// construct DaGenericGroup object
	theGroup = new DaGenericGroup();
	if (theGroup == NULL) {
		res = E_OUTOFMEMORY;
		goto SDAddGroupExit0;
	}

	// create DaGenericGroup object
	res = theGroup->Create(
		daGenericServer_,
		theName,
		bIntActive,
		pTempTimeBias,
		RequestedUpdateRate,
		ClientGroupHandle,
		pPercentDeadband,
		LCID,
		FALSE,
		0,
		&ServerGroupHandle
	);
	if (FAILED(res)) {
		goto SDAddGroupExit1;
	}

	res = daGenericServer_->GetCOMGroup(theGroup, IID_IDispatch, (LPUNKNOWN *)ppDisp);
	if (FAILED(res)) {
		goto SDAddGroupExit1;
	}

	*pRevisedUpdateRate = theGroup->m_RevisedUpdateRate;
	*pServerGroupHandle = ServerGroupHandle;

	WSTRFree(theName, NULL);

	LeaveCriticalSection(&daGenericServer_->m_GroupsCritSec);
	return S_OK;

	// function failed
SDAddGroupExit1:
	delete theGroup;

SDAddGroupExit0:
	WSTRFree(theName, NULL);

SDAddGroupExit00:
	LeaveCriticalSection(&daGenericServer_->m_GroupsCritSec);
	return res;
}




//=========================================================================
// IOPCServerDisp.GetErrorString
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::GetErrorString(
	/*[in]         */ long   Error,
	/*[in]         */ long   Locale,
	/*[out, retval]*/ BSTR * ErrorString)
{
	HRESULT  hres;
	LPWSTR   szError;

	LOGFMTI("IOPCServerDisp::GetErrorString for error code %lX", Error);

	*ErrorString = NULL;

	hres = OpcCommon::GetErrorString(Error, Locale, &szError);
	if (SUCCEEDED(hres)) {
		*ErrorString = SysAllocString(szError);
		pIMalloc->Free(szError);
		if (*ErrorString == NULL) {
			hres = E_OUTOFMEMORY;
		}
	}
	return hres;
}




//=========================================================================
// IOPCServerDisp.GetGroupByName
//    get the IDispatch interface of the private group with 
//    given name
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::GetGroupByName(
	/*[in]         */ BSTR         Name,
	/*[out, retval]*/ IDispatch ** ppDisp)
{
	long                    idx;
	DaGenericGroup           *group;
	HRESULT                 res;
	long                    pgh;

	USES_CONVERSION;
	LOGFMTI("IOPCServerDisp.GetGroupByName:  Name=%s", W2A(Name));
	if (ppDisp == NULL) {
		LOGFMTE("GetGroupByName() failed with invalid argument(s):  ppDisp is NULL");
		return E_INVALIDARG;
	}
	*ppDisp = NULL;

	EnterCriticalSection(&daGenericServer_->m_GroupsCritSec);

	res = daGenericServer_->m_GroupList.First(&idx);
	while (res != E_FAIL) {
		daGenericServer_->m_GroupList.GetElem(idx, &group);
		if (group->GetPublicInfo(&pgh) == FALSE) {
			if (wcscmp(group->m_Name, Name) == 0) {
				// found it!
				res = daGenericServer_->GetCOMGroup(group,
					IID_IDispatch, (LPUNKNOWN *)ppDisp);
				goto SDGroupByNameExit1;
			}
		}
		res = daGenericServer_->m_GroupList.Next(idx, &idx);
	}
	res = E_INVALIDARG;

SDGroupByNameExit1:
	LeaveCriticalSection(&daGenericServer_->m_GroupsCritSec);
	return res;
}



//=========================================================================
// IOPCServerDisp.RemoveGroup
//
// Inputs:
//    - Force is ignored, but rem that the values are VARIANT_FALSE and VARIANT_TRUE
//       and not TRUE and FALSE
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::RemoveGroup(
	/*[in]*/ long    ServerGroupHandle,
	/*[in]*/ VARIANT_BOOL Force)
{
	HRESULT                 res;
	DaGenericGroup           *group;
	long                    pgh;

	LOGFMTI("IOPCServerDisp.RemoveGroup: Handle=%Xh, Forced=%d", ServerGroupHandle, Force);

	res = daGenericServer_->GetGenericGroup(ServerGroupHandle, &group);
	if (FAILED(res)) {
		return res;
	}
	EnterCriticalSection(&(group->m_CritSec));

	if (group->GetPublicInfo(&pgh) == TRUE) {
		LeaveCriticalSection(&(group->m_CritSec));
		res = OPC_E_PUBLIC;
		goto SDRemoveGroupExit1;
	}
	res = daGenericServer_->RemoveGenericGroup(ServerGroupHandle);
	LeaveCriticalSection(&(group->m_CritSec));

SDRemoveGroupExit1:
	daGenericServer_->ReleaseGenericGroup(ServerGroupHandle);
	return res;
}





// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
// [[[[[[[[[[[[[[[[[[[[ IOPCServerPublicGroupsDisp [[[[[[[[[[[[[[[[[[[[[[[[
// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]

//=========================================================================
// IOPCServerPublicGroupsDisp.GetPublicGroupByName
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::GetPublicGroupByName(
	/*[in]         */ BSTR         Name,
	/*[out, retval]*/ IDispatch ** ppDisp)
{
	HRESULT     res;
	IUnknown    *pUnk;

	USES_CONVERSION;
	LOGFMTI("IOPCServerPublicGroupsDisp.GetPublicGroupByName: Name=%s", W2A(Name));
	if (ppDisp == NULL) {
		LOGFMTE("GetPublicGroupByName() failed with invalid argument(s):  ppDisp is NULL");
		return E_INVALIDARG;
	}
	*ppDisp = NULL;

	res = GetPublicGroupByName(Name, IID_IDispatch, &pUnk);
	if (SUCCEEDED(res)) {
		*ppDisp = (IDispatch *)pUnk;
	}
	return res;
}


//=========================================================================
// IOPCServerPublicGroupsDisp.RemovePublicGroup
//
// This function simply calls the custom interface implementation of this 
// method
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::RemovePublicGroup(
	/*[in]*/ long    ServerGroupHandle,
	/*[in]*/ VARIANT_BOOL Force)
{
	HRESULT res;

	LOGFMTI("IOPCServerPublicGroupsDisp.RemovePublicGroup: Handle=%Xh, Force=%d", ServerGroupHandle, Force);

	// call the custom interface  RemovePublicGroup
	if (Force == VARIANT_FALSE) {
		res = RemovePublicGroup((OPCHANDLE)ServerGroupHandle, FALSE);
	}
	else {
		res = RemovePublicGroup((OPCHANDLE)ServerGroupHandle, TRUE);
	}
	return res;
}



// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
// [[[[[[[[[[[[[[[[[ IOPCBrowseServerAddressSpaceDisp [[[[[[[[[[[[[[[[[[[[[
// ]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]

//=========================================================================
// IOPCBrowseServerAddressSpaceDisp.get__NewEnum
//
// parameter:
//    -
// return:
//    -
//=========================================================================
//[ propget, restricted, id( DISPID_NEWENUM )   ]
HRESULT STDMETHODCALLTYPE DaComBaseServer::get__NewEnum_BROWSE(
	/*[out, retval]*/ IUnknown ** ppUnk)
{
	IOpcEnumVariant   *temp;
	HRESULT           res;
	LPWSTR            *pElements;
	DWORD             nElements;
	DWORD             i;

	LOGFMTI("IOPCBrowseServerAddressSpaceDisp.get__NewEnum");
	if (ppUnk == NULL) {
		LOGFMTE("get__NewEnum_BROWSE() failed with invalid argument(s):  ppUnk is NULL");
		return E_INVALIDARG;
	}
	*ppUnk = NULL;

	if (daGenericServer_->m_EnumeratingItemIDs == TRUE) {
		res = daBaseServer_->OnBrowseItemIdentifiers(
			daGenericServer_->m_BrowseData.m_bstrServerBrowsePosition,
			(OPCBROWSETYPE)daGenericServer_->m_BrowseFilterType,
			daGenericServer_->m_FilterCriteria,
			daGenericServer_->m_DataTypeFilter,
			daGenericServer_->m_AccessRightsFilter,
			&nElements,
			&pElements,
			&daGenericServer_->m_BrowseData.m_pCustomData
		);
	}
	else {
		// enumerating access paths
		res = daBaseServer_->OnBrowseAccessPaths(
			daGenericServer_->m_BrowseItemIDForAccessPath,
			&nElements,
			&pElements,
			&daGenericServer_->m_BrowseData.m_pCustomData
		);
	}
	if (FAILED(res)) {
		return res;
	}

	// Create the Enumerator 
	temp = new IOpcEnumVariant(VT_BSTR, nElements, (void *)pElements,
		(IOPCServerDisp *)this, pIMalloc
	);
	if (temp == NULL) {
		res = E_OUTOFMEMORY;
		goto NewEnumBrowseExit0;
	}

	// Query the interface for the client
	res = temp->QueryInterface(IID_IEnumVARIANT, (LPVOID*)ppUnk);
	if (FAILED(res)) {
		delete temp;
		goto NewEnumBrowseExit0;
	}

	res = S_OK;

NewEnumBrowseExit0:
	if (pElements != NULL) {
		i = 0;
		while (i < nElements) {
			if (pElements[i] != NULL) {
				SysFreeString(pElements[i]);
			}
			i++;
		}
		delete pElements;
	}

	return res;
}




//=========================================================================
// IOPCBrowseServerAddressSpaceDisp.get_Organization
//
//=========================================================================
////[ propget ]
HRESULT STDMETHODCALLTYPE DaComBaseServer::get_Organization(
	/*[out, retval]*/ long * pNameSpaceType)
{
	LOGFMTI("IOPCBrowseServerAddressSpaceDisp::get_Organization");

	if (!pNameSpaceType) {
		LOGFMTE("get_Organization() failed with invalid argument(s):  pNameSpaceType is NULL");
		return E_INVALIDARG;
	}
	*pNameSpaceType = daBaseServer_->OnBrowseOrganization();
	return S_OK;
}



//=========================================================================
// IOPCBrowseServerAddressSpaceDisp.ChangeBrowsePosition
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::ChangeBrowsePosition(
	/*[in]*/ long BrowseDirection,
	/*[in]*/ BSTR Position)
{
	OPCWSTOAS(Position)
		LOGFMTI("IOPCBrowseServerAddressSpaceDisp.ChangeBrowsePosition: from %s, Direction=%d", OPCastr, BrowseDirection);
}

if ((BrowseDirection == OPC_BROWSE_DOWN) && (Position == NULL)) {
	return E_INVALIDARG;                      // Branch must not be NULL.
}

// Use current position as initial position.
BSTR bstrNewPosition = daGenericServer_->m_BrowseData.m_bstrServerBrowsePosition;

HRESULT hr = daBaseServer_->OnBrowseChangeAddressSpacePosition(
(OPCBROWSEDIRECTION)BrowseDirection,
(LPCWSTR)Position,
&bstrNewPosition,
&daGenericServer_->m_BrowseData.m_pCustomData);
if (SUCCEEDED(hr)) {
	SysFreeString(daGenericServer_->m_BrowseData.m_bstrServerBrowsePosition);
	daGenericServer_->m_BrowseData.m_bstrServerBrowsePosition = bstrNewPosition;
}
return hr;
}




//=========================================================================
// IOPCBrowseServerAddressSpaceDisp.SetItemIDEnumerator
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::SetItemIDEnumerator(
	/*[in]*/          long         BrowseFilterType,
	/*[in]*/          BSTR         FilterCriteria,
	/*[in]*/          VARIANT      DataTypeFilter,
	/*[in]*/          long         AccessRightsFilter)
{
	LPWSTR TempStr;

	OPCWSTOAS(FilterCriteria)
		LOGFMTI("IOPCBrowseServerAddressSpaceDisp.SetItemIDEnumerator: Filter=%s, Type=%d", OPCastr, BrowseFilterType);
}
// Check the passed arguments
if (AccessRightsFilter & ~(OPC_READABLE | OPC_WRITEABLE)) {
	LOGFMTE("SetItemIDEnumerator() failed with invalid argument(s): AccessRightsFilter with unknown type");
	return E_INVALIDARG;                      // Illegal Access Rights Filter
}
if ((BrowseFilterType != OPC_BRANCH) &&
(BrowseFilterType != OPC_LEAF) &&
(BrowseFilterType != OPC_FLAT)) {
	LOGFMTE("SetItemIDEnumerator() failed with invalid argument(s): BrowseFilterType with unknown type");
	return E_INVALIDARG;
}

if (FilterCriteria == NULL) {
	TempStr = NULL;
}
else {
	TempStr = SysAllocString(FilterCriteria);
	if (TempStr == NULL) {
		return E_OUTOFMEMORY;
	}
}
daGenericServer_->m_BrowseFilterType = BrowseFilterType;

if (daGenericServer_->m_FilterCriteria != NULL) {
	SysFreeString(daGenericServer_->m_FilterCriteria);
}
daGenericServer_->m_FilterCriteria = TempStr;
daGenericServer_->m_DataTypeFilter = V_VT(&DataTypeFilter);
daGenericServer_->m_AccessRightsFilter = AccessRightsFilter;
daGenericServer_->m_EnumeratingItemIDs = TRUE;
return S_OK;
}



//=========================================================================
// IOPCBrowseServerAddressSpaceDisp.GetItemIDString
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::GetItemIDString(
	/*[in]         */ BSTR   ItemDataID,
	/*[out, retval]*/ BSTR * ItemID)
{
	HRESULT res;
	LPWSTR   FullItemID;

	USES_CONVERSION;
	LOGFMTI("IOPCBrowseServerAddressSpaceDisp.GetItemIDString of item %s", W2A(ItemDataID));

	if (ItemID == NULL) {
		LOGFMTE("GetItemIDString() failed with invalid argument(s):  ItemID is NULL");
		return E_INVALIDARG;
	}
	*ItemID = NULL;
	res = daBaseServer_->OnBrowseGetFullItemIdentifier(
		daGenericServer_->m_BrowseData.m_bstrServerBrowsePosition,
		ItemDataID, &FullItemID,
		&daGenericServer_->m_BrowseData.m_pCustomData);
	if (FAILED(res)) {
		return res;
	}
	*ItemID = FullItemID;
	return res;
}


//=========================================================================
// IOPCBrowseServerAddressSpaceDisp.SetAccessPathEnumerator
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaComBaseServer::SetAccessPathEnumerator(
	/*[in]*/          BSTR         ItemID)
{
	LPWSTR TempStr;

	USES_CONVERSION;
	LOGFMTI("IOPCBrowseServerAddressSpaceDisp.SetAccessPathEnumerator for item %s", W2A(ItemID));

	if (ItemID == NULL) {
		TempStr = NULL;
	}
	else {
		TempStr = SysAllocString(ItemID);
		if (TempStr == NULL) {
			return E_OUTOFMEMORY;
		}
	}

	if (daGenericServer_->m_BrowseItemIDForAccessPath != NULL) {
		SysFreeString(daGenericServer_->m_BrowseItemIDForAccessPath);
	}

	daGenericServer_->m_BrowseItemIDForAccessPath = TempStr;

	// the next call of NewEnum will cause enumeration of access paths (and not intem IDs!)
	daGenericServer_->m_EnumeratingItemIDs = FALSE;

	return S_OK;
}


//DOM-IGNORE-END
