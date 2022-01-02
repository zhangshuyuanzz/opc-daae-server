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
#include "DaGroup.h"
#include "DaGenericGroup.h"

#include "UtilityFuncs.h"
#include "DaComServer.h"
#include "enumclass.h"

//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCItemMgtDisp [[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

//=========================================================================
// IOPCItemMgtDisp.get_Count
// -------------------------
// Return the number of items in the group.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::get_Count(
	/*[out, retval]*/ long * pCount )
{
	DaGenericGroup *group;
	HRESULT res;

	res = GetGenericGroup( &group );
	if( FAILED(res)) {
      LOGFMTE("IOPCItemMgtDisp.get_Count: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI("IOPCItemMgtDisp.get_Count of group %s", W2A(group->m_Name));

	if( pCount == NULL ) {
      LOGFMTE("get_Count() failed with invalid argument(s):  pCount is NULL");
		ReleaseGenericGroup();
		return E_INVALIDARG;
	}

	*pCount = -1;
	EnterCriticalSection( &(group->m_ItemsCritSec) );
	*pCount = group->m_oaItems.TotElem();
	LeaveCriticalSection( &(group->m_ItemsCritSec) );
	ReleaseGenericGroup();
	return S_OK;
}




//=========================================================================
// IOPCItemMgtDisp.get__NewEnum
// ----------------------------
// Item enumeration.
// Create a new enumerator class and build an enumeration array with
// all elements.
//=========================================================================
////[ propget, 
////  restricted, 
////  id( DISPID_NEWENUM )
////]
HRESULT STDMETHODCALLTYPE DaGroup::get__NewEnum(
	/*[out, retval]*/ IUnknown ** ppUnk  )
{
	IOpcEnumVariant   *temp;
	long               ItemCount;
	HRESULT            res;
	LPUNKNOWN         *ItemList;
	DaGenericGroup     *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
		LOGFMTE("IOPCItemMgtDisp.get_New_Enum: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI("IOPCItemMgtDisp.get__NewEnum for group %s", W2A(group->m_Name));

	if( ppUnk == NULL ) {
		LOGFMTE("get__NewEnum() failed with invalid argument(s): ppUnk is NULL");
		res = E_INVALIDARG;
		goto ItemMgtDispNewEnumExit0;
	}

	*ppUnk = NULL;

	// avoid deletion of groups while creating the list
	EnterCriticalSection( &(group->m_ItemsCritSec) );


	// Get a snapshot of the group list 
	// Note: this does NOT do AddRefs to the groups
	res = GetCOMItemList( group, &ItemList, &ItemCount);

	LeaveCriticalSection( &group->m_ItemsCritSec );
	if( FAILED(res) ) {
		goto ItemMgtDispNewEnumExit0;
	}


	// Create the Enumerator using the snapshot
	// Note that the enumerator will AddRef the server 
	// and also all of the groups.
	temp = new IOpcEnumVariant( VT_UNKNOWN, ItemCount, ItemList, 
		(IOPCServerDisp *)this, pIMalloc
		);
	if( temp == NULL ) {
		res = E_OUTOFMEMORY;
		goto ItemMgtDispNewEnumExit1;
	}

	FreeCOMItemList( ItemList, ItemCount );

	// Then QI for the interface ('temp') actually is the interface
	// but QI is the 'proper' way to get it.
	// Note QI will do an AddRef of the Enum which will also do
	// an AddRef of the 'parent' - i.e. the 'this' pointer passed above.
	res = temp->QueryInterface( IID_IEnumVARIANT, (LPVOID*)ppUnk);
	if( FAILED( res ) ) {
		delete temp;
	}

	ReleaseGenericGroup( );
	return res;

ItemMgtDispNewEnumExit1:
	FreeCOMItemList( ItemList, ItemCount );

ItemMgtDispNewEnumExit0:
	ReleaseGenericGroup( );
	return res;
}


//=========================================================================
// IOPCItemMgtDisp.Item
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::Item(
	/*[in]         */ VARIANT      ItemSpecifier, 
	/*[out, retval]*/ IDispatch ** ppDisp    )
{
	long                    idx, nidx;
	long                    AutoIdx, i;
	HRESULT                 res;
	DaGenericGroup           *group;
	BOOL                    created;
	DaItem                *pCOMItem;
	DaGenericItem            *item;
	BOOL                    found;
	LPWSTR                  Name;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCItemMgtDisp.Item: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCItemMgtDisp.Item in group %s", W2A(group->m_Name));

	if( ppDisp == NULL ) {
      LOGFMTE("Item() failed with invalid argument(s):  ppDisp is NULL");
		res = E_INVALIDARG;
		goto ServerGroupItemExit0;
	}
	*ppDisp = NULL;

	// only integer and string collection member ids are accepted
	if( ( V_VT(&ItemSpecifier) != VT_I2 ) && ( V_VT(&ItemSpecifier) != VT_I4 ) && (V_VT(&ItemSpecifier) != VT_BSTR) ) {
		res = E_INVALIDARG;
		goto ServerGroupItemExit0;
	}


	if( ( V_VT(&ItemSpecifier) == VT_I2 ) || ( V_VT(&ItemSpecifier) == VT_I4 ) ) {
		AutoIdx = 0;
		if( V_VT(&ItemSpecifier) == VT_I4 ) {
			AutoIdx = V_I4( &ItemSpecifier );
		} else if( V_VT(&ItemSpecifier) == VT_I2 ) {
			AutoIdx = V_I2( &ItemSpecifier );
		}
		if( AutoIdx <= 0 ) {
			res = E_INVALIDARG;
			goto ServerGroupItemExit0;
		}

		// avoid changes in list while getting item
		EnterCriticalSection( &(group->m_ItemsCritSec) );

		i = 1;
		res = group->m_oaItems.First( &idx );
		while ( (SUCCEEDED(res)) && (i < AutoIdx) ) {
			i ++;
			res = group->m_oaItems.Next( idx, &nidx );
			idx = nidx;
		}

		if( FAILED(res) ) {
			// invalid index
			res = E_FAIL;
			goto GroupItemItemExit1;
		}

		group->m_oaItems.GetElem( idx, &item );

	} else { // Item type is VT_BSTR
		// search by name in array of groups

		// avoid changes in list while getting item
		EnterCriticalSection( &(group->m_ItemsCritSec) );

		found = FALSE;
		res = group->m_oaItems.First( &idx );
		while ( (found == FALSE) && (SUCCEEDED(res)) ) {
			group->m_oaItems.GetElem( idx, &item );
			// get the name. must be freed!
			res = item->get_ItemIDCopy( &Name );
			if( SUCCEEDED( res )) {
				if( (Name != NULL) && ( item->Killed() == FALSE ) && ( wcscmp( Name, V_BSTR( &ItemSpecifier ) ) == 0 ) ) {
					found = TRUE;
				} else {
					res = group->m_oaItems.Next( idx, &nidx );
					idx = nidx;
				}
				if (Name) {
					WSTRFree( Name, NULL );
				}
			}
		}

		if( FAILED(res) ) {
			// invalid index
			res = E_FAIL;
			goto GroupItemItemExit1;
		}

		if( FAILED( res ) ) {
			goto GroupItemItemExit1;
		}
	}

	res = group->GetCOMItem(    
		idx,
		item,
		&pCOMItem,
		&created );

	if( FAILED(res) ) {
		goto GroupItemItemExit1;
	}

	// query interface
	res = pCOMItem->QueryInterface( IID_IDispatch, (void **)ppDisp );
	if( FAILED(res) ) {
		// if failed then delete the Item because no one will do it
		// ( this because the item is deleted only with release
		//   when the RefCount goes to 0 )
		if( created == TRUE ) {
			// reset created COM Item
			delete pCOMItem;
		}
		goto GroupItemItemExit1;
	}

	LeaveCriticalSection( &(group->m_ItemsCritSec) );
	ReleaseGenericGroup( );

	return S_OK;

GroupItemItemExit1:
	LeaveCriticalSection( &(group->m_ItemsCritSec) );

ServerGroupItemExit0:
	ReleaseGenericGroup( );
	return res;

}




//=========================================================================
// IOPCItemMgtDisp.AddItems
// ------------------------
// Add items to the group.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::AddItems(
	/*[in]               */ long      NumItems,
	/*[in]               */ VARIANT   ItemIDs,
	/*[in]               */ VARIANT   ActiveStates,
	/*[in]               */ VARIANT   ClientHandles,
	/*[out]              */ VARIANT * pServerHandles,
	/*[out]              */ VARIANT * errors,
	/*[out]              */ VARIANT * pItemObjects,
	/*[in, optional]     */ VARIANT   AccessPaths,
	/*[in, optional]     */ VARIANT   RequestedDataTypes,
	/*[in, out, optional]*/ VARIANT * pBlobs,
	/*[out, optional]    */ VARIANT * pCanonicalDataTypes,
	/*[out, optional]    */ VARIANT * pAccessRights
	)
{
	DaGenericItem    *NewGI;
	DaDeviceItem    **ppDItems;
	HRESULT          res;
	long             i;
	long             hServerHandle;
	OPCITEMDEF     * pItemArray;
	HRESULT        * pTempErrors;
	DaItem       * pCOMItem;
	DaGenericGroup  *group;
	OPCITEMDEF     *pID;
	BYTE           *pBlob;
	DWORD           BlobSize;
	VARIANT_BOOL    bActive;
	HRESULT         theRes;
	VARIANT         theBlobVar;   
	VARIANT         *pTheBlobVar; 
	BOOL            created;
	long            pgh;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCItemMgtDisp.AddItems: Internal error: No generic group.");
		goto IMDAddItemsBadExit00;
	}
   OPCWSTOAS( group->m_Name )
      LOGFMTI("IOPCItemMgtDisp.AddItems: Add %ld items to group %s", NumItems, OPCastr );
   }
	if( pServerHandles != NULL ) {
		VariantInit( pServerHandles );
	}
	if( errors != NULL ) {
		VariantInit( errors );
	}
	if( pItemObjects != NULL ) {
		VariantInit( pItemObjects );
	}
	if(( NumItems <= 0 ) 
		|| ( V_VT(&ActiveStates) != (VT_ARRAY | VT_BOOL) ) 
		|| ( V_VT(&ItemIDs) != (VT_ARRAY | VT_BSTR) ) 
		|| ( V_VT(&ClientHandles) != (VT_ARRAY | VT_I4) ) 
		|| ( (V_VT(&AccessPaths) != (VT_ARRAY | VT_BSTR ) ) &&      (V_VT(&AccessPaths) != VT_NULL) ) 
		|| ( (V_VT(&RequestedDataTypes) != ( VT_ARRAY | VT_I2) ) && (V_VT(&RequestedDataTypes) != VT_NULL) ) 
		|| ( (pBlobs != NULL) && (V_VT(pBlobs) != ( VT_ARRAY | VT_VARIANT ) ) && (V_VT(pBlobs) != VT_NULL) ) 
		|| ( pServerHandles == NULL ) 
		|| ( errors == NULL ) 
		|| ( pItemObjects == NULL )  ) {

         LOGFMTE( "AddItems() failed with invalid argument(s):" );
         if( NumItems <= 0 ) {
            LOGFMTE( "      NumItems is %ld", NumItems );
         }
         if( V_VT(&ActiveStates) != (VT_ARRAY | VT_BOOL) ) {
            LOGFMTE( "      ActiveStates is not array of VT_BOOL" );
         }
         if( V_VT(&ItemIDs) != (VT_ARRAY | VT_BSTR) ) {
            LOGFMTE( "      ItemIDs is not array of VT_BSTR" );
         }
         if( V_VT(&ClientHandles) != (VT_ARRAY | VT_I4) ) {
            LOGFMTE( "      ClientHandles is not array of VT_I4" );
         }
         if( (V_VT(&AccessPaths) != (VT_ARRAY | VT_BSTR ) ) && (V_VT(&AccessPaths) != VT_NULL) ) {
            LOGFMTE( "      AccessPaths is not VT_NULL or array of VT_BSTR" );
         }
         if( (V_VT(&RequestedDataTypes) != ( VT_ARRAY | VT_I2) ) && (V_VT(&RequestedDataTypes) != VT_NULL) ) {
            LOGFMTE( "      RequestedDataTypes is not VT_NULL or array of VT_I2" );
         }
         if( (pBlobs != NULL) && (V_VT(pBlobs) != ( VT_ARRAY | VT_VARIANT ) ) && (V_VT(pBlobs) != VT_NULL) ) {
            LOGFMTE( "      pBlobs is not VT_NULL or array of VT_VARIANT" );
         }
         if( pServerHandles == NULL ) {
            LOGFMTE( "      pServerHandles is NULL" );
         }
         if( errors == NULL ) {
            LOGFMTE( "      errors is NULL" );
         }
         if( pItemObjects == NULL ) {
            LOGFMTE( "      pItemObjects is NULL" );
         }

			res = E_INVALIDARG;
			goto IMDAddItemsBadExit00;
	}

	if( group->GetPublicInfo( &pgh ) == TRUE ) {
		res = OPC_E_PUBLIC;
		goto IMDAddItemsBadExit0;
	}

	// all output VARIANTs must be created with
	// COM memory management object

	// create errors array ( of HRESULT)
	res = VariantInitArrayOf( errors, NumItems, VT_I4 );
	if( FAILED(res) ) {
		goto IMDAddItemsBadExit0;
	}
	// initialize with E_FAIL
	{
		long *temp;
		long i;

		temp = (long *)(V_ARRAY(errors)->pvData);

		i = 0;
		while ( i < NumItems ) {
			temp[i] = E_FAIL;
			i++;
		}
	}



	// create itemobjects array (of Dispatch *)
	res = VariantInitArrayOf( pItemObjects, NumItems, VT_DISPATCH );
	if( FAILED(res) ) {
		goto IMDAddItemsBadExit1;
	}


	// create canonicaldatatypes array (of VT_I4)
	res = VariantInitArrayOf( pServerHandles, NumItems, VT_I4 );
	if( FAILED(res) ) {
		goto IMDAddItemsBadExit2;
	}


	if( ( pCanonicalDataTypes != NULL ) && ( V_VT( pCanonicalDataTypes ) != VT_NULL ) ) {
		// create canonicaldatatypes array (of VT_I4)
		res = VariantInitArrayOf( pCanonicalDataTypes, NumItems, VT_I2 );
		if( FAILED(res) ) {
			goto IMDAddItemsBadExit3;
		}
	}


	if( ( pAccessRights != NULL ) && ( V_VT( pAccessRights ) != VT_NULL ) ) {
		// create access rights array (of VT_I2 *)
		res = VariantInitArrayOf( pAccessRights, NumItems, VT_I4 );
		if( FAILED(res) ) {
			goto IMDAddItemsBadExit4;
		}
	}


	// !!!??? here assuming a Blob is a Variant representing
	// an array of VT_UI1 (that is a byte stream)


	ppDItems = new DaDeviceItem*[ NumItems ];
	if( ppDItems == NULL ) {
		goto IMDAddItemsBadExit4b;
	}

	pItemArray = new OPCITEMDEF[ NumItems ];
	if( pItemArray == NULL ) {
		goto IMDAddItemsBadExit6;
	}

	pTempErrors = new HRESULT[ NumItems ];
	if( pItemArray == NULL ) {
		goto IMDAddItemsBadExit6b;
	}


	// fill the pItemArray with VARIANT arrays passed as input
	i = 0;
	while ( i < NumItems ) {

		pTempErrors[i] = S_OK;
		ppDItems[i]    = NULL;

		SafeArrayGetElement( V_ARRAY( &ItemIDs ), &i, &pItemArray[i].szItemID );
		SafeArrayGetElement( V_ARRAY( &ActiveStates ), &i, &bActive );

		if( bActive == VARIANT_FALSE ) {
			pItemArray[i].bActive = FALSE;
		} else {
			pItemArray[i].bActive = TRUE;
		}

		SafeArrayGetElement( V_ARRAY( &ClientHandles ), &i, &(pItemArray[i].hClient) );
		if( V_VT( &AccessPaths ) != VT_NULL ) {
			SafeArrayGetElement( V_ARRAY( &AccessPaths ), &i, &(pItemArray[i].szAccessPath) );
		} else {
			pItemArray[i].szAccessPath = SysAllocString(L"");
		}

		if( V_VT(&RequestedDataTypes) != VT_NULL ) {
			SafeArrayGetElement( V_ARRAY( &RequestedDataTypes ), &i, &(pItemArray[i].vtRequestedDataType) );
		} else {
			pItemArray[i].vtRequestedDataType = VT_EMPTY;
		}

		if( ( pBlobs != NULL ) & ( V_VT( pBlobs ) != VT_NULL ) ) {
			// there is an input

			// get the BLOB 
			SafeArrayGetElement( V_ARRAY( pBlobs ), &i, &theBlobVar );

			ConvertBlobVariantToBytes( &theBlobVar, &(pItemArray[i].dwBlobSize), &(pItemArray[i].pBlob), NULL );

			//AM
			VariantClear( &theBlobVar );


		} else {

			pItemArray[i].pBlob = NULL;
			pItemArray[i].dwBlobSize = 0;
		}

		i++;
	}

	// check the item objects
	theRes = m_pServerHandler->OnValidateItems(OPC_VALIDATEREQ_DEVICEITEMS,
		TRUE,   // Blob Update
		NumItems,
		pItemArray,
		ppDItems,
		NULL,
		pTempErrors );
	if( FAILED( theRes ) ) {
		goto IMDAddItemsBadExit7;
	}


	EnterCriticalSection( &(group->m_ItemsCritSec) );
	// insert the successfully items in the groups list
	// and itemresult datastructures
	i = 0;
	while ( i < NumItems ) {
		if( SUCCEEDED( pTempErrors[i] ) ) {
			if (ppDItems[i]->Killed()) {
				ppDItems[i]->Detach();              // Attach from OnValidateItems() no longer required
				pTempErrors[i] = OPC_E_UNKNOWNITEMID;
			}
		}
		if( SUCCEEDED( pTempErrors[i] ) ) {


			// OPCITEMDEF struct
			pID = &pItemArray[i];

			res = S_OK;

			// construct item
			NewGI = new DaGenericItem( );
			if( NewGI == NULL ) {
				res = E_OUTOFMEMORY;
				goto IMDAddItemsBadExitX;
			}

			// create generic item
			res = NewGI->Create( pID->bActive,           // active mode flag
				pID->hClient,           // client handle
				pID->vtRequestedDataType,
				group,                  // this group
				ppDItems[i],            // link to DeviceItem
				&hServerHandle,
				FALSE );                // Since CALL-R has no Automation
			// Interface an Item cannot be an
			// Item with Physival value.
			if( FAILED(res) ) {
				goto IMDAddItemsBadExitX;
			}


			// Fill result arrays
			SafeArrayPutElement( V_ARRAY( pServerHandles ), &i, &hServerHandle);
			if( ( pCanonicalDataTypes != NULL ) & ( V_VT( pCanonicalDataTypes ) != VT_NULL ) ) {
				VARTYPE         vt;
				vt = NewGI->get_CanonicalDataType();
				SafeArrayPutElement( V_ARRAY( pCanonicalDataTypes ), &i, &vt );
			}
			if( ( pAccessRights != NULL ) && ( V_VT( pAccessRights ) != VT_NULL ) ) { 
				DWORD ar;
				NewGI->get_AccessRights( &ar );
				// !!!??? check for error
				SafeArrayPutElement( V_ARRAY( pAccessRights ), &i, &ar );
			}
			if( ( pBlobs != NULL ) && ( V_VT( pBlobs ) != VT_NULL ) ) {
				// there is an output
				NewGI->get_Blob( &pBlob, &BlobSize );

				// !!!??? check for errors!
				ConvertBlobBytesToVariant(    
					BlobSize, 
					pBlob, 
					&pTheBlobVar, 
					pIMalloc );


				SafeArrayPutElement( V_ARRAY( pBlobs ), &i, pTheBlobVar );

			}

			//Get or Create the COM item and set the pItemObjects array
			res = group->GetCOMItem( hServerHandle, NewGI, &pCOMItem, &created);
			if( FAILED(res)) {
				goto IMDAddItemsBadExitX;
			}

			// QueryInterface seems not to be necessary
			// it is done outside!
			SafeArrayPutElement( V_ARRAY( pItemObjects ), &i, pCOMItem); 


			// used to skip
IMDAddItemsBadExitX:
			ppDItems[i]->Detach();              // Attach from OnValidateItems() no longer required
			; // keep this!

		} else { // erroneous item
			// leave the values as they were passed as parameter
			res = pTempErrors[i];
		}

		SafeArrayPutElement( V_ARRAY( errors ), &i, &res );

		if( FAILED( res ) ) {
			theRes = S_FALSE;
		}

		// Release all in OPCITEMDEF struct
		pID = &pItemArray[i];
		if (pID->szItemID)      SysFreeString( pID->szItemID );
		if (pID->szAccessPath)  SysFreeString( pID->szAccessPath );
		if (pID->pBlob)         delete [] pBlob;
		i++;
	}

	LeaveCriticalSection( &group->m_ItemsCritSec );

	// unnail the group
	ReleaseGenericGroup();

	delete [] ppDItems;
	delete [] pItemArray;
	delete [] pTempErrors;

	return theRes;


IMDAddItemsBadExit7:
	delete [] pTempErrors;

IMDAddItemsBadExit6b:
	delete [] pItemArray;

IMDAddItemsBadExit6:

	delete [] ppDItems;

IMDAddItemsBadExit4b:
	if( ( pAccessRights != NULL ) && ( V_VT( pAccessRights ) != VT_NULL ) ) {
		VariantUninitArrayOf( pAccessRights );
	}

IMDAddItemsBadExit4:
	if( ( pCanonicalDataTypes != NULL ) && ( V_VT( pCanonicalDataTypes ) != VT_NULL ) ) {
		VariantUninitArrayOf( pCanonicalDataTypes );
	}

IMDAddItemsBadExit3:
	VariantUninitArrayOf( pServerHandles );


IMDAddItemsBadExit2:
	VariantUninitArrayOf( pItemObjects );

IMDAddItemsBadExit1:
	VariantUninitArrayOf( errors );

IMDAddItemsBadExit0:
	ReleaseGenericGroup();

IMDAddItemsBadExit00:
	return res;
}




//=========================================================================
// IOPCItemMgtDisp.ValidateItems
// -----------------------------
// Check if the specified items ccould be added to the group.
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::ValidateItems(
	/*[in]               */ long      NumItems,
	/*[in]               */ VARIANT   ItemIDs,
	/*[out]              */ VARIANT * errors,
	/*[in, optional]     */ VARIANT   AccessPaths,
	/*[in, optional]     */ VARIANT   RequestedDataTypes,
	/*[in, optional]     */ VARIANT BlobUpdate,
	/*[in, out, optional]*/ VARIANT * pBlobs,
	/*[out, optional]    */ VARIANT * pCanonicalDataTypes,
	/*[out, optional]    */ VARIANT * pAccessRights
	)
{
	HRESULT          res;
	long             i;
	OPCITEMDEF     * pItemArray;
	HRESULT        * pTempErrors;
	OPCITEMRESULT  * pTempItemResults;
	DaGenericGroup  * group;
	OPCITEMDEF     *pID;
	OPCITEMRESULT  *pItemResult;
	HRESULT         theRes;
	VARIANT         theBlobVar;   
	VARIANT         *pTheBlobVar; 
	BOOL            BlobsUpdated;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCItemMgtDisp.ValidateItems: Internal error: No generic group.");
		goto IMDValidateItemsBadExit00;
	}
   OPCWSTOAS( group->m_Name )
      LOGFMTI("IOPCItemMgtDisp.ValidateItems: Validate %ld items to group %s", NumItems, OPCastr );
   }

	if( errors != NULL ) {
		VariantInit( errors );
	}

	if(( NumItems <= 0 )
		|| ( V_VT(&ItemIDs) != (VT_ARRAY | VT_BSTR) ) 
		|| ( (V_VT(&AccessPaths) != (VT_ARRAY | VT_BSTR ) ) && (V_VT(&AccessPaths) != VT_NULL) ) 
		|| ( (V_VT(&RequestedDataTypes) != ( VT_ARRAY | VT_I2) ) && (V_VT(&RequestedDataTypes) != VT_NULL) ) 
		|| ( (V_VT(&BlobUpdate) != VT_NULL) && (V_VT(&BlobUpdate) != VT_BOOL) ) 
		|| ( errors == NULL )  ) {

         LOGFMTE( "ValidateItems() failed with invalid argument(s):" );
         if( NumItems <= 0 ) {
            LOGFMTE( "      NumItems is %ld", NumItems );
         }
         if( V_VT( &ItemIDs ) != (VT_ARRAY | VT_BSTR) ) {
            LOGFMTE( "      ItemIDs is not array of VT_BSTR" );
         }
         if( (V_VT( &AccessPaths ) != (VT_ARRAY | VT_BSTR ) ) && (V_VT( &AccessPaths ) != VT_NULL) ) {
            LOGFMTE( "      AccessPaths is not VT_NULL or array of VT_BSTR" );
         }
         if( (V_VT( &RequestedDataTypes ) != ( VT_ARRAY | VT_I2) ) && (V_VT( &RequestedDataTypes ) != VT_NULL) ) {
            LOGFMTE( "      RequestedDataTypes is not VT_NULL or array of VT_I2" );
         }
         if( (V_VT( &BlobUpdate ) != VT_NULL) && (V_VT( &BlobUpdate ) != VT_BOOL) ) {
            LOGFMTE( "      pBlobs is not VT_NULL or VT_BOOL" );
         }
         if( errors == NULL ) {
            LOGFMTE( "      errors is NULL" );
         }

			res = E_INVALIDARG;
			goto IMDValidateItemsBadExit0;
	}

	// dont check for public group!
	//   //  works also for public groups?
	//if( group->GetPublicInfo( &pgh ) == TRUE ) {
	//   res = OPC_E_PUBLIC;
	//   goto IMDValidateItemsBadExit0;
	//}

	if( V_VT(&BlobUpdate) != VT_NULL ) {
		if( (V_BOOL( &BlobUpdate ) != VARIANT_FALSE) ) {
			// Update blobs is True
			if( (pBlobs == NULL) || (V_VT( pBlobs ) != (VT_ARRAY | VT_VARIANT)) ) {
				// cannot update if no blobs are passed as inputs
				res = E_INVALIDARG;
				goto IMDValidateItemsBadExit0;
			}

			BlobsUpdated = TRUE;
		} else {
			BlobsUpdated = FALSE;
		}
	} else { // VT_NULL
		BlobsUpdated = FALSE;
	}

	if( BlobsUpdated == FALSE ) {
		// only return the actual blobs! (no inputs)
		if( ( pBlobs != NULL ) && ( V_VT( pBlobs )!= VT_NULL ) ) {
			// client is interested in the blobs (outputs)
			res = VariantInitArrayOf( pBlobs, NumItems, VT_VARIANT );
			if( FAILED(res) ) {
				goto IMDValidateItemsBadExit0;
			}
		}
	}



	// all output VARIANTs must be created with
	// COM memory management object

	// create errors array ( of HRESULT)
	res = VariantInitArrayOf( errors, NumItems, VT_I4 );
	if( FAILED(res) ) {
		goto IMDValidateItemsBadExit1;
	}
	// initialize with E_FAIL
	{
		long *temp;
		long i;

		temp = (long *)(V_ARRAY( errors )->pvData);

		i = 0;
		while ( i < NumItems ) {
			temp[i] = E_FAIL;
			i++;
		}
	}

	if( ( pCanonicalDataTypes != NULL ) && ( V_VT( pCanonicalDataTypes ) != VT_NULL ) ) {
		// create canonicaldatatypes array (of VT_I4)
		res = VariantInitArrayOf( pCanonicalDataTypes, NumItems, VT_I2 );
		if( FAILED(res) ) {
			goto IMDValidateItemsBadExit3;
		}
	}

	if( (pAccessRights != NULL) && ( V_VT( pAccessRights ) != VT_NULL ) ) {
		// create access rights array (of VT_I2 *)
		res = VariantInitArrayOf( pAccessRights, NumItems, VT_I4 );
		if( FAILED(res) ) {
			goto IMDValidateItemsBadExit4;
		}
	}


	// !!!??? here assuming a Blob is a Variant representing
	// an array of VT_UI1 (that is a byte stream)


	res = E_OUTOFMEMORY;
	pItemArray = new OPCITEMDEF[ NumItems ];
	if( pItemArray == NULL ) {
		goto IMDValidateItemsBadExit6;
	}

	pTempErrors = new HRESULT[ NumItems ];
	if( pItemArray == NULL ) {
		goto IMDValidateItemsBadExit6b;
	}

	pTempItemResults = new OPCITEMRESULT [NumItems];
	if (!pTempItemResults) {
		goto IMDValidateItemsBadExit6c;
	}
	res = S_OK;

	// fill the pItemArray with VARIANT arrays passed as input
	for (i=0; i < NumItems; i++) {

		pTempErrors[i] = S_OK;

		pTempItemResults[i].hServer               = 0;
		pTempItemResults[i].vtCanonicalDataType   = VT_EMPTY;
		pTempItemResults[i].wReserved             = 0;
		pTempItemResults[i].dwAccessRights        = 0;
		pTempItemResults[i].dwBlobSize            = 0;
		pTempItemResults[i].pBlob                 = NULL;

		SafeArrayGetElement( V_ARRAY( &ItemIDs ), &i, &pItemArray[i].szItemID );

		if( V_VT(&RequestedDataTypes) != VT_NULL ) {
			SafeArrayGetElement( V_ARRAY( &RequestedDataTypes ), &i, &pItemArray[i].vtRequestedDataType );
		} else {
			pItemArray[i].vtRequestedDataType = VT_EMPTY;
		}

		if( (BlobsUpdated == TRUE) && (V_VT( pBlobs )!= VT_NULL) ){
			// get the BLOB if there is one as input
			SafeArrayGetElement( V_ARRAY( pBlobs ), &i, &theBlobVar );

			ConvertBlobVariantToBytes( &theBlobVar, &pItemArray[i].dwBlobSize, &pItemArray[i].pBlob, NULL );

		} else {

			pItemArray[i].pBlob = NULL;
			pItemArray[i].dwBlobSize = 0;
		}
	}

	// check the item objects
	theRes = m_pServerHandler->OnValidateItems(
		OPC_VALIDATEREQ_ITEMRESULTS,
		BlobsUpdated,
		NumItems, 
		pItemArray, 
		NULL,
		pTempItemResults,
		pTempErrors );
	if( FAILED( theRes ) ) {
		res = theRes;
		goto IMDValidateItemsBadExit7;
	}


	for (i=0; i < NumItems; i++) {
		if (SUCCEEDED( pTempErrors[i] )) {

			pID = &pItemArray[i] ;                 // OPCITEMDEF struct
			pItemResult = &pTempItemResults[i];    // OPCITEMRESULT struct

			if( ( pCanonicalDataTypes != NULL ) && ( V_VT( pCanonicalDataTypes )!= VT_NULL ) ) {
				SafeArrayPutElement( V_ARRAY( pCanonicalDataTypes ), &i, &pItemResult->vtCanonicalDataType );
			}
			if( (pAccessRights != NULL) && ( V_VT( pAccessRights ) != VT_NULL ) ) { 
				SafeArrayPutElement( V_ARRAY( pAccessRights ), &i, &pItemResult->dwAccessRights );
			}
			if( ( pBlobs != NULL ) && (V_VT( pBlobs ) != VT_NULL ) ) {

				ConvertBlobBytesToVariant(          // Now recreate the arrays of blobs!
					pItemResult->dwBlobSize, 
					pItemResult->pBlob, 
					&pTheBlobVar, 
					pIMalloc );

				SafeArrayPutElement( V_ARRAY( pBlobs ), &i, pTheBlobVar );
			}
		}
		else {
			// invalid item
			// leave them as they were passed as parameter
			theRes = S_FALSE;
		}
		SafeArrayPutElement( V_ARRAY( errors ), &i, &pTempErrors[i] );
	}

	ReleaseGenericGroup();

	for (i=0; i < NumItems; i++) {
		// Release pItemArray members
		SysFreeString( pItemArray[i].szItemID );
		if (pItemArray[i].pBlob) {
			delete [] pItemArray[i].pBlob;
		}

		// Release pTempItemResults members
		if (pTempItemResults[i].pBlob) {
			pIMalloc->Free( pTempItemResults[i].pBlob );
		}
	}
	delete [] pItemArray;
	delete [] pTempErrors;
	delete [] pTempItemResults;

	return theRes;


IMDValidateItemsBadExit7:
	for (i=0; i < NumItems; i++) {
		// Release pItemArray members
		SysFreeString( pItemArray[i].szItemID );
		if (pItemArray[i].pBlob) {
			delete [] pItemArray[i].pBlob;
		}

		// Release pTempItemResults members
		if (pTempItemResults[i].pBlob) {
			pIMalloc->Free( pTempItemResults[i].pBlob );
		}
	}
	delete pTempItemResults;

IMDValidateItemsBadExit6c:
	delete pTempErrors;

IMDValidateItemsBadExit6b:
	delete pItemArray;

IMDValidateItemsBadExit6:
	if( (pAccessRights != NULL) && ( V_VT( pAccessRights ) != VT_NULL ) ) { 
		VariantUninitArrayOf( pAccessRights );
	}

IMDValidateItemsBadExit4:
	if( ( pCanonicalDataTypes != NULL ) && ( V_VT( pCanonicalDataTypes ) != VT_NULL ) ) {
		VariantUninitArrayOf( pCanonicalDataTypes );
	}

IMDValidateItemsBadExit3:
	VariantUninitArrayOf( errors );

IMDValidateItemsBadExit1:
	if( BlobsUpdated == FALSE ) {
		// no inputs
		if( ( pBlobs != NULL ) && ( V_VT( pBlobs ) != VT_NULL ) ) {
			// wanted outputs
			/*res =*/ VariantUninitArrayOf( pBlobs );
		}
	}

IMDValidateItemsBadExit0:
	ReleaseGenericGroup();

IMDValidateItemsBadExit00:
	return res;
}




//=========================================================================
// IOPCItemMgtDisp.RemoveItems
// ---------------------------
// Items are removed even though some operation is performed over them!
// So the client must care to avoid access to them while removing
// 
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::RemoveItems( 
	/*[in] */ long      NumItems,
	/*[in] */ VARIANT   ServerHandles,
	/*[out]*/ VARIANT * errors,
	/*[in] */ VARIANT_BOOL   Force  )
{
	long                    sh;
	HRESULT                 res;
	BOOL                    SomeBadHandle;
	long                    i;
	DaGenericGroup           *group;
	long                    pgh;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCItemMgtDisp.RemoveItems: Internal error: No generic group.");
		goto IMDRemoveItemsExit0;
	}
   OPCWSTOAS( group->m_Name )
      LOGFMTI("IOPCItemMgtDisp.RemoveItems: Remove %ld items to group %s", NumItems, OPCastr );
   }

	if( errors != NULL ) {
		VariantInit( errors );
	}
	if( ( NumItems <= 0 ) || (errors == NULL) ) {
         LOGFMTE( "RemoveItems() failed with invalid argument(s):" );
         if( NumItems <= 0 ) {
            LOGFMTE( "      NumItems is %ld", NumItems );
         }
         if( errors == NULL ) {
            LOGFMTE( "      errors is NULL" );
         }
		res = E_INVALIDARG;
		goto IMDRemoveItemsExit1;
	}

	if( group->GetPublicInfo( &pgh ) == TRUE ) {
		res = OPC_E_PUBLIC;
		goto IMDRemoveItemsExit1;
	}

	if(V_VT( &ServerHandles) != (VT_ARRAY | VT_I4) ) {
		res = E_INVALIDARG;
		goto IMDRemoveItemsExit1;
	}

	res = VariantInitArrayOf( errors, NumItems, VT_I4 ); 
	if( FAILED(res) ) {
		goto IMDRemoveItemsExit1;
	}

	EnterCriticalSection( &(group->m_ItemsCritSec) );

	SomeBadHandle = FALSE;
	i = 0;
	while ( i < NumItems ) {
		// get the server handle from input
		SafeArrayGetElement( V_ARRAY( &ServerHandles ), &i, &sh );

		res = group->RemoveGenericItemNoLock( sh );
		if( FAILED(res) ) {
			SomeBadHandle = TRUE;
		}
		SafeArrayPutElement( V_ARRAY( errors ), &i, &res );
		i++;
	}

	LeaveCriticalSection( &(group->m_ItemsCritSec) );

	if( SomeBadHandle == TRUE ) {
		res = S_FALSE;
	} else {
		res = S_OK;
	}

IMDRemoveItemsExit1:
	ReleaseGenericGroup();

IMDRemoveItemsExit0:
	return res;
}




//=========================================================================
// IOPCItemMgtDisp.SetActiveState
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::SetActiveState(
	/*[in] */ long      NumItems,
	/*[in] */ VARIANT   ServerHandles,
	/*[in] */ VARIANT_BOOL   ActiveState,
	/*[out]*/ VARIANT * errors  )
{
	long           sh;
	HRESULT        res;
	DaGenericItem   *pItem;
	BOOL           SomeBadHandle;
	long           i;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
		LOGFMTE("IOPCItemMgtDisp.SetActiveState: Internal error: No generic group.");
		goto IMDSetActiveStateExit0;
	}
   OPCWSTOAS( group->m_Name )
      LOGFMTI("IOPCItemMgtDisp.SetActiveState of %ld items in group %s", NumItems, OPCastr );
   }

	if( errors != NULL ) {
		VariantInit( errors );
	}

	if( (errors == NULL) || (NumItems <= 0) ) {

         LOGFMTE( "SetActiveState() failed with invalid argument(s):" );
         if( errors == NULL ) {
            LOGFMTE( "      errors is NULL" );
         }
         if( NumItems <= 0 ) {
            LOGFMTE( "      NumItems is %ld", NumItems );
         }

		res = E_INVALIDARG;
		goto IMDSetActiveStateExit1;
	}

	if( V_VT(&ServerHandles) != (VT_ARRAY | VT_I4) ) {
		res = E_INVALIDARG;
		goto IMDSetActiveStateExit1;

	}

	res = VariantInitArrayOf( errors, NumItems, VT_I4 ); 
	if( FAILED(res) ) {
		goto IMDSetActiveStateExit1;
	}

	SomeBadHandle = FALSE;
	i = 0;
	while ( i < NumItems ) {
		// get the server handle from input
		SafeArrayGetElement( V_ARRAY( &ServerHandles ), &i, &sh );

		// get and lock item
		res = group->GetGenericItem( sh, &pItem );
		if( FAILED(res) ) {
			SomeBadHandle = TRUE;
		} else {
			if( ActiveState == VARIANT_FALSE ) {
				res = pItem->set_Active( FALSE );
			} else {
				res = pItem->set_Active( TRUE );
			}
			if( FAILED(res) ) {
				SomeBadHandle = TRUE;
			}
			group->ReleaseGenericItem( sh );
		}

		SafeArrayPutElement( V_ARRAY( errors ), &i, &res );
		i++;
	}

	if( SomeBadHandle == TRUE ) {
		res = S_FALSE;
	} else {
		res = S_OK;
	}

IMDSetActiveStateExit1:
	ReleaseGenericGroup();

IMDSetActiveStateExit0:
	return res;
}



//=========================================================================
// IOPCItemMgtDisp.SetClientHandles
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::SetClientHandles(
	/*[in] */ long      NumItems,
	/*[in] */ VARIANT   ServerHandles,
	/*[in] */ VARIANT   ClientHandles,
	/*[out]*/ VARIANT * errors  )
{
	long           sh;
	long           ch;
	HRESULT        res;
	DaGenericItem   *pItem;
	BOOL           SomeBadHandle;
	long           i;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCItemMgtDisp.SetClientHandles: Internal error: No generic group.");
		goto IMDSetClientHandlesExit0;
	}
   OPCWSTOAS( group->m_Name )
      LOGFMTI("IOPCItemMgtDisp.SetClientHandles of %ld items in group %s", NumItems, OPCastr );
   }

	if( errors != NULL ) {
		VariantInit( errors );
	}

	if( (errors == NULL) || (NumItems <= 0) ) {

         LOGFMTE( "SetActiveState() failed with invalid argument(s):" );
         if( errors == NULL ) {
            LOGFMTE( "      errors is NULL" );
         }
         if( NumItems <= 0 ) {
            LOGFMTE( "      NumItems is %ld", NumItems );
         }

		res = E_INVALIDARG;
		goto IMDSetClientHandlesExit1;
	}
	if( V_VT(&ServerHandles) != (VT_ARRAY | VT_I4) ) {
		res = E_INVALIDARG;
		goto IMDSetClientHandlesExit1;
	}
	if( V_VT(&ClientHandles) != (VT_ARRAY | VT_I4) ) {
		res = E_INVALIDARG;
		goto IMDSetClientHandlesExit1;
	}

	res = VariantInitArrayOf( errors, NumItems, VT_I4 ); 
	if( FAILED(res) ) {
		goto IMDSetClientHandlesExit1;
	}

	SomeBadHandle = FALSE;
	i = 0;
	while ( i < NumItems ) {
		// get the server handle from input
		SafeArrayGetElement( V_ARRAY( &ServerHandles ), &i, &sh );
		// get the new client handle from input
		SafeArrayGetElement( V_ARRAY( &ClientHandles ), &i, &ch );

		// get and lock item
		res = group->GetGenericItem( sh, &pItem );
		if( FAILED(res) ) {
			SomeBadHandle = TRUE;
		} else {
			pItem->set_ClientHandle( ch );
			group->ReleaseGenericItem( sh );
		}

		SafeArrayPutElement( V_ARRAY( errors ), &i, &res );
		i++;
	}

	if( SomeBadHandle == TRUE ) {
		res = S_FALSE;
	} else {
		res = S_OK;
	}

IMDSetClientHandlesExit1:
	ReleaseGenericGroup();

IMDSetClientHandlesExit0:
	return res;
}



//=========================================================================
// IOPCItemMgtDisp.SetDatatypes
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::SetDatatypes(
	/*[in] */ long      NumItems,
	/*[in] */ VARIANT   ServerHandles,
	/*[in] */ VARIANT   RequestedDataTypes,    
	/*[out]*/ VARIANT * errors       )
{
	long           sh;
	VARTYPE        rdt;
	HRESULT        res;
	DaGenericItem   *pItem;
	BOOL           SomeBadHandle;
	long           i;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
		LOGFMTE("IOPCItemMgtDisp.SetDataTypes: Internal error: No generic group.");
		goto IMDSetDataTypesExit0;
	}
   OPCWSTOAS( group->m_Name )
      LOGFMTI("IOPCItemMgtDisp.SetDataTypes of %ld items in group %s", NumItems, OPCastr );
   }

	if( errors != NULL ) {
		VariantInit( errors );
	}

	if( (errors == NULL) || (NumItems <= 0) ) {

         LOGFMTE( "SetDatatypes() failed with invalid argument(s):" );
         if( errors == NULL ) {
            LOGFMTE( "      errors is NULL" );
         }
         if( NumItems <= 0 ) {
            LOGFMTE( "      NumItems is %ld", NumItems );
         }
		res = E_INVALIDARG;
		goto IMDSetDataTypesExit1;
	}

	if( V_VT(&ServerHandles) != (VT_ARRAY | VT_I4) ) {
		res = E_INVALIDARG;
		goto IMDSetDataTypesExit1;
	}
	if( V_VT(&RequestedDataTypes) != (VT_ARRAY | VT_I2) ) {
		res = E_INVALIDARG;
		goto IMDSetDataTypesExit1;
	}

	res = VariantInitArrayOf( errors, NumItems, VT_I4 ); 
	if( FAILED(res) ) {
		goto IMDSetDataTypesExit1;
	}

	SomeBadHandle = FALSE;
	i = 0;
	while ( i < NumItems ) {
		// get the server handle from input
		SafeArrayGetElement( V_ARRAY( &ServerHandles ), &i, &sh );
		// get the new client handle from input
		SafeArrayGetElement( V_ARRAY( &RequestedDataTypes ), &i, &rdt );

		// get and lock item
		res = group->GetGenericItem( sh, &pItem );
		if( FAILED(res) ) {
			SomeBadHandle = TRUE;
		} else {
			res = pItem->set_RequestedDataType( rdt );
			if( FAILED(res) ) {
				SomeBadHandle = TRUE;
			}
			group->ReleaseGenericItem( sh );
		}

		SafeArrayPutElement( V_ARRAY( errors ), &i, &res );
		i++;
	}

	if( SomeBadHandle == TRUE ) {
		res = S_FALSE;
	} else {
		res = S_OK;
	}

IMDSetDataTypesExit1:
	ReleaseGenericGroup();

IMDSetDataTypesExit0:
	return res;
}






//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[ IOPCGroupStateMgtDisp [[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

//=========================================================================
// IOPCGroupStateMgtDisp.get_ActiveStatus
//=========================================================================
////[ propget ]
HRESULT STDMETHODCALLTYPE DaGroup::get_ActiveStatus(
	//        /*[out, retval]*/ boolean * pActiveStatus
	/*[out, retval]*/ VARIANT_BOOL * pActiveStatus )
{
	DaGenericGroup  *group;
	HRESULT        res;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.get_ActiveStatus: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.get_ActiveStatus of group %s", W2A(group->m_Name));

	if( pActiveStatus == NULL ) {
      LOGFMTE( "get_ActiveStatus() failed with invalid argument(s):  pActiveStatus is NULL" );
		ReleaseGenericGroup();
		return E_INVALIDARG;
	}

	if( group->GetActiveState() == FALSE ) {
		*pActiveStatus = VARIANT_FALSE;  
	} else {
		*pActiveStatus = VARIANT_TRUE; 
	}

	ReleaseGenericGroup();
	return S_OK;
}




//=========================================================================
// IOPCGroupStateMgtDisp.put_ActiveStatus
//
// if Active set to false then all callbacks are stopped!
//
//=========================================================================
////[ propput ]
HRESULT STDMETHODCALLTYPE DaGroup::put_ActiveStatus(
	/*[in]*/ VARIANT_BOOL ActiveStatus  )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.put_ActiveStatus: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.put_ActiveStatus of group %s", W2A(group->m_Name));

	if( ActiveStatus == VARIANT_FALSE ) {
		res = group->SetActiveState( FALSE );
	} else {                      
		res = group->SetActiveState( TRUE );
	}

	ReleaseGenericGroup();
	return res;
}




//=========================================================================
// IOPCGroupStateMgtDisp.get_ClientGroupHandle
//
//=========================================================================
////[ propget ]
HRESULT STDMETHODCALLTYPE DaGroup::get_ClientGroupHandle(
	/*[out, retval]*/ long * phClientGroupHandle  )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.get_ClientGroupHandle: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.get_ClientGroupHandle of group %s", W2A(group->m_Name) );

	if (phClientGroupHandle == NULL) {
      LOGFMTE( "get_ClientGroupHandle() failed with invalid argument(s):  phClientGroupHandle is NULL" );
		res = E_INVALIDARG;
	}
	else {
		*phClientGroupHandle = group->m_hClientGroupHandle;
		res = S_OK;
	}

	ReleaseGenericGroup();
	return res;
}




//=========================================================================
// IOPCGroupStateMgtDisp.put_ClientGroupHandle
//
//=========================================================================
////[ propput ]
HRESULT STDMETHODCALLTYPE DaGroup::put_ClientGroupHandle(
	/*[in]*/ long ClientGroupHandle  )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.put_ClientGroupHandle: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.put_ClientGroupHandle of group %s", W2A(group->m_Name));

	group->m_hClientGroupHandle = ClientGroupHandle;
	ReleaseGenericGroup();
	return S_OK;
}





//=========================================================================
// IOPCGroupStateMgtDisp.get_ServerGroupHandle
//
//=========================================================================
////[ propget ]
HRESULT STDMETHODCALLTYPE DaGroup::get_ServerGroupHandle(
	/*[out, retval]*/ long * phServerGroupHandle  )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.get_ServerGroupHandle: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.get_ServerGroupHandle of group %s", W2A(group->m_Name));

	if (phServerGroupHandle == NULL) {
      LOGFMTE( "get_ServerGroupHandle() failed with invalid argument(s):  phServerGroupHandle is NULL" );
		res = E_INVALIDARG;
	}
	else {
		*phServerGroupHandle = group->m_hServerGroupHandle;
		res = S_OK;
	}

	ReleaseGenericGroup();
	return res;
}




//=========================================================================
// IOPCGroupStateMgtDisp.get_Name
//
//=========================================================================
////[ propget ]
HRESULT STDMETHODCALLTYPE DaGroup::get_Name(
	/*[out, retval]*/ BSTR * pName )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.get_Name: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.get_Name of group %s", W2A(group->m_Name));

	if( pName == NULL ) {
      LOGFMTE( "get_Name() failed with invalid argument(s):  pName is NULL" );
		ReleaseGenericGroup();
		return E_INVALIDARG;
	}

	// lock the private group name
	// while creating the public group
	EnterCriticalSection ( &(m_pServer->m_GroupsCritSec) );

	*pName = SysAllocString( group->m_Name );
	// unlock the group->name_
	LeaveCriticalSection ( &(m_pServer->m_GroupsCritSec) );

	if( *pName == NULL ) {
		res = E_OUTOFMEMORY;
	} else {
		res = S_OK;
	}
	ReleaseGenericGroup();
	return res;
}




//=========================================================================
// IOPCGroupStateMgtDisp.put_Name
//  
//=========================================================================
////[ propput ]
HRESULT STDMETHODCALLTYPE DaGroup::put_Name(
	/*[in]*/ BSTR Name )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.put_Name: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.put_Name of group %s", W2A(group->m_Name));

	res = m_pServer->ChangePrivateGroupName( group, Name );
	ReleaseGenericGroup();
	return res;
}




//=========================================================================
// IOPCGroupStateMgtDisp.get_UpdateRate
// 
//=========================================================================
////[ propget ]
HRESULT STDMETHODCALLTYPE DaGroup::get_UpdateRate(
	/*[out, retval]*/ long * pUpdateRate  )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.get_UpdateRate: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.get_UpdateRate of group %s", W2A(group->m_Name));

	if (pUpdateRate == NULL) {
      LOGFMTE( "get_UpdateRate() failed with invalid argument(s):  pUpdateRate is NULL" );
		res = E_INVALIDARG;
	}
	else {
		*pUpdateRate = group->m_RevisedUpdateRate;
		res = S_OK;
	}

	ReleaseGenericGroup();
	return res;
}




//=========================================================================
// IOPCGroupStateMgtDisp.put_UpdateRate
// 
//=========================================================================
////[ propput ]
HRESULT STDMETHODCALLTYPE DaGroup::put_UpdateRate(
	/*[in]*/ long UpdateRate  )
{
	HRESULT        res;
	DaGenericGroup *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.put_UpdateRate: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.put_UpdateRate of group %s", W2A(group->m_Name));

	group->ReviseUpdateRate( UpdateRate );
	ReleaseGenericGroup();
	return S_OK;
}





//=========================================================================
// IOPCGroupStateMgtDisp.get_TimeBias
// 
//=========================================================================
////[ propget ]
HRESULT STDMETHODCALLTYPE DaGroup::get_TimeBias(
	/*[out, retval]*/ long * pTimeBias  )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.get_TimeBias: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.get_TimeBias of group %s", W2A(group->m_Name));

	if (pTimeBias == NULL) {
      LOGFMTE( "get_TimeBias() failed with invalid argument(s): pTimeBias is NULL" );
		res = E_INVALIDARG;
	}
	else {
		*pTimeBias = group->m_TimeBias;
		res = S_OK;
	}

	ReleaseGenericGroup();
	return res;
}




//=========================================================================
// IOPCGroupStateMgtDisp.put_TimeBias
// 
//=========================================================================
////[ propput ]
HRESULT STDMETHODCALLTYPE DaGroup::put_TimeBias(
	/*[in]*/ long TimeBias  )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.put_TimeBias: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.put_TimeBias of group %s", W2A(group->m_Name));

	group->m_TimeBias = TimeBias;
	ReleaseGenericGroup();
	return S_OK;
}




//=========================================================================
// IOPCGroupStateMgtDisp.get_PercentDeadBand
// 
//=========================================================================
////[ propget ]
HRESULT STDMETHODCALLTYPE DaGroup::get_PercentDeadBand(
	/*[out, retval]*/ float * pPercentDeadBand  )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.get_PercentDeadBand: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.get_PercentDeadBand of group %s", W2A(group->m_Name));

	if (pPercentDeadBand == NULL) {
      LOGFMTE( "get_PercentDeadBand() failed with invalid argument(s):   pPercentDeadBand is NULL" );
		res = E_INVALIDARG;
	}
	else {
		*pPercentDeadBand = group->m_PercentDeadband;
		res = S_OK;
	}

	ReleaseGenericGroup();
	return res;
}





//=========================================================================
// IOPCGroupStateMgtDisp.put_PercentDeadBand
// 
//=========================================================================
////[ propput ]
HRESULT STDMETHODCALLTYPE DaGroup::put_PercentDeadBand(
	/*[in]*/ float PercentDeadBand  )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.put_PercentDeadBand: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.put_PercentDeadBand of group %s", W2A(group->m_Name));

	if (PercentDeadBand < 0.0 || PercentDeadBand > 100.0) {
      LOGFMTE( "put_PercentDeadBand() failed with invalid argument(s): Percent Deadband out of range (0...100)" );
		res = E_INVALIDARG;
	}
	else {
		group->m_PercentDeadband = PercentDeadBand;
		res = S_OK;
	}
	ReleaseGenericGroup();
	return res;
}





//=========================================================================
// IOPCGroupStateMgtDisp.get_LCID
// 
//=========================================================================
////[ propget ]
HRESULT STDMETHODCALLTYPE DaGroup::get_LCID(
	/*[out, retval]*/ long * pLCID  )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.get_LCID: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.get_LCID of group %s", W2A(group->m_Name));

	if (pLCID == NULL) {
      LOGFMTE( "get_LCID() failed with invalid argument(s):  pLCID is NULL" );
		res = E_INVALIDARG;
	}
	else {
		*pLCID = group->m_dwLCID;
		res = S_OK;
	}

	ReleaseGenericGroup();
	return res;
}




//=========================================================================
// IOPCGroupStateMgtDisp.put_LCID
// 
//=========================================================================
////[ propput ]
HRESULT STDMETHODCALLTYPE DaGroup::put_LCID(
	/*[in]*/ long LCID  )
{
	HRESULT        res;
	DaGenericGroup  *group;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.put_LCID: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.put_LCID of group %s", W2A(group->m_Name));

	group->m_dwLCID = LCID;
	ReleaseGenericGroup();
	return S_OK;
}




//=========================================================================
// IOPCGroupStateMgtDisp.CloneGroup
// 
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::CloneGroup(
	/*[in, optional]*/ VARIANT      Name,
	/*[out, retval] */ IDispatch ** ppDisp  )
{
	DaGenericGroup  *group;
	HRESULT res;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCGroupStateMgtDisp.CloneGroup: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCGroupStateMgtDisp.CloneGroup of group %s", W2A(group->m_Name));

	if( ppDisp != NULL ) {
		*ppDisp = NULL;
	}

	if( ( ppDisp == NULL ) || ( (V_VT(&Name) != VT_BSTR) && ( V_VT(&Name) != VT_EMPTY ) && (V_VT(&Name) != VT_NULL) ) ) {
         LOGFMTE( "CloneGroup() failed with invalid argument(s): ");
         if( ppDisp == NULL ) {
            LOGFMTE( "      ppDisp is NULL" );
         }
         if( (V_VT( &Name ) != VT_BSTR) && (V_VT( &Name ) != VT_EMPTY ) && (V_VT( &Name ) != VT_NULL) ) {
            LOGFMTE( "      Name must be VT_BSTR, VT_EMPTY or VT_NULL" );
         }

		ReleaseGenericGroup();
		return E_INVALIDARG;
	}


	// call custom interface implementation
	if( V_VT(&Name) != VT_BSTR ) {            // VT_EMPTY or VT_NULL
		res = CloneGroup( L"", IID_IDispatch, (IUnknown **)ppDisp); 
	} else {
		res = CloneGroup( V_BSTR( &Name ), IID_IDispatch, (IUnknown **)ppDisp); 
	}
	ReleaseGenericGroup();
	return res;
}



//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCSyncIODisp [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

//=========================================================================
// IOPCSyncIODisp.OPCRead
///
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::OPCRead(
	/*[in]           */ short     Source,
	/*[in]           */ long      NumItems,
	/*[in]           */ VARIANT   ServerHandles,
	/*[out]          */ VARIANT * pValues,
	/*[out, optional]*/ VARIANT * pQualities,
	/*[out, optional]*/ VARIANT * pTimeStamps,
	/*[out, optional]*/ VARIANT * errors    )
{
	OPCITEMSTATE      *pItemValues;
	//DaGenericItem     **ppGItems;
	DaGenericItem      *GItem;
	DaDeviceItem      **ppDItems;
	HRESULT           *pTempErrors;
	long               i;
	HRESULT            theRes, res;
	SAFEARRAY         *harr;
	long               h;
	DaGenericGroup     *group;
	VARIANT           TempV;
	HRESULT           Err;
	DWORD             AccessRight;
	BOOL              fQualityOutOfService;
	BOOL              *pfItemActiveState = NULL;    // Array with active state of all generic items
	DATE              timstamp;
	DWORD             dwNumOfDItems = 0;            // Number of successfully attached DeviceItems


	if ((Source != OPC_DS_CACHE) && (Source != OPC_DS_DEVICE) ) {
		return Err = E_INVALIDARG;
	}

	res = GetGenericGroup( &group );       // check group state and get ptr
	if( FAILED(res) ) {
      LOGFMTE("IOPCSyncIODisp.Read: Internal error: No generic Group");
		return res;
	}
   OPCWSTOAS( group->m_Name );
      LOGFMTI( "IOPCSyncIODisp.Read %ld items of group %s", NumItems, OPCastr );
   }

	if( pValues != NULL ) {
		VariantInit( pValues );
	}
	if( ( pValues == NULL ) || ( NumItems <= 0 ) || ( V_VT(&ServerHandles) != (VT_ARRAY | VT_I4) ) ) {
         LOGFMTE( "OPCRead() failed with invalid argument(s):" );
         if( pValues == NULL ) {
            LOGFMTE( "      pValues is NULL");
         }
         if( NumItems <= 0 ) {
            LOGFMTE( "      NumItems is %ld", NumItems );
         }
         if( V_VT( &ServerHandles ) != (VT_ARRAY | VT_I4) ) {
            LOGFMTE( "      ServerHandles must be an array of VT_I4" );
         }

		res = E_INVALIDARG;
		goto SynchIOReadBadExit0;
	}

	// create itemvalues array (of VARIANT)
	// the read values will be returned here
	res = VariantInitArrayOf( pValues, NumItems, VT_VARIANT );
	if( FAILED( res ) ) {
		goto SynchIOReadBadExit0;  
	}

	if( ( pQualities != NULL ) && (V_VT( pQualities ) != VT_NULL) ) {
		// create qualities array (of Integer)
		res = VariantInitArrayOf( pQualities, NumItems, VT_I2 );
		if( FAILED( res ) ) {
			goto SynchIOReadBadExit1;  
		}
	}

	if( ( pTimeStamps != NULL ) && (V_VT( pTimeStamps ) != VT_NULL) ) {
		// create timestamps array  (of Real)
		res = VariantInitArrayOf( pTimeStamps, NumItems, VT_DATE );
		if( FAILED( res ) ) {
			goto SynchIOReadBadExit2;  
		}
	}

	// create itemvalues array which content will be converted
	// to the separated arrays described above after the read operation
	pItemValues = new OPCITEMSTATE[ NumItems ];
	if( pItemValues == NULL ) {
		res = E_OUTOFMEMORY;
		goto SynchIOReadBadExit3;
	}
	i = 0;
	while ( i < NumItems ) {
		pItemValues[i].wQuality = 0;
		i ++;
	}

	if( ( errors != NULL ) && (V_VT( errors ) != VT_NULL) ) {
		// create errors array (VARIANT)
		res = VariantInitArrayOf( errors, NumItems, VT_I4 );
		if( FAILED( res ) ) {
			goto SynchIOReadBadExit4;  
		}
	}

	// Custom Errors array
	pTempErrors = new HRESULT [ NumItems ];
	if( pTempErrors == NULL ) {
		res = E_OUTOFMEMORY;
		goto SynchIOReadBadExit6;
	}
	i = 0;
	while ( i < NumItems ) {
		pTempErrors[i] = S_OK;
		i ++;
	}


	// convert server handles -> generic items
	res = E_INVALIDARG;
	if( V_VT(&ServerHandles) != (VT_ARRAY | VT_I4) ) {
		goto SynchIOReadBadExit7;
	} 

	harr = V_ARRAY( &ServerHandles );
	if( (harr == NULL) || (SafeArrayGetDim( harr ) != 1 )
		|| ( SafeArrayGetElemsize( harr ) != 4 ) ) {
			goto SynchIOReadBadExit7;
	}


	ppDItems = new DaDeviceItem*[ NumItems ];
	if( ppDItems == NULL ) {
		res = E_OUTOFMEMORY;
		goto SynchIOReadBadExit7;
	}
	i = 0;
	while ( i < NumItems ) {
		ppDItems[ i ] = NULL;
		i ++;
	}

	// if read from cache and group inactive then
	// all items must return quality OPC_QUALITY_OUT_OF_SERVICE
	if( ( Source == OPC_DS_CACHE ) && ( group->GetActiveState() == FALSE ) ) {
		fQualityOutOfService = TRUE;
	} else {
		fQualityOutOfService = FALSE;
	}

	// Generate array with active state of all generic items
	pfItemActiveState = new BOOL [NumItems];
	if( pfItemActiveState == NULL ) {
		res = E_OUTOFMEMORY;
		goto SynchIOReadBadExit8;
	}


	for( i=0 ; i < NumItems ; ++i ) {
		SafeArrayGetElement( harr, &i, &h);      // get the server handle

		Err = S_OK;
		// get the item with that handle
		res = group->GetGenericItem( h, &GItem );
		if( SUCCEEDED(res) ) {       
			if( GItem->AttachDeviceItem( &(ppDItems[i]) ) > 0 ) {
				// store the active state of the generic item
				pfItemActiveState[i] = GItem->get_Active();
				res = GItem->get_AccessRights( &AccessRight );
				if( SUCCEEDED(res) ) {
					if( (AccessRight & OPC_READABLE) != 0 ) {

						dwNumOfDItems++;

						// allowed to read this item
						pItemValues[i].hClient = GItem->get_ClientHandle( );
						// set the requested data type!
						V_VT(&pItemValues[i].vDataValue) = GItem->get_RequestedDataType();

					} else {
						// cannot read this item
						Err = OPC_E_BADRIGHTS;     
					}
					if( FAILED( Err ) ) {
						// this item cannot be read
						ppDItems[i]->Detach();
						ppDItems[i] = NULL;
					}
				} else {
					// access rights failed
					Err = res;     
				}
			} else {
				// device item not found
				Err = OPC_E_INVALIDHANDLE;
			}

			group->ReleaseGenericItem( h );

		} else {
			Err = OPC_E_INVALIDHANDLE;
		}

		SafeArrayPutElement( V_ARRAY( errors ), &i, &Err );
	}

	if (dwNumOfDItems) {
		// Read the current values
		theRes = group->InternalRead( 
			Source, 
			NumItems, 
			ppDItems, 
			pItemValues,
			pTempErrors
			);
	}
	else {
		theRes = S_FALSE;
	}

	VariantInit( &TempV );

	// now that the Read is done convert output
	// "Custom" to "Automation" ( struct to variant)
	for( i=0 ; i < NumItems ; ++i ) {
		// get the server handle
		SafeArrayGetElement( harr, &i, &h);


		// fill return array only if item not already excluded from read
		// and read not completely failed!
		if( ppDItems[i] != NULL ) {
			// Note : theRes is always S_OK or S_FALSE

			//VariantCopy( &TempV, &(pItemValues[i].vDataValue) );
			//SafeArrayPutElement( V_ARRAY( pValues, &i, (void *)&(TempV) );
			//VariantClear( &TempV );
			//VariantClear( &(pItemValues[i].vDataValue) );

			SafeArrayPutElement( V_ARRAY( pValues ), &i, (void *)&(pItemValues[i].vDataValue) );
			VariantClear( &(pItemValues[i].vDataValue) );

			// set quality
			// Set Quality to 'Bad: Out of Service' if
			// Group or Item is inactive.
			if ( fQualityOutOfService ||                                   // Group is inactive
				( (Source == OPC_DS_CACHE) &&                               // Item is inactive
				(pfItemActiveState[i] == FALSE) ) ) {

					pItemValues[i].wQuality = 
						OPC_QUALITY_BAD |                               // new quality is 'bad'
						OPC_QUALITY_OUT_OF_SERVICE |                    // new substatus is 'out of service'
						(pItemValues[i].wQuality & OPC_LIMIT_MASK);     // left the limit unchanged
			}

			SafeArrayPutElement( V_ARRAY( pQualities ), &i, 
				(void *)&(pItemValues[i].wQuality) );
			// convert timestamp to date
			FileTimeToDATE( &(pItemValues[i].ftTimeStamp), timstamp );
			SafeArrayPutElement( V_ARRAY( pTimeStamps ), &i, (void *)&(timstamp) );

			SafeArrayPutElement( V_ARRAY( errors ), &i, &(pTempErrors[i]) );

			ppDItems[i]->Detach();
		}

	}

	delete [] pfItemActiveState;
	delete [] ppDItems;
	delete [] pTempErrors;
	delete pItemValues; 
	ReleaseGenericGroup();
	return theRes;

SynchIOReadBadExit8:
	delete [] ppDItems;

SynchIOReadBadExit7:
	delete [] pTempErrors;

SynchIOReadBadExit6:
	//delete ppGItems;

	//SynchIOReadBadExit5:
	if( ( errors != NULL ) && (V_VT( errors ) != VT_NULL) ) {
		VariantUninitArrayOf( errors );
	}

SynchIOReadBadExit4:
	delete pItemValues; 

SynchIOReadBadExit3:
	if( ( pTimeStamps != NULL ) && (V_VT( pTimeStamps ) != VT_NULL) ) {
		VariantUninitArrayOf( pTimeStamps );
	}

SynchIOReadBadExit2:
	if( ( pQualities != NULL ) && (V_VT( pQualities ) != VT_NULL) ) {
		VariantUninitArrayOf( pQualities );
	}

SynchIOReadBadExit1:
	VariantUninitArrayOf( pValues );

SynchIOReadBadExit0:
	ReleaseGenericGroup();
	return res;

}





//=========================================================================
// IOPCSyncIODisp.OPCWrite
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::OPCWrite(
	/*[in]           */ long      NumItems,
	/*[in]           */ VARIANT   ServerHandles,
	/*[in]           */ VARIANT   Values,        
	/*[out, optional]*/ VARIANT * errors   )
{
	long               i;
	HRESULT            theRes, res;
	SAFEARRAY         *harr;
	long               h;
	DaGenericGroup     *group;
	HRESULT           Err;
	DWORD             AccessRight;
	DaGenericItem      *GItem;
	DWORD             dwNumOfDItems = 0;         // Number of successfully attached DeviceItems

	// members to store memory allocated with new operator
	HRESULT        *  pTempErrors;
	DaDeviceItem    ** ppDItems;
	CAutoVectorPtr<OPCITEMVQT> avpVQTs;

	res = GetGenericGroup( &group );                // check group state and get ptr
	if( FAILED(res) ) {
      LOGFMTE("IOPCSyncIODisp.OPCWrite: Internal error: No generic Group");
		return res;
	}
   OPCWSTOAS( group->m_Name );
      LOGFMTI( "IOPCSyncIODisp.OPCWrite %ld items of group %s", NumItems, OPCastr );
   }

	if(( NumItems <= 0 ) 
		|| ( V_VT(&ServerHandles) != (VT_ARRAY | VT_I4) ) 
		|| ( V_VT(&Values) != (VT_ARRAY | VT_VARIANT) ) ) {
         LOGFMTE( "OPCWrite() failed with invalid argument(s):" );
         if( NumItems <= 0 ) {
            LOGFMTE( "      NumItems is %ld", NumItems );
         }
         if( V_VT(&ServerHandles) != (VT_ARRAY | VT_I4) ) {
            LOGFMTE( "      ServerHandles must be an array of VT_I4" );
         }
         if( V_VT(&Values) != (VT_ARRAY | VT_VARIANT) ) {
            LOGFMTE( "      Values must be an array of VT_VARIANT" );
         }
      res = E_INVALIDARG;                 // Version 1.3, AM
			goto SynchIOWriteBadExit0;
	}

	if( ( errors != NULL ) && (V_VT( errors ) != VT_NULL) ) {
		// create errors array (VARIANT)
		res = VariantInitArrayOf( errors, NumItems, VT_I4 );
		if( FAILED( res ) ) {
			goto SynchIOWriteBadExit0;  
		}
	}

	// Custom Errors array
	pTempErrors = new HRESULT [ NumItems ];
	if( pTempErrors == NULL ) {
		res = E_OUTOFMEMORY;
		goto SynchIOWriteBadExit2;
	}
	// set error results to S_OK
	i = 0;
	while ( i < NumItems ) {
		pTempErrors[ i ] = S_OK;
		i ++;
	}



	// convert server handles -> generic items
	res = E_INVALIDARG;
	if( V_VT(&ServerHandles) != (VT_ARRAY | VT_I4) ) {
		goto SynchIOWriteBadExit3;
	} 

	harr = V_ARRAY( &ServerHandles );
	if( (harr == NULL) || (SafeArrayGetDim( harr ) != 1 )
		|| ( SafeArrayGetElemsize( harr ) != 4 ) ) {
			goto SynchIOWriteBadExit3;
	}


	ppDItems = new DaDeviceItem*[ NumItems ];
	if( ppDItems == NULL ) {
		res = E_OUTOFMEMORY;
		goto SynchIOWriteBadExit3;
	}
	i = 0;
	while ( i < NumItems ) {
		ppDItems[i] = NULL;
		i ++;
	}

	if (!avpVQTs.Allocate( NumItems )) {
		res = E_OUTOFMEMORY;
		goto SynchIOWriteBadExit4;
	}
	memset( avpVQTs, 0, sizeof (OPCITEMVQT) * NumItems );

	for( i=0 ; i < NumItems ; ++i ) {
		SafeArrayGetElement( harr, &i, &h);      // get the server handle

		Err = S_OK;
		// get the item with that handle
		res = group->GetGenericItem( h, &GItem );
		if( SUCCEEDED(res) ) {       
			if( GItem->AttachDeviceItem( &(ppDItems[i]) ) > 0 ) {

				res = GItem->get_AccessRights( &AccessRight );
				if( SUCCEEDED(res) ) {
					if( (AccessRight & OPC_WRITEABLE) != 0 ) {

						VariantInit( &avpVQTs[i].vDataValue );
						res = SafeArrayGetElement( V_ARRAY( &Values ), &i, &avpVQTs[i].vDataValue );
						if( SUCCEEDED(res) ) {
							dwNumOfDItems++;
						} else {
							Err = E_OUTOFMEMORY;
						}
					} else {
						// cannot read this item
						Err = OPC_E_BADRIGHTS;     
					}
					if( FAILED( Err ) ) {
						// this item cannot be read
						ppDItems[i]->Detach();
						ppDItems[i] = NULL;
					}
				} else {
					// access rights failed
					Err = res;     
				}
			} else {
				// device item not found
				Err = OPC_E_INVALIDHANDLE;
			}

			group->ReleaseGenericItem( h );

		} else {
			Err = OPC_E_INVALIDHANDLE;
		}

		SafeArrayPutElement( V_ARRAY( errors ), &i, &Err );
	}

	// Read the current values
	if (dwNumOfDItems) {
		theRes = group->InternalWriteVQT(
			NumItems,
			ppDItems,
			avpVQTs,
			pTempErrors
			);

		// now that the Read is done convert output
		// "Custom" to "Automation" ( struct to variant)
		for( i=0 ; i < NumItems ; ++i ) {
			if( ppDItems[i] != NULL ) {
				ppDItems[i]->Detach();
				SafeArrayPutElement( V_ARRAY( errors ), &i, &(pTempErrors[i]) );
			}
		}
	}
	else {
		theRes = S_FALSE;                      // There are no Device Items to write, detailed error codes
	}                                         // are stored in the error array.

	for (i=0; i<NumItems; i++) {
		VariantClear( &avpVQTs[i].vDataValue );
	}
	avpVQTs.Free();

	delete [] ppDItems;
	delete [] pTempErrors;
	ReleaseGenericGroup();
	return theRes;

SynchIOWriteBadExit4:
	delete [] ppDItems;

SynchIOWriteBadExit3:
	delete [] pTempErrors;

SynchIOWriteBadExit2:
	if( ( errors != NULL ) && (V_VT( errors ) != VT_NULL) ) {
		VariantUninitArrayOf( errors );
	}

SynchIOWriteBadExit0:
	ReleaseGenericGroup();
	return res;
}





//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[ IOPCAsyncIODisp [[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[

//=========================================================================
// IOPCAsyncIODisp.OPCRead
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::OPCRead(
	/*[in]           */ short     Source,
	/*[in]           */ long      NumItems,
	/*[in]           */ VARIANT   ServerHandles,
	/*[out, optional]*/ VARIANT * errors,
	/*[out, retval]  */ long    * pTransactionID  )
{
	DaGenericGroup     *group;
	DaGenericItem      *GItem ;
	HRESULT            res;
	HRESULT            Err;
	long               i, h;
	SAFEARRAY         *harr;
	long               th;
	WORD               QualityOutOfService;
	BOOL               bInvalidHandles;
	OPCITEMSTATE      *pItemStates;
	DaDeviceItem      **ppDItems;
	DaAsynchronousThread     *pAdv ;

	if ((Source != OPC_DS_CACHE) && (Source != OPC_DS_DEVICE) ) {
		return E_INVALIDARG;
	}

	res = GetGenericGroup( &group );       // check group state and get ptr
	if( FAILED(res) ) {
      LOGFMTE("IOPCAsyncIODisp.OPCRead: Internal error: No generic Group");
		return res;
	}
   OPCWSTOAS( group->m_Name );
      LOGFMTI( "IOPCAsyncIODisp.OPCRead %ld items of group %s", NumItems, OPCastr );
   }

	if( pTransactionID != NULL ) {
		*pTransactionID = 0;
	}

	if( ( pTransactionID == NULL ) || ( NumItems <= 0 ) || ( V_VT(&ServerHandles) != (VT_ARRAY | VT_I4) ) ) {
         LOGFMTE( "OPCRead() failed with invalid argument(s):" );
         if( pTransactionID == NULL ) {
            LOGFMTE( "      pTransactionID is NULL");
         }
         if( NumItems <= 0 ) {
            LOGFMTE( "      NumItems is %ld", NumItems );
         }
         if( V_VT(&ServerHandles) != (VT_ARRAY | VT_I4) ) {
            LOGFMTE( "      ServerHandles must be an array of VT_I4" );
         }

		res = E_INVALIDARG;
		goto ASReadExit0;
	}

	harr = V_ARRAY( &ServerHandles );

	// flag to indicate if there are invalid handles
	// used also as condition for freeing memory
	bInvalidHandles = FALSE;

	if( ( errors != NULL ) && (V_VT( errors ) != VT_NULL) ) {
		// create errors array (VARIANT)
		res = VariantInitArrayOf( errors, NumItems, VT_I4 );
		if( FAILED(res) ) {
			goto ASReadExit0 ;
		}
	}

	// create client handle array
	pItemStates = new OPCITEMSTATE [ NumItems ];
	if( pItemStates == NULL ) {
		res = E_OUTOFMEMORY;
		goto ASReadExit1;
	}
	i = 0;
	while ( i < NumItems ) {
		pItemStates[i].wQuality = 0; 
		i ++;
	}

	// create deviceItem array
	ppDItems = new DaDeviceItem*[ NumItems ];
	if( ppDItems == NULL ) {
		res = E_OUTOFMEMORY;
		goto ASReadExit2;
	}
	for (i=0; i<NumItems; i++) {
		ppDItems[i] = NULL;
	}


	// if read from cache and group inactive then
	// all items must return quality OPC_QUALITY_OUT_OF_SERVICE
	if( ( Source == OPC_DS_CACHE ) && ( group->GetActiveState() == FALSE ) ) {
		QualityOutOfService = OPC_QUALITY_OUT_OF_SERVICE;
	} else {
		QualityOutOfService = 0;
	}


	bInvalidHandles = FALSE;

	i = 0; 
	while ( i < NumItems ) {
		SafeArrayGetElement( harr, &i, &h);      // get the server handle

		Err = S_OK;
		// get the item with that handle
		res = group->GetGenericItem( h, &GItem );
		if( SUCCEEDED(res) ) {        
			if( GItem->AttachDeviceItem( &(ppDItems[i]) ) > 0 ) { 
				// ... inverse operation is done when thread ended with DeviceItem.Detach() !
				// set client handle into result array
				pItemStates[i].hClient = GItem->get_ClientHandle() ;  
				// set requested data type
				V_VT(&pItemStates[i].vDataValue) = GItem->get_RequestedDataType();
				// set quality
				pItemStates[i].wQuality |= QualityOutOfService;

			} else { 
				// could not attach
				Err = OPC_E_INVALIDHANDLE;
			}

			group->ReleaseGenericItem( h );

		} else {
			Err = OPC_E_INVALIDHANDLE;
		}

		if( ( errors != NULL ) && (V_VT( errors ) != VT_NULL) ) {
			// care for errors
			SafeArrayPutElement( V_ARRAY( errors ), &i, &Err );
		}

		i++;
	}

	pAdv =  new DaAsynchronousThread( group );          // create advice object
	if( pAdv == NULL ) {
		res = E_OUTOFMEMORY;
		goto ASReadExit4;
	}

	// add async read advise thread to the thread list of the group
	EnterCriticalSection( &( group->m_AsyncThreadsCritSec ) );
	th = group->m_oaAsyncThread.New();
	res = group->m_oaAsyncThread.PutElem( th, pAdv );
	LeaveCriticalSection( &( group->m_AsyncThreadsCritSec ) );
	if( FAILED(res) ) {
		// could not add to Thread list
		goto ASReadExit5;
	}

	res = pAdv->CreateAutomationRead(   (OPCDATASOURCE)Source,  // from cache or device
		NumItems,               // number of items to handle
		ppDItems,               // array of DeviceItems to be read
		pItemStates,            // client handles 
		NULL,
		pTransactionID );       // out parameter

	// Note :   Do no longer use pAdr after Create function because object deletes itself.
	//          Also all provided arrays will be deleted.
	ReleaseGenericGroup();
	return res;
	//-------------

ASReadExit5:
	delete pAdv;

ASReadExit4:
	for (i=0; i<NumItems; i++) {                 // Release the DeviceItems
		if (ppDItems[i]) {
			ppDItems[i]->Detach() ;
		}
	}
	delete [] ppDItems;

ASReadExit2:
	delete [] pItemStates;

ASReadExit1:
	if ((errors != NULL) && (V_VT( errors ) != VT_NULL)) {
		VariantUninitArrayOf( errors );
	}

ASReadExit0:
	ReleaseGenericGroup();
	return res;
}





//=========================================================================
// IOPCAsyncIODisp.OPCWrite
//
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::OPCWrite(
	/*[in]           */ long      NumItems,
	/*[in]           */ VARIANT   ServerHandles,
	/*[in]           */ VARIANT   Values,        
	/*[out, optional]*/ VARIANT * errors,
	/*[out, retval]  */ long    * pTransactionID )
{
	DaGenericItem      *GItem ;
	HRESULT            res, Err;
	SAFEARRAY         *harr;
	long               i, h;
	DaGenericGroup     *group;
	VARIANT            TempV;
	long               th;
	BOOL               bInvalidHandles;
	HRESULT            theRes;
	// members to store memory allocated with new operator
	long           *  pClientHandles ;
	DaDeviceItem    ** ppDItems;
	VARIANT        *  paVariantValues;
	DaAsynchronousThread  *  pAdv;

	res = GetGenericGroup( &group );       // check group state and get ptr
	if( FAILED(res) ) {
      LOGFMTE("IOPCAsyncIODisp.OPCWrite: Internal error: No generic Group");
		return res;
	}
   OPCWSTOAS( group->m_Name );
      LOGFMTI( "IOPCAsyncIODisp.OPCWrite %ld items of group %s", NumItems, OPCastr );
   }

	if( pTransactionID != NULL ) {
		*pTransactionID = 0;
	}

	if(( pTransactionID == NULL ) 
		|| ( NumItems <= 0 )
		|| ( V_VT(&ServerHandles) != (VT_ARRAY | VT_I4) ) 
		|| ( V_VT(&Values) != (VT_ARRAY | VT_VARIANT) ) ) {
         LOGFMTE( "OPCWrite() failed with invalid argument(s):" );
         if( pTransactionID == NULL ) {
            LOGFMTE( "      pTransactionID is NULL");
         }
         if( NumItems <= 0 ) {
            LOGFMTE( "      NumItems is %ld", NumItems );
         }
         if( V_VT(&ServerHandles) != (VT_ARRAY | VT_I4) ) {
            LOGFMTE( "      ServerHandles must be an array of VT_I4" );
         }
         if( V_VT(&Values) != (VT_ARRAY | VT_VARIANT) ) {
            LOGFMTE( "      Values must be an array of VT_VARIANT" );
         }

			res = E_INVALIDARG;                 // Version 1.3, AM
			goto ASWriteExit0;  
	}

	harr = V_ARRAY( &ServerHandles );

	// flag to indicate if there are invalid handles
	// used also as condition for freeing memory
	bInvalidHandles = FALSE;

	if( ( errors != NULL ) && (V_VT( errors ) != VT_NULL) ) {
		// create errors array (VARIANT)
		res = VariantInitArrayOf( errors, NumItems, VT_I4 );
		if( FAILED(res) ) {
			goto ASWriteExit0;  
		} 
	}
	// create client handle array
	pClientHandles = new long[ NumItems ];
	if( pClientHandles == NULL ) {
		res = E_OUTOFMEMORY;
		goto ASWriteExit1;
	}
	memset( pClientHandles, 0, NumItems * sizeof(long) ); // fill with 0's (invalid handle)


	ppDItems = new DaDeviceItem*[ NumItems ];
	if( ppDItems == NULL ) {
		res = E_OUTOFMEMORY;
		goto ASWriteExit3;
	}
	i = 0;
	while ( i < NumItems ) {
		ppDItems[i] = NULL;
		i ++;
	}

	paVariantValues = new VARIANT[ NumItems ];
	if( paVariantValues == NULL ) {
		res = E_OUTOFMEMORY;
		goto ASWriteExit4;
	}
	i = 0;
	while ( i < NumItems ) {
		VariantInit( &(paVariantValues[i]) );
		i ++;
	}

	theRes = S_OK;
	i = 0;
	while ( i < NumItems ) {

		SafeArrayGetElement( harr, &i, &h);      // get the server handle

		Err = S_OK;

		// get the item with that handle
		res = group->GetGenericItem( h, &GItem );
		if( SUCCEEDED(res) ) {        
			if( GItem->AttachDeviceItem( &(ppDItems[i]) ) > 0 ) {
				// ... inverse operation is done when thread ended with DeviceItem.Detach() !

				res = SafeArrayGetElement( V_ARRAY( &Values ), &i, &TempV );
				if( SUCCEEDED(res) ) {
					res = VariantCopy( &(paVariantValues[i]), &TempV );
					if( SUCCEEDED(res) ) {
						// set client handle into result array
						pClientHandles[ i ] = GItem->get_ClientHandle() ;
					} else {
						Err = E_OUTOFMEMORY;
					}
					VariantClear( &TempV );
				} else {
					Err = E_OUTOFMEMORY;
				}

				if( FAILED( Err ) ) {
					// this item cannot be read
					ppDItems[i]->Detach();
					ppDItems[i] = NULL;
					// bad error: all fails 
					theRes = Err;
				}
			} else {
				bInvalidHandles = TRUE;
				Err = OPC_E_INVALIDHANDLE;
			}
			group->ReleaseGenericItem( h );
		} else {
			Err = OPC_E_INVALIDHANDLE;
			bInvalidHandles = TRUE;
		}

		if( ( errors != NULL ) && (V_VT( errors ) != VT_NULL) ) {
			// care for errors
			SafeArrayPutElement( V_ARRAY( errors ), &i, &Err );
		}

		i ++;
	}

	if( FAILED(theRes) ) {
		res = theRes;
		goto ASWriteExit6;
	}

	if( bInvalidHandles == TRUE ) {
		res = S_FALSE;
		goto ASWriteExit6;
	}


	pAdv =  new DaAsynchronousThread( group );          // create advice object
	if( pAdv == NULL ) {
		goto ASWriteExit6;
	}

	// add async read advise thread to the thread list of the group
	EnterCriticalSection( &( group->m_AsyncThreadsCritSec ) );
	th = group->m_oaAsyncThread.New();
	res = group->m_oaAsyncThread.PutElem( th, pAdv );
	LeaveCriticalSection( &( group->m_AsyncThreadsCritSec ) );
	if( FAILED(res) ) {
		// could not add to Thread list
		goto ASWriteExit7;
	}

	res = pAdv->CreateAutomationWrite(  NumItems,               // number of items to handle
		ppDItems,               // array of DeviceItems to be read
		pClientHandles,         // client handle result array
		paVariantValues,        // itesm values to set
		NULL,
		pTransactionID );       // out parameter

	// Note :   Do no longer use pAdr after Create function because object deletes itself.
	//          Also all provided arrays will be deleted.
	ReleaseGenericGroup();
	return res;

ASWriteExit7:
	delete pAdv;

ASWriteExit6:
	for (i=0; i<NumItems; i++) {
		if (ppDItems[i]) {
			ppDItems[i]->Detach() ;
		}
		VariantClear( &paVariantValues[i] );
	}
	delete [] paVariantValues;

ASWriteExit4:
	delete [] ppDItems;

ASWriteExit3:
	delete [] pClientHandles;

ASWriteExit1:
	if (bInvalidHandles == FALSE) {
		if ((errors != NULL) && (V_VT( errors ) != VT_NULL)) {
			VariantUninitArrayOf( errors );
		}
	}

ASWriteExit0:
	ReleaseGenericGroup();
	return res;

}




//=========================================================================
// IOPCAsyncIODisp.Cancel
// 
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::Cancel(
	/*[in]*/ long TransactionID  )
{
	HRESULT res;

   LOGFMTI( "IOPCAsyncIODisp.Cancel Transaction %ld", TransactionID );

	res = Cancel( (DWORD) TransactionID );
	return res;
}




//=========================================================================
// IOPCAsyncIODisp.Refresh
// 
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::Refresh(
	/*[in]         */ short   Source,
	/*[out, retval]*/ long  * pTransactionID  )
{
	DaGenericGroup     *group;
	DaGenericItem      *GItem ;
	DaDeviceItem      **ppDItems;
	HRESULT            res;
	OPCITEMSTATE      *pItemStates ;
	long               i;
	DaAsynchronousThread     *pAdv ;
	long               th;
	long               ih, newih;
	long               nGroupItems, nItemsToSend;

	if ((Source != OPC_DS_CACHE) && (Source != OPC_DS_DEVICE) ) {
		return E_INVALIDARG;
	}

	res = GetGenericGroup( &group );       // check group state and get ptr
	if( FAILED(res) ) {
      LOGFMTE("IOPCAsyncIODisp.Refresh: Internal error: No generic Group");
		return res;
	}
   OPCWSTOAS( group->m_Name );
      LOGFMTI( "IOPCAsyncIODisp.Refresh group %s from %d", OPCastr, Source );
   }

	if( pTransactionID != NULL ) {
		*pTransactionID = 0;
	}

	if( pTransactionID == NULL ) {
		LOGFMTE( "Refresh() failed with invalid argument(s):  pTransactionID is NULL" );
		res = E_INVALIDARG;
		goto AsRefreshExit0;
	}

	// refresh is only allowed on active groups
	if( ( Source == OPC_DS_CACHE ) && ( group->GetActiveState() == FALSE ) ) {
		res = E_FAIL;
		goto AsRefreshExit0;
	}

	EnterCriticalSection( &(group->m_ItemsCritSec) );
	nGroupItems = (DWORD) group->m_oaItems.TotElem();

	// create client handle array
	pItemStates = new OPCITEMSTATE [ nGroupItems ];
	if( pItemStates == NULL ) {
		LeaveCriticalSection( &( group->m_ItemsCritSec ) );
		res = E_OUTOFMEMORY;
		goto AsRefreshExit0;
	}
	memset( pItemStates, 0, nGroupItems * sizeof(OPCITEMSTATE) ); // fill with 0's

	// create deviceItem array
	ppDItems = new DaDeviceItem*[ nGroupItems ];
	if( ppDItems == NULL ) {
		LeaveCriticalSection( &( group->m_ItemsCritSec ) );
		res = E_OUTOFMEMORY;
		goto AsRefreshExit1;
	}
	for (i=0; i<nGroupItems; i++) {
		ppDItems[i] = NULL;
	}


	// get through the items of a group and
	// and build the arrays taking into account only those
	// which are active
	// !!!??? in theory only those R or RW should be taken into account!?
	nItemsToSend = 0;
	res = group->m_oaItems.First( &ih );
	while ( SUCCEEDED( res ) ) {
		group->m_oaItems.GetElem( ih, &GItem );
		if( GItem->get_Active() == TRUE ) {
			// include the active item
			if( GItem->AttachDeviceItem( &ppDItems[ nItemsToSend ] ) > 0 )  { 
				// DeviceItem attached
				pItemStates[ nItemsToSend ].hClient = GItem->get_ClientHandle() ;
				// set requested data type
				V_VT(&pItemStates[ nItemsToSend ].vDataValue) = GItem->get_RequestedDataType() ;

				nItemsToSend ++;
			}
		}
		res = group->m_oaItems.Next( ih, &newih );
		ih = newih;
	}

	LeaveCriticalSection( &( group->m_ItemsCritSec ) );

	if( nItemsToSend == 0 ) {
		res = E_FAIL;
		goto AsRefreshExit3;
	}        

	// create advise object
	pAdv =  new DaAsynchronousThread( group );          
	if( pAdv == NULL ) {
		res = E_OUTOFMEMORY;
		goto AsRefreshExit3;
	}

	// add async refresh advise thread to the thread list of the group
	EnterCriticalSection( &( group->m_AsyncThreadsCritSec ) );
	th = group->m_oaAsyncThread.New();
	res = group->m_oaAsyncThread.PutElem( th, pAdv );
	LeaveCriticalSection( &( group->m_AsyncThreadsCritSec ) );
	if( FAILED(res) ) {
		// could not add to Thread list
		goto AsRefreshExit4;
	}

	res = pAdv->CreateAutomationRead(   (OPCDATASOURCE)Source,  // from cache or device
		nItemsToSend,           // number of items to handle
		ppDItems,               // array of DeviceItems to be read
		pItemStates,            // client handles 
		NULL,
		pTransactionID );       // out parameter

	// Note :   Do no longer use pAdr after Create function because object deletes itself.
	//          Also all provided arrays will be deleted.
	ReleaseGenericGroup();
	return res;

	//-------------
AsRefreshExit4:
	delete pAdv;

AsRefreshExit3:
	for (i=0; i<nGroupItems; i++) {
		if (ppDItems[i]) {
			ppDItems[i]->Detach() ;
		}
	}
	delete [] ppDItems;

AsRefreshExit1:
	delete [] pItemStates;

AsRefreshExit0:
	ReleaseGenericGroup();
	return res;
}




//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[ IOPCPublicGroupStateMgtDisp [[[[[[[[[[[[[[[[[[[[[[[[
//[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[


//=========================================================================
// IOPCPublicGroupStateMgtDisp.get_State
//    [ propget ]
// 
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::get_State(
	/*[out, retval]*/ VARIANT_BOOL * pPublic  )
{
	HRESULT        res;
	DaGenericGroup  *group;
	long           pgh;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCPublicGrpupStateMgtDisp.get_State: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCPublicGroupStateMgtDisp.get_State of group %s", W2A(group->m_Name));

	if( pPublic == NULL ) {
		LOGFMTE( "get_State() failed with invalid argument(s):  pPublic is NULL" );
		ReleaseGenericGroup();
		return E_INVALIDARG;
	}

	if( group->GetPublicInfo( &pgh ) == FALSE ) {
		*pPublic = VARIANT_FALSE;
	} else {
		*pPublic = VARIANT_FALSE;
	}

	ReleaseGenericGroup();
	return S_OK;
}


//=========================================================================
// IOPCPublicGroupStateMgtDisp.MoveToPublic
//=========================================================================
HRESULT STDMETHODCALLTYPE DaGroup::MoveToPublic_DISP(
	void  )
{
	DaGenericGroup  *group;
	HRESULT res;

	res = GetGenericGroup( &group );
	if( FAILED(res) ) {
      LOGFMTE("IOPCPublicGrpupStateMgtDisp.MoveToPublic: Internal error: No generic group.");
		return res;
	}
	USES_CONVERSION;
	LOGFMTI( "IOPCPublicGroupStateMgtDisp.MoveToPublic of group %s", W2A(group->m_Name));

	res = MoveToPublic();
	ReleaseGenericGroup();
	return res;
}

//DOM-IGNORE-END

