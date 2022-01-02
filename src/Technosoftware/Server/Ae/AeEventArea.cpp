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

#ifdef   _OPC_SRV_AE                            // Alarms & Events Server

//DOM-IGNORE-BEGIN
//-------------------------------------------------------------------------
// INLCUDE
//-------------------------------------------------------------------------
#include "stdafx.h"
#include "UtilityDefs.h"
#include "MatchPattern.h"
#include "AeEvent.h"
#include "AeEventArea.h"

//-------------------------------------------------------------------------
// STATIC MEMBERS
//-------------------------------------------------------------------------
// Default delimiter character between the branches and the leafs.
WCHAR EventArea::delimiter_[2] = L".";


//-------------------------------------------------------------------------
// CODE
//-------------------------------------------------------------------------

EventArea::EventArea()
{
   areaId_			= AREAID_UNSPECIFIED;
   parent_			= NULL;
   root_			= NULL;
   partialName_		= NULL;
}


HRESULT EventArea::Create( EventArea* parent, DWORD areaId, LPCWSTR areaName )
{
   _ASSERTE( areaName != NULL );				// Must not be NULL

                                                // Set the root object
   root_     = (parent != NULL) ? parent->root_ : this;
   parent_   = parent;
                                                // Check if the ID is alread used
   EventArea* pTmp;
   HRESULT hres = GetArea( areaId, &pTmp );
   if (SUCCEEDED( hres )) return E_OBJECT_IN_LIST;

   areaId_  = areaId;							// The ID must not be set before
                                                // call of function GetArea().
                                                // Set the fully qualified area name
   if (parent_ && !parent_->IsRoot()) {			// It's a sub-area
      LPWSTR szQualifiedName = new WCHAR [wcslen( parent_->Name() ) + wcslen( areaName ) + 2 ];
                                                // +2 for EOS and the delimiter character
      if (!szQualifiedName) {
         hres = E_OUTOFMEMORY;
      }
      else {
         wcscpy( szQualifiedName, parent_->Name() );
         wcscat( szQualifiedName, delimiter_ );
         wcscat( szQualifiedName, areaName );
         hres = name_.SetString( szQualifiedName );
         delete [] szQualifiedName;             // Temporary area name no longer used
                                                // Initialize pointer to the partial area name
         partialName_ = name_ + wcslen( parent_->Name() ) + 1;
      }
   }
   else {                                       // It is not a sub-area
      hres = name_.SetString( areaName );		// Set the fully qualified area name
      partialName_ = name_;						// Initialize pointer to the partial area name
   }

   if (SUCCEEDED( hres )) {
      if (parent_) {							// Add it as sub-area to the parent
         parent_->subAreasLock_.Lock();
         hres = parent_->subAreas_.Add( static_cast<EventArea*>(this) ) ? S_OK : E_OUTOFMEMORY;
         parent_->subAreasLock_.Unlock();
      }
   }

   return hres;
}


EventArea::~EventArea()
{
   subAreasLock_.Lock();
   //
   // TODO for dynamic Area Space
   //
   // The data member 'current area' of the browser object
   // must be protected if areas can be removed dynamically
   // from the Area Space.
   //
   // Remove this area from the parent area.
   // Note: This moves the entries within the array.
   //
   // if (parent_) {
   //    parent_->subAreasLock_.Lock();
   //    hres = parent_->subAreas_.Remove( (EventArea*)this );
   //    parent_->subAreasLock_.Unlock();
   // }

   for (int i=0; i < subAreas_.GetSize(); i++) {
      delete subAreas_[i];
   }
   subAreas_.RemoveAll();
   subAreasLock_.Unlock();

   sourcesLock_.Lock();
   sources_.RemoveAll();
   sourcesLock_.Unlock();
}



//-------------------------------------------------------------------------
// OPERATIONS
//-------------------------------------------------------------------------

HRESULT EventArea::PositionDown( LPCWSTR areaName, EventArea** area)
{
   *area = NULL;

   subAreasLock_.Lock();
   for (int i=0; i < subAreas_.GetSize(); i++) {

      if (wcscmp( subAreas_[i]->partialName_, areaName) == 0) {
         *area = subAreas_[i];
         subAreasLock_.Unlock();
         return S_OK;
      }
   }
   subAreasLock_.Unlock();
   return OPC_E_INVALIDBRANCHNAME;
}


HRESULT EventArea::PositionTo( LPCWSTR areaName, EventArea** area )
{
   HRESULT		hres;
   LPCWSTR		name;
   WideString	nameCopy;
   EventArea*	eventArea = root_;
   WCHAR*		ptr;

   *area = NULL;

   hres = nameCopy.SetString(areaName);			// make a copy because wcstok()
   if (FAILED( hres ))                          // modifies the string
      return hres;

                                                // get top area name
   name = wcstok_s( nameCopy, delimiter_, &ptr);
   if (!name) {									// move to the root if the string is empty
      *area = root_;
      return S_OK;
   }

   while (name) {								// move to next sub-area
      hres = eventArea->PositionDown( name, &eventArea );
      if (FAILED( hres )) {                     // invalid area name
         return OPC_E_INVALIDBRANCHNAME;
      }                                         // get next sub-area name
      name = wcstok_s(NULL, reinterpret_cast<LPCWSTR>(&delimiter_), &ptr);
   }
   *area = eventArea;
   return S_OK;
}


HRESULT EventArea::GetQualifiedAreaName( LPCWSTR areaName, LPWSTR* qualifiedAreaName)
{
   HRESULT     hres;
   EventArea* eventArea;
                                                
   hres = PositionDown(areaName, &eventArea );   // get the sub-area
   if (SUCCEEDED( hres )) {                     // get the fully qualified name
      *qualifiedAreaName = eventArea->Name().CopyCOM();
      if (*qualifiedAreaName == NULL) {
         hres = E_OUTOFMEMORY;
      }
   }
   else if (hres == OPC_E_INVALIDBRANCHNAME) {
      hres = E_INVALIDARG;
   }
   return hres;
}


HRESULT EventArea::GetQualifiedSourceName( LPCWSTR sourceName, LPWSTR* qualifiedSourceName)
{
   HRESULT        hres = E_INVALIDARG;
   AeSource*  pSource;

   sourcesLock_.Lock();                          // search the name in the source list of this area
   for (int i=0; i < sources_.GetSize(); i++) {
      pSource = sources_[i];
      if (wcscmp( pSource->PartialName(), sourceName) == 0) {
         *qualifiedSourceName = pSource->Name().CopyCOM();
         hres = (*qualifiedSourceName == NULL) ? E_OUTOFMEMORY : S_OK;
         break;
      }
   }
   sourcesLock_.Unlock();

   return hres;
}


HRESULT EventArea::BrowseAreas( EventAreaBrowser*		areaBrowser,
                                 COpcComEnumString*     comEnum,
                                 OPCAEBROWSETYPE        browseFilterType,
                                 LPCWSTR                filterCriteria)
{
   HRESULT     hres = S_OK;
   LPWSTR*     names = NULL;					// areas/sources that match the filter criteria
   LPWSTR      tempName;
   DWORD       match = 0;

   (browseFilterType == OPC_AREA) ? subAreasLock_.Lock() : sourcesLock_.Lock();
   try {
                                                // max. number of names
      DWORD dwSize   = (browseFilterType == OPC_AREA) ?
                                    subAreas_.GetSize() : sources_.GetSize();

      names = new LPWSTR [ dwSize ];
      if (names == NULL)
         throw E_OUTOFMEMORY;

      for (DWORD i=0; i < dwSize; i++) {
         tempName = (browseFilterType == OPC_AREA) ?
                                    const_cast<LPWSTR>(subAreas_[i]->partialName_) :
                                    const_cast<LPWSTR>(sources_[i]->PartialName());
												// filter the name
         if (*filterCriteria) {					// there is a server specific filter defined
            if (FilterName( tempName, filterCriteria)) {
               names[match] = tempName;			// name passes the filter
               match++;
            }
         }
         else {									// no filter
            names[match] = tempName;      
            match++;
         }
      }
   }
   catch (HRESULT hresEx) {
      hres = hresEx;
   }
   (browseFilterType == OPC_AREA) ? subAreasLock_.Unlock() : sourcesLock_.Unlock();

   if (SUCCEEDED( hres )) {
      hres = comEnum->Init(
                        names, &names[ match ],
                        reinterpret_cast<LPUNKNOWN>(areaBrowser), // this will do an AddRef to the owner
                        AtlFlagCopy );          // makes a copy
      if (SUCCEEDED( hres ) && match == 0) {
         hres = S_FALSE;                        // result is S_FALSE if there is nothing to enumerate
      }
   }
   if (names) {									// free the temporary array of names
      delete [] names;
   }
   return hres;
}


HRESULT EventArea::GetArea( DWORD areaId, EventArea** area)
{
   *area = NULL;
   return root_->GetSubArea(areaId, area);
}


HRESULT EventArea::AttachSource( AeSource* source)
{
   HRESULT hres;

   sourcesLock_.Lock();
   hres = sources_.Add(source) ? S_OK : E_OUTOFMEMORY;
   sourcesLock_.Unlock();
   return hres;
}


HRESULT EventArea::DetachSource( AeSource* source)
{
   _ASSERTE(source != NULL );					// Must not be NULL

   HRESULT hres;
   sourcesLock_.Lock();
   hres = sources_.Remove(source) ? S_OK : E_FAIL;
   sourcesLock_.Unlock();
   return hres;
}


HRESULT EventArea::EnableConditions( BOOL enable, CSimpleValArray<AeEventArray*>& events)
{
   AeEventArray* parEvents;
   HRESULT        hres;
   HRESULT        hresRet = S_OK;

   sourcesLock_.Lock();                          // Search the name in the source list of this area

   for (int i=0; i < sources_.GetSize(); i++) {

      parEvents = new AeEventArray;
      if (!parEvents) {
         hresRet = E_OUTOFMEMORY;
         break;
      }

      hres = sources_[i]->EnableConditions(enable, parEvents );
      if (FAILED( hres )) {
         hresRet = hres;
         break;
      }
      if (hres == S_FALSE) {
         hresRet = hres;						// Not succeeded for all conditions.
      }

      if (parEvents->NumOfEvents()) {			// Add it only to the array if at least one state has changed.
         hres = events.Add( parEvents );
         if (FAILED( hres )) {
            hresRet = S_FALSE;					// Not succeeded for all conditions.
         }
      }
      else {
         delete parEvents;
      }
   }
   sourcesLock_.Unlock();

   return hresRet;
}


BOOL EventArea::ExistArea( LPCWSTR areaName)
{
   if (MatchPattern( name_, areaName)) {
      return TRUE;                              // This Area matches to the specified name.
   }

   BOOL fExist = FALSE;                         // If this Area doesn't match to the specified
   subAreasLock_.Lock();                        // name then test if a Sub-Area matches.
   int i = subAreas_.GetSize();
   while (i--) {
      if ((fExist = subAreas_[i]->ExistArea(areaName))) {
         break;                                 // This Sub-Area matches to the specified name.
      }
   }
   subAreasLock_.Unlock();
   return fExist;
}



//-------------------------------------------------------------------------
// IMPLEMENTATION
//-------------------------------------------------------------------------

//=========================================================================
// GetSubArea
// ----------
//    Searches the Area with the specified ID from the current instance.
//=========================================================================
HRESULT EventArea::GetSubArea( DWORD dwAreaID, EventArea** ppArea )
{
   if (areaId_ == dwAreaID) {
      *ppArea = this;
      return S_OK;
   }

   subAreasLock_.Lock();
   for (int i=0; i < subAreas_.GetSize(); i++) {

      if (SUCCEEDED( subAreas_[i]->GetSubArea( dwAreaID, ppArea ))) {;
         subAreasLock_.Unlock();
         return S_OK;
      }
   }
   subAreasLock_.Unlock();
   return E_INVALID_HANDLE;
}
//DOM-IGNORE-BEGIN

#endif