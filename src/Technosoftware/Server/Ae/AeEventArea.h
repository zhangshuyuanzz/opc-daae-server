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

#ifndef __EventArea_H
#define __EventArea_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "enumclass.h"
#include "WideString.h"
#include "AeSource.h"

class EventAreaBrowser;

//-----------------------------------------------------------------------
// DEFINES
//-----------------------------------------------------------------------
/** @brief	The root Area Id. */
#define  AREAID_ROOT          (0xFFFFFFFE)
/** @brief	The Unspecified Area Id. */
#define  AREAID_UNSPECIFIED   (0xFFFFFFFD)

//-----------------------------------------------------------------------
// CLASS
//-----------------------------------------------------------------------
class EventArea
{
// Construction / Destruction
public:

   /**
    * @fn	EventArea::EventArea();
    *
    * @brief	Default constructor.
    */

   EventArea();

   /**
    * @fn	HRESULT EventArea::Create( EventArea* parent, DWORD areaId, LPCWSTR areaName );
    *
    * @brief	Must be called after construction.
    * 			Initializes the new area and add it as sub-area to the parent area.
    *
    * @param [in,out]	parent	If non-null, the parent.
    * @param	areaId		   	Identifier for the area.
    * @param	areaName	   	Name of the area.
    *
    * @return	A hResult.
    */

   HRESULT Create(EventArea* parent, DWORD areaId, LPCWSTR areaName);

   /**
    * @fn	EventArea::~EventArea();
    *
    * @brief	Destructor.
    */

   ~EventArea();

// Attributes
public:
   inline EventArea*	Parent() const    { return parent_; }
   inline EventArea*    Root()   const    { return root_; }
   inline BOOL          IsRoot() const    { return (root_ == this) ? TRUE : FALSE; }
   inline WideString&   Name()            { return name_; }
   inline LPCWSTR       PartialName()     { return partialName_; }
   inline LPCWSTR       Delimiter() const { return delimiter_; }

// Operations
public:

   /**
    * @fn	HRESULT EventArea::PositionDown( LPCWSTR areaName, EventArea** area );
    *
    * @brief	Gets the specified sub-area to move into.
    *
    * @param	areaName		Name of the area.
    * @param [in,out]	area	If non-null, the area.
    *
    * @return	A hResult.
    */

   HRESULT  PositionDown( LPCWSTR areaName, EventArea** area );

   /**
    * @fn	HRESULT EventArea::PositionTo( LPCWSTR areaName, EventArea** area );
    *
    * @brief	Gets the specified area from the hierarchical event area space to move into.
    *
    * @param	areaName		Name of the area.
    * @param [in,out]	area	If non-null, the area.
    *
    * @return	A hResult.
    */

   HRESULT  PositionTo( LPCWSTR areaName, EventArea** area );

   /**
    * @fn	HRESULT EventArea::GetQualifiedAreaName( LPCWSTR areaName, LPWSTR* qualifiedAreaName );
    *
    * @brief	Gets the fully qualified area name from a partial sub-area name.
    *
    * @param	areaName				 	Name of the area.
    * @param [in,out]	qualifiedAreaName	If non-null, name of the qualified area.
    *
    * @return	The qualified area name.
    */

   HRESULT  GetQualifiedAreaName( LPCWSTR areaName, LPWSTR* qualifiedAreaName );

   /**
    * @fn	HRESULT EventArea::GetQualifiedSourceName( LPCWSTR sourceName, LPWSTR* qualifiedSourceName );
    *
    * @brief	Gets the fully qualified source name from a partial source name.
    *
    * @param	sourceName				   	Name of the source.
    * @param [in,out]	qualifiedSourceName	If non-null, name of the qualified source.
    *
    * @return	The qualified source name.
    */

   HRESULT  GetQualifiedSourceName( LPCWSTR sourceName, LPWSTR* qualifiedSourceName );

   /**
    * @fn	HRESULT EventArea::BrowseAreas( EventAreaBrowser* areaBrowser, COpcComEnumString* comEnum, OPCAEBROWSETYPE browseFilterType, LPCWSTR filterCriteria );
    *
    * @brief	Initializes the string enumerator object with the name of the sub-areas or the name of the sources.
    *
    * @param [in,out]	areaBrowser	If non-null, the area browser.
    * @param [in,out]	comEnum	   	If non-null, the com enum.
    * @param	browseFilterType   	Type of the browse filter. OPC_AREA or OPC_SOURCE
    * @param	filterCriteria	   	The filter criteria.
    *
    * @return	A hResult.
    */

   HRESULT  BrowseAreas( EventAreaBrowser* areaBrowser, COpcComEnumString* comEnum,
                         OPCAEBROWSETYPE browseFilterType, LPCWSTR filterCriteria );

   /**
    * @fn	static void EventArea::SetDelimiter( WCHAR wc )
    *
    * @brief	Functions called from the Server Class Handler.
    *
    * @param	wc	The wc.
    */

   static void SetDelimiter( WCHAR delimiter ) {
      delimiter_[0] = delimiter;
   }

   /**
    * @fn	HRESULT EventArea::GetArea( DWORD areaId, EventArea** area );
    *
    * @brief	Searches the Area with the specified ID from the entire Area Space.
    *
    * @param	areaId			Identifier for the area.
    * @param [in,out]	area	If non-null, the area.
    *
    * @return	The area.
    */

   HRESULT  GetArea( DWORD areaId, EventArea** area );

   /**
    * @fn	HRESULT EventArea::AttachSource( AeSource* source );
    *
    * @brief	Attaches a link to source to this area.
    *
    * @param [in,out]	source	If non-null, source for the.
    *
    * @return	A hResult.
    */

   HRESULT  AttachSource( AeSource* source );

   /**
    * @fn	HRESULT EventArea::DetachSource( AeSource* source );
    *
    * @brief	Detaches a link to a source from this area.
    *
    * @param [in,out]	source	If non-null, source for the.
    *
    * @return	A hResult.
    */

   HRESULT  DetachSource( AeSource* source );

   /**
    * @fn	HRESULT EventArea::EnableConditions( BOOL enable, CSimpleValArray<AeEventArray*>& events );
    *
    * @brief	Enables or disables all conditions for all sources within this area.
    *
    * @param	enable		  	true to enable, false to disable.
    * @param [in,out]	events	[in,out] Array of pointers to arrays with created event instances 
    * 							for each condition with changed enable state.
    *
    * @return	A hResult.		S_OK                    All succeeded
    * 							S_FALSE                 Not succeeded for all conditions.
    * 							E_XXX                   Error occured.    
    */

   HRESULT  EnableConditions( BOOL enable, CSimpleValArray<AeEventArray*>& events );

   /**
    * @fn	BOOL EventArea::ExistArea( LPCWSTR areaName );
    *
    * @brief	Checks if the name of this area or the name of a a sub area matches the specified name. 
    * 			It is possible to specify an area name using the wildcard sytax described in Appendix A 
    * 			of the AE Specification.
    *
    * @param	areaName	Name of the area.
    *
    * @return	true if it succeeds, false if it fails.
    */

   BOOL     ExistArea( LPCWSTR areaName );

   /**
    * @fn	BOOL EventArea::FilterName( LPCWSTR name, LPCWSTR filterCriteria );
    *
    * @brief	Filters an Process Area or Event Source name with the specified filter. 
    * 			This function is called if the client browse the Process Area Space with the function 
    * 			IOPCEventAreaBrowser::BrowseOPCAreas().
    * 			
    * 			Function located in the server specific part.
    *
    * @param	name		  	The Process Area or Event Source name
    * @param	filterCriteria	A filter string.
    *
    * @return	If name matches filterCriteria, return TRUE; if there is no match, return is FALSE. 
    * 			If either name or filterCriteria is a NULL pointer, return is FALSE.
    */

   BOOL     FilterName( LPCWSTR name, LPCWSTR filterCriteria );

// Implementation
protected:

   /**
    * @fn	HRESULT EventArea::GetSubArea( DWORD dwAreaID, EventArea** ppArea );
    *
    * @brief	Searches the Area with the specified ID from the current instance.
    *
    * @param	dwAreaID	  	Identifier for the area.
    * @param [in,out]	ppArea	If non-null, the area.
    *
    * @return	The sub area.
    */

   HRESULT GetSubArea( DWORD dwAreaID, EventArea** ppArea );

   static WCHAR					delimiter_[2];
   DWORD						areaId_;
   EventArea*					parent_;
   EventArea*					root_;
   WideString					name_;			// The fully qualified name of the area
   LPCWSTR						partialName_;   // Points to the partially area name in
                                                // name_. Do not free this string !

   /** @brief	Array with the Sub-Areas of this area. */
   CSimpleArray<EventArea*>     subAreas_;

   /** @brief	Array with the Sources of this area. */
   CSimpleArray<AeSource*>	sources_;

   /** @brief	The critical section to lock/unlock the array of Sub-Areas. */
   CComAutoCriticalSection		subAreasLock_;

   /** @brief	The critical section to lock/unlock the array of Sources. */
   CComAutoCriticalSection		sourcesLock_;
};
//DOM-IGNORE-END

#endif // __EventArea_H

