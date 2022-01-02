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

#ifndef __AttributeValueMap_H
#define __AttributeValueMap_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include <comdef.h>                             // for _variant_t

/**
 * @class	AttributeValueMap
 *
 * @brief	This class implements a map with attribute values and the associoated attribute IDs. 
 * 			Entries can be stored by index. Attribute values can be accessed by ID.
 */

class AttributeValueMap
{
public:

   /**
    * @fn	AttributeValueMap::AttributeValueMap();
    *
    * @brief	Default constructor.
    */

   AttributeValueMap();

   /**
    * @fn	HRESULT AttributeValueMap::Create( DWORD dwSize );
    *
    * @brief	Creates a new Attribute Value Map. 
    * 			Must be called after construction.
    * 			Parameter size specifies the maximum number of elements.
    *
    * @param	size	The size.
    *
    * @return	A hResult.
    */

   HRESULT Create( DWORD size );

   /**
    * @fn	AttributeValueMap::~AttributeValueMap();
    *
    * @brief	Destructor.
    */

   ~AttributeValueMap();

// Operations
public:

   /**
    * @fn	HRESULT AttributeValueMap::SetAtIndex( DWORD index, DWORD attributeId, LPVARIANT variant );
    *
    * @brief	Sets the attribute value/ID pair at specified index in the map.
    * 			The maximum number of elements is specified by the Create() function.
    *
    * @param	index	   	Zero-based index of the.
    * @param	attributeId	Identifier for the attribute.
    * @param	variant	   	The variant.
    *
    * @return	A hResult.
    */

   HRESULT     SetAtIndex( DWORD index, DWORD attributeId, LPVARIANT variant );

   /**
    * @fn	LPVARIANT AttributeValueMap::Lookup( DWORD attributeId );
    *
    * @brief	Returns a pointer to the attribute value associated with the specified attribute ID if available.
    * 			Otheriwse a NULL pointer is returned.
    *
    * @param	attributeId	Identifier for the attribute.
    *
    * @return	A LPVARIANT.
    */

   LPVARIANT   Lookup( DWORD attributeId );

// Impmementation
protected:

   /**
    * @typedef	struct
    *
    * @brief	Defines the Attribute Id/Value structure.
    */

   typedef struct {
      DWORD       attributeId_;
      _variant_t  attributeValue_;
   } ATTRIBUTE, *LPATTRIBUTE;

   /** @brief	The map. */
   LPATTRIBUTE  map_;
   /** @brief	The size. */
   DWORD          size_;
};
//DOM-IGNORE-END

#endif // __AttributeValueMap_H

