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
//-------------------------------------------------------------------------
// INLCUDE
//-------------------------------------------------------------------------
#include "stdafx.h"
#include "AeAttributeValueMap.h"

//-------------------------------------------------------------------------
// CODE
//-------------------------------------------------------------------------

//=========================================================================
// Construction
//=========================================================================
AttributeValueMap::AttributeValueMap()
{
   size_ = 0;
   map_   = NULL;
}


HRESULT AttributeValueMap::Create( DWORD size )
{
   map_ = new ATTRIBUTE [size];
   if (!map_) return E_OUTOFMEMORY;
   size_ = size;
   return S_OK;
}



//=========================================================================
// Destructor
//=========================================================================
AttributeValueMap::~AttributeValueMap()
{
   if (map_) {
      delete [] map_;
   }
   size_=0;
   map_= NULL;
}



//-------------------------------------------------------------------------
// OPERATIONS
//-------------------------------------------------------------------------

HRESULT AttributeValueMap::SetAtIndex( DWORD index, DWORD attributeId, LPVARIANT variant)
{
   _ASSERTE(index < size_ );
   try {
      map_[index].attributeId_		= attributeId;
      map_[index].attributeValue_   = variant;
   }
   catch(_com_error e) {
      return e.Error();
   }
   return S_OK;
}


LPVARIANT AttributeValueMap::Lookup( DWORD attributeId)
{
   for (DWORD i = 0; i < size_; i++) {
      if (map_[ i ].attributeId_ == attributeId) {
         return &map_[ i ].attributeValue_;
      }
   }
   return NULL;
}
//DOM-IGNORE-END