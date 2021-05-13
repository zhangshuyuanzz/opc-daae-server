/*
 * Copyright (c) 2011-2021 Technosoftware GmbH. All rights reserved
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

#ifndef __EVENTATTRIBUTE_H
#define __EVENTATTRIBUTE_H

 //DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "OpcString.h"

//-----------------------------------------------------------------------
// DEFINES
//-----------------------------------------------------------------------
//    IDs of internal handled attributes
#define  ATTRID_ACKCOMMENT    (0xFFFFFFFE)
#define  ATTRID_AREAS         (0xFFFFFFFD)
//    Indexes of internal handled attributes
#define  ATTRNDX_ACKCOMMENT   (0)
#define  ATTRNDX_AREAS        (1)


//-----------------------------------------------------------------------
// CLASS
//-----------------------------------------------------------------------
class AeAttribute
{
	// Construction
public:
	AeAttribute()
	{
		m_dwAttrID = 0;
		m_vt = VT_EMPTY;
	}
	HRESULT Create(DWORD dwAttrID, LPCWSTR szDescr, VARTYPE vt)
	{
		m_dwAttrID = dwAttrID;
		m_vt = vt;
		m_wsDescr = szDescr;
		return S_OK;
	}

	// Destruction
	~AeAttribute() {};

	// Attributes
public:
	inline DWORD			AttrID()    const { return m_dwAttrID; }
	inline COpcString&		Descr() { return m_wsDescr; }
	inline VARTYPE			VarType()   const { return m_vt; }

	// Implementation
protected:
	DWORD       m_dwAttrID;
	COpcString	m_wsDescr;
	VARTYPE     m_vt;
};
//DOM-IGNORE-END

#endif // __EVENTATTRIBUTE_H
