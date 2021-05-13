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

#ifndef __EVENTCONDITIONDEF_H
#define __EVENTCONDITIONDEF_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "OpcString.h"

class AeCategory;


//-----------------------------------------------------------------------
// CLASS AeSubConditionDefiniton
//-----------------------------------------------------------------------
class AeSubConditionDefiniton
{
// Construction
public:
   AeSubConditionDefiniton();
   HRESULT Create( LPCWSTR szName, LPCWSTR szDef,
                   DWORD dwSeverity, LPCWSTR szDescr,
                   BOOL fAckRequired );

// Destruction
   ~AeSubConditionDefiniton();

// Attributes
public:
   inline COpcString&  Name()               { return m_wsName;      }
   inline COpcString&  Definition()         { return m_wsDef;       }
   inline DWORD         Severity() const     { return m_dwSeverity;  }
   inline COpcString&  Description()        { return m_wsDescr;     }
   inline BOOL          AckRequired() const  { return m_fAckRequired;}

// Implementation
protected:
	COpcString    m_wsName;
	COpcString    m_wsDef;
	DWORD          m_dwSeverity;
	COpcString    m_wsDescr;
	BOOL           m_fAckRequired;               // Defines if acknowledge is required if the
                                                // sub condition becomes active.
};



//-----------------------------------------------------------------------
// CLASS AeConditionDefiniton
//-----------------------------------------------------------------------
class AeConditionDefiniton {
// Construction
public:
   AeConditionDefiniton();                        // Default Constructor
                                                // Single State Condition
   HRESULT Create( AeCategory* pCategory, DWORD m_dwCondDefID,
                   LPCWSTR szName, LPCWSTR szDef, DWORD dwSeverity, LPCWSTR szDescr,
                   BOOL fAckRequired );

                                                // Multi State Condition
   HRESULT Create( AeCategory* pCategory, DWORD dwCondDefID, LPCWSTR szName );
   HRESULT AddSubCondDef( DWORD dwSubCondDefID, LPCWSTR szName, LPCWSTR szDef,
                          DWORD dwSeverity, LPCWSTR szDescr, BOOL fAckRequired );

// Destruction
public:
   ~AeConditionDefiniton();

// Attributes
public:
   inline DWORD CondDefID() const { return m_dwCondDefID; }
   inline COpcString& Name() { return m_wsName; }
   inline AeCategory* Category() { return m_pCategory; }
   inline AeSubConditionDefiniton* DefaultSubCondition() { return m_mapSubCond.m_aVal[0]; }
   inline AeSubConditionDefiniton* SubCondition( DWORD dwSubCondDefID ) { return m_mapSubCond.Lookup( dwSubCondDefID ); }
   inline BOOL IsMultiState() const { return (m_mapSubCond.GetSize() > 1) ? TRUE : FALSE; }
                                       
// Operations
public:

   HRESULT GetSubConditionInfo(  LPDWORD pdwNumSCs,
                                 LPWSTR** pszSCNames, LPWSTR** pszSCDefinitions,
                                 LPDWORD* pdwSCSeverities, LPWSTR** pszSCDescriptions );

   HRESULT QuerySubConditionNames(
   /* [out] */                   DWORD       *  pdwCount,
   /* [size_is][size_is][out] */ LPWSTR      ** ppszSubConditionNames );

   HRESULT GetSubCondDefID(
   /* [in] */                    LPCWSTR szSubCondName,
   /* [out] */                   LPDWORD pdwSubCondDefID );

   HRESULT GetSubCondDefID(
   /* [in] */                    AeSubConditionDefiniton* pSubCond,
   /* [out] */                   LPDWORD pdwSubCondDefID );

// Implementation
protected:
   // Data Members

   CComAutoCriticalSection m_csMem;             // lock/unlock data members
   DWORD                   m_dwCondDefID;
   AeCategory*         m_pCategory;
   COpcString             m_wsName;

   CSimpleMap<DWORD, AeSubConditionDefiniton*> m_mapSubCond;
};
//DOM-IGNORE-END

#endif // __EVENTCONDITIONDEF_H

