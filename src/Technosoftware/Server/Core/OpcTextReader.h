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

#ifndef _COpcTextReader_H
#define _COpcTextReader_H

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "OpcDefs.h"
#include "OpcText.h"

//==============================================================================
// CLASS:   COpcTextReader
// PURPOSE: Extracts tokens from a stream.

class OPCUTILS_API COpcTextReader
{
    OPC_CLASS_NEW_DELETE();

public:

    //==========================================================================
    // Operators

    // Constructor
    COpcTextReader(const COpcString& cBuffer);  
    COpcTextReader(LPCSTR szBuffer, UINT uLength = -1);  
    COpcTextReader(LPCWSTR szBuffer, UINT uLength = -1);  
 
    // Destructor
    ~COpcTextReader(); 

    //==========================================================================
    // Public Methods
  
    // GetNext
    bool GetNext(COpcText& cText);

    // GetBuf
    LPCWSTR GetBuf() const { return m_szBuf; }

private:

    //==========================================================================
    // Private Methods

    // ReadData
    bool ReadData();

    // FindToken
    bool FindToken(COpcText& cText);

    // FindLiteral
    bool FindLiteral(COpcText& cText);

    // FindNonWhitespace
    bool FindNonWhitespace(COpcText& cText);

    // FindWhitespace
    bool FindWhitespace(COpcText& cText);
    
    // FindDelimited
    bool FindDelimited(COpcText& cText);

    // FindEnclosed
    bool FindEnclosed(COpcText& cText);

    // CheckForHalt
    bool CheckForHalt(COpcText& cText, UINT uIndex);
    
    // CheckForDelim
    bool CheckForDelim(COpcText& cText, UINT uIndex);

    // SkipWhitespace
    UINT SkipWhitespace(COpcText& cText);

    // CopyData
    void CopyData(COpcText& cText, UINT uStart, UINT uEnd);

    //==========================================================================
    // Private Members

    LPWSTR m_szBuf;
    UINT   m_uLength;
    UINT   m_uEndOfData;
};

#endif //ndef _COpcTextReader_H
