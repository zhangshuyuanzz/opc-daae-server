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

#include "StdAfx.h"
#include "OpcText.h"

//==============================================================================
// COpcText

// Constructor
COpcText::COpcText()
{
   Reset();
}

// Reset
void COpcText::Reset()
{
   m_cData.Empty();

   // Search Crtieria
   m_eType = COpcText::NonWhitespace;
   m_cHaltChars.Empty();
   m_uMaxChars = 0;
   m_bNoExtract = false;
   m_cText.Empty();
   m_bSkipLeading = false;
   m_bSkipWhitespace = false;
   m_bIgnoreCase = false;
   m_bEofDelim = false;
   m_bNewLineDelim = false;
   m_cDelims.Empty();
   m_bLeaveDelim = true;
   m_zStart = L'"';
   m_zEnd = L'"';
   m_bAllowEscape = true;

   // Search Results
   m_uStart = 0;
   m_uEnd = 0;
   m_zHaltChar = 0;
   m_uHaltPos = 0;
   m_zDelimChar = 0;
   m_bEof = false;
   m_bNewLine = false;
}

// CopyData
void COpcText::CopyData(LPCWSTR szData, UINT uLength)
{
    m_cData.Empty();

    if (uLength > 0 && szData != NULL)
    {
        LPWSTR wszData = OpcArrayAlloc(WCHAR, uLength+1);
        wcsncpy(wszData, szData, uLength);
        wszData[uLength] = L'\0';
        
        m_cData = wszData;
        OpcFree(wszData);
    }
}

// SetType
void COpcText::SetType(COpcText::Type eType)
{
   Reset();

   m_eType = eType;

   switch (eType)
   {
      case Literal:
      {
         m_cText.Empty();
         m_bSkipLeading = false;
         m_bSkipWhitespace = true;
         m_bIgnoreCase = false;
         break;
      }

      case Whitespace:
      {
         m_bSkipLeading = false;
         m_bEofDelim = true;
         break;
      }

      case NonWhitespace:
      {
         m_bSkipWhitespace = true;
         m_bEofDelim = true;
         break;
      }

      case Delimited:
      {
         m_cDelims.Empty();
         m_bSkipWhitespace = false;
         m_bIgnoreCase = false;
         m_bEofDelim = false;
         m_bNewLineDelim = false;
         m_bLeaveDelim = false;
         break;
      }

      default:
      {
         break;
      }
   }
}
