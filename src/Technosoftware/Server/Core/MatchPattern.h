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

#ifndef __MatchPattern_H
#define __MatchPattern_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000


// By redefining MCHAR, _M and _ismdigit you may alter the type
// of string MatchPattern() works with. For example to operate on
// wide strings, make the following definitions:
// #define MCHAR		WCHAR
// #define _M(x)		L ## x
// #define _ismdigit   iswdigit



#ifndef MCHAR
                              // The OPC Toolkit requires a wide-char version
#define MCHAR		   WCHAR
#define _M(a)		   L ## a
#define _ismdigit    iswdigit

//#define MCHAR		   TCHAR
//#define _M(a)		   _T(a)
//#define _ismdigit    _istdigit

#endif



extern BOOL  MatchPattern( const MCHAR* String, const MCHAR * Pattern, BOOL bCaseSensitive = FALSE );

//DOM-IGNORE-END

#endif

