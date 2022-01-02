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

#if !defined(AFX_STDAFX_H__5F66E434_FC32_11D0_A25F_0000E81E9085__INCLUDED_)
#define AFX_STDAFX_H__5F66E434_FC32_11D0_A25F_0000E81E9085__INCLUDED_

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#define WIN32_LEAN_AND_MEAN      // Exclude not used services from Windows headers.
#define _WIN32_WINNT 0x0501
#define _ATL_APARTMENT_THREADED
#define _ATL_STATIC_REGISTRY     // Forces static linking of the registar module.

#include <windows.h>
#include <comdef.h>                 // For _variant_t and _bstr_t
#include <crtdbg.h>                 // For _ASSERTE

#include <atlbase.h>
#include "CoreGenericMain.h"    // Defines the global _Module used by ATL.
#include "opcauto10a.h"          // Inlcude the modified header 
                                 // Must be included before OpcSdk.h !
#include "OpcSdk.h"               // Definition of all supported Interfaces.

//#define _ATL_DEBUG_QI          // Display ATL Debug Informations
#include <atlcom.h>

#include "OpcError.h"

//----------------------------------------------------------------------------
// The __stdcall calling convention must be used for all
// functions exported by this DLL.
//----------------------------------------------------------------------------
#define DLLEXP
#define DLLCALL __stdcall

#ifdef _OPC_DLL
#define DLLIMP __declspec(dllimport)
#endif

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

//DOM-IGNORE-END

#endif // !defined(AFX_STDAFX_H__5F66E434_FC32_11D0_A25F_0000E81E9085__INCLUDED)
