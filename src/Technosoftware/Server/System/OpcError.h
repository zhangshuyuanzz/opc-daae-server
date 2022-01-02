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

#ifndef  __OPCERROR_H_14EB9DC1_6E64_11d2_88EE_00104B965F5E
#define  __OPCERROR_H_14EB9DC1_6E64_11d2_88EE_00104B965F5E

#ifdef _WIN32

      #include "inc32\opcerror.h"		// IIDs and CLSIDs for OPC Common Definitions
      #include "inc32\opcae_er.h"		// IIDs and CLSIDs for OPC Server List / Enumerator

#else

   #ifdef _WIN64

      #include "inc64\opcerror.h"		// IIDs and CLSIDs for OPC Common Definitions
      #include "inc64\opcae_er.h"		// IIDs and CLSIDs for OPC Server List / Enumerator

   #else

      #error "Platform not supported"

   #endif // _WIN64

#endif // _WIN32

#endif // __OPCERROR_H_14EB9DC1_6E64_11d2_88EE_00104B965F5E
