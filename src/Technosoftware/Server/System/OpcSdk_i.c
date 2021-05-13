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

#ifndef  __OPCSDK_I_C_
#define  __OPCSDK_I_C_

#ifdef _WIN32

      #include "inc32\opccomn_i.c"         // IIDs and CLSIDs for OPC Common Definitions
      #include "inc32\OpcEnum_i.c"         // IIDs and CLSIDs for OPC Server List / Enumerator
      #include "inc32\opcauto10a_i.c"      // IIDs and CLSIDs for OPC 1.0A Automation Data Access
      #include "inc32\opcda_i.c"           // IIDs and CLSIDs for OPC Data Access
      #include "inc32\opc_ae_i.c"          // IIDs and CLSIDs for OPC Alarsm & Events

#else

   #ifdef _WIN64

      #include "inc64\opccomn_i.c"         // IIDs and CLSIDs for OPC Common Definitions
      #include "inc64\OpcEnum_i.c"         // IIDs and CLSIDs for OPC Server List / Enumerator
      #include "inc64\opcda_i.c"           // IIDs and CLSIDs for OPC Data Access
      #include "inc64\opc_ae_i.c"          // IIDs and CLSIDs for OPC Alarsm & Events

   #else

      #error "Platform not supported"

   #endif // _WIN64

#endif // _WIN32
#endif // __OPCSDK_I_C_
