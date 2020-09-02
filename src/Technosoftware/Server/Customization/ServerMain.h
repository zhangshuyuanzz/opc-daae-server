/*
 * Copyright (c) 2020 Technosoftware GmbH. All rights reserved
 * Web: https://technosoftware.com 
 * 
 * The source code in this file is covered under a dual-license scenario:
 *   - Owner of a purchased license: RPL 1.5
 *   - GPL V3: everybody else
 *
 * RPL license terms accompanied with this source code.
 * See https://technosoftware.com/license/RPLv15License.txt
 *
 * GNU General Public License as published by the Free Software Foundation;
 * version 3 of the License are accompanied with this source code.
 * See https://technosoftware.com/license/GPLv3License.txt
 *
 * This source code is distributed in the hope that it will be useful, but
 * WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY
 * or FITNESS FOR A PARTICULAR PURPOSE.
 */
                                                                        
#ifndef __APP_MAIN_H_
#define __APP_MAIN_H_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "DaServer.h"
#include "AeServer.h"
#include "Logger.h"

//-------------------------------------------------------------------------
// MACRO
//-------------------------------------------------------------------------
#define CHECK_RESULT(f) {HRESULT hres = f; if (FAILED( hres )) throw hres;}
#define CHECK_PTR(p) {if ((p)== NULL) throw E_OUTOFMEMORY;}
#define CHECK_ADD(f) {if (!(f)) throw E_OUTOFMEMORY;}
#define CHECK_BOOL(f) {if (!(f)) throw E_FAIL;}


#endif // __APP_MAIN_H_
