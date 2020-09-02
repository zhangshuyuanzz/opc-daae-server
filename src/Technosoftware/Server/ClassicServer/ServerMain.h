/*
 * Copyright (c) 2011-2020 Technosoftware GmbH. All rights reserved
 * Web: https://technosoftware.com 
 * 
 * Purpose: 
 *
 * The Software is subject to the Technosoftware GmbH Source Code License Agreement, 
 * which can be found here:
 * https://technosoftware.com/documents/Source_License_Agreement.pdf
 */
                                                                        
#ifndef __APP_MAIN_H_
#define __APP_MAIN_H_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "DaServer.h"
#include "AeServer.h"

//-------------------------------------------------------------------------
// MACRO
//-------------------------------------------------------------------------
#define CHECK_RESULT(f) {HRESULT hres = f; if (FAILED( hres )) throw hres;}
#define CHECK_PTR(p) {if ((p)== NULL) throw E_OUTOFMEMORY;}
#define CHECK_ADD(f) {if (!(f)) throw E_OUTOFMEMORY;}
#define CHECK_BOOL(f) {if (!(f)) throw E_FAIL;}


#endif // __APP_MAIN_H_
