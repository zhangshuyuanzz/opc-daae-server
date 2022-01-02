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

#ifndef _OpcUtils_H_
#define _OpcUtils_H_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "opccomn.h"

#define OPCUTILS_API

#include "OpcDefs.h"
#include "COpcString.h"
#include "COpcFile.h"
#include "OpcMatch.h"
#include "COpcCriticalSection.h"
#include "COpcArray.h"
#include "COpcList.h"
#include "COpcMap.h"
#include "COpcSortedArray.h"
#include "COpcText.h"
#include "COpcTextReader.h"
#include "COpcThread.h"
#include "OpcCategory.h"
#include "OpcRegistry.h"
#include "OpcXmlType.h"
#include "COpcXmlAnyType.h"
#include "COpcXmlElement.h"
#include "COpcXmlDocument.h"
#include "COpcVariant.h"
#include "COpcComObject.h"
#include "COpcClassFactory.h"
#include "COpcComModule.h"
#include "COpcCommon.h"
#include "COpcConnectionPoint.h"
#include "COpcCPContainer.h"
#include "COpcEnumCPs.h"
#include "COpcEnumString.h"
#include "COpcEnumUnknown.h"
#include "COpcSecurity.h"
#include "COpcThreadPool.h"
#include "COpcBrowseElement.h"

#endif // _OpcUtils_H_

