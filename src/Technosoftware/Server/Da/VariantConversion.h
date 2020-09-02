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

#ifndef __VARIANTCONVERSION_H
#define __VARIANTCONVERSION_H

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

HRESULT VariantFromVariant(
            /* [out] */    LPVARIANT pvDest,
            /* [in]  */    VARTYPE vtRequested,
            /* [in]  */    const LPVARIANT pvSrc );
//DOM-IGNORE-END

#endif
