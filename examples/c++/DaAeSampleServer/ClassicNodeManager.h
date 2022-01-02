/*
 * Copyright (c) 2011-2022 Technosoftware GmbH. All rights reserved
 * Web: http://www.technosoftware.com 
 * 
 * License:
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:

 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.

 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 * SPDX-License-Identifier: MIT
 */

#if !defined(CLASSICNODEMANAGER_H)
#define CLASSICNODEMANAGER_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/*
 * ----- BEGIN SAMPLE IMPLEMENTATION -----
 */

/*
 * Application Definitions (SAMPLE)
 */
#define UPDATE_PERIOD         200            /* Data Cache update rate in milliseconds */


/*
 * Signal ( Item ) Types (SAMPLE)
 /
#define SIGMASK_INTERN_X      0x010000       /* static In/Out/InOut */
#define SIGMASK_INTERN_IN     0x110000       /* input */
#define SIGMASK_INTERN_OUT    0x210000       /* output */
#define SIGMASK_INTERN_INOUT  0x410000       /* readable output */
#define SIGMASK_SIMULATED     0x510000       /* internal item with simulated data */
#define SIGMASK_REMOVABLE     0x610000       /* internal removable item */

/*
 * ----- END SAMPLE IMPLEMENTATION ----- 
 */

#endif // !defined(CLASSICNODEMANAGER_H)
