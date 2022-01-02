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

#ifndef __DaPublicGroupManager_H_
#define __DaPublicGroupManager_H_

//DOM-IGNORE-BEGIN

#include "DaPublicGroup.h"
#include "openarray.h"


/////////////////////////////////////////////////////////////////
//
// DaPublicGroupManager:
// ==================
// handles a collection of public groups (type DaPublicGroup)
// once a Public Group is added to the handler it is marked
// as owned and the handler will take care of it, included
// it's deletion
//
// Implementation:
// is implemented as an array
// there is a semaphore for each group to avoid race conditions
// so a group cannot be removed while someone is using Next
//
/////////////////////////////////////////////////////////////////

class DaPublicGroupManager {
   friend DaPublicGroup;

public:

   DaPublicGroupManager();
   ~DaPublicGroupManager();

      // returns the handle of the added Public Group
   HRESULT AddGroup( DaPublicGroup *pGroup, long *hPGHandle );

      // removes the Public Group with the given handle;
      // if the group is accessed by other threads then
      // the group is phisically removed only when the 
      // last has called ReleaseGroup
   HRESULT RemoveGroup( long hPGHandle );

      // returns the group with the given handle and increments ref count
   HRESULT GetGroup( long hPGHandle, DaPublicGroup **pGroup );

      // decrements the ref count on the group 
   HRESULT ReleaseGroup( long hPGHandle );

      // returns the group with the given name
   HRESULT FindGroup( BSTR Name, long *hPGHandle );

      // returns the first group available
   HRESULT First( long *hPGHandle );

      // starting from the next
   HRESULT Next( long hStartHandle, long *hNextHandle );
   
      // the number of public groups
   HRESULT TotGroups( long *nrGroups );

private:
   HRESULT FindGroupNoLock( BSTR Name, long *hPGHandle );

private:
   long m_TotGroups;

   OpenArray<DaPublicGroup *> m_array;

      // used to synchronize access to m_array
   CRITICAL_SECTION m_CritSec;
};
//DOM-IGNORE-END

#endif // __DaPublicGroupManager_H_
