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

#ifndef _COpcMap_H_
#define _COpcMap_H_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "OpcDefs.h"
#include "OpcString.h"

//==============================================================================
// TYPE:    OPC_POS
// PURPOSE: A position when enumerating the hash table.

#ifndef _OPC_POS
#define _OPC_POS
typedef struct TOpcPos{}* OPC_POS;
#endif //_OPC_POS

//==============================================================================
// CLASS:   COpcMap<KEY, VALUE>
// PURPOSE: Defines a hash table template class.

template<class KEY,class VALUE>
class COpcMap 
{
   OPC_CLASS_NEW_DELETE_ARRAY()

public:

    //==========================================================================
    // COpcEntry
    class COpcEntry
    {
        OPC_CLASS_NEW_DELETE_ARRAY()

    public:

        COpcEntry* pNext;
        KEY        cKey;
        VALUE      cValue;
        UINT       uHash;

        COpcEntry() : pNext(NULL), uHash(0) {}
    };

    //==========================================================================
    // COpcBlock
    class COpcBlock
    {
        OPC_CLASS_NEW_DELETE_ARRAY()

    public:

        COpcBlock(UINT uBlockSize)
        :
           pNext(NULL),
           pEntries(NULL)
        {
            pEntries = new COpcEntry[uBlockSize];
        }

        ~COpcBlock()
        {
            delete [] pEntries;
        }

        COpcBlock* pNext;
        COpcEntry* pEntries;
    };

public:

    //==========================================================================
    // Constructor
    COpcMap(int nBlockSize = 10, int uTableSize = 17)
    :
        m_pUnusedEntries(NULL),
        m_uCount(0),
        m_pBlocks(NULL),
        m_uBlockSize(nBlockSize),
        m_ppHashTable(NULL),
        m_uTableSize(0)
    {
        InitHashTable(uTableSize);
    }

    //==========================================================================
    // Copy Constructor
    COpcMap(const COpcMap& cMap)
    :
        m_pUnusedEntries(NULL),
        m_uCount(0),
        m_pBlocks(NULL),
        m_uBlockSize(cMap.m_uBlockSize),
        m_ppHashTable(NULL),
        m_uTableSize(0)
    {
        *this = cMap;
    }

    //==========================================================================
    // Destructor
    ~COpcMap()
    {
        RemoveAll();
        delete [] m_ppHashTable;
    }

    //==========================================================================
    // Assignment
    COpcMap& operator=(const COpcMap& cMap)
    {
        InitHashTable(cMap.m_uTableSize);

        KEY cKey;
        VALUE cValue;
        OPC_POS pos = cMap.GetStartPosition();

        while (pos != NULL)
        {
            cMap.GetNextAssoc(pos, cKey, cValue);
            SetAt(cKey, cValue);
        }

        return *this;
    }

    //==========================================================================
    // GetCount
    int GetCount() const
    {
        return m_uCount;
    }

    //==========================================================================
    // IsEmpty
    bool IsEmpty() const
    {
        return (m_uCount == 0);
    }

    //==========================================================================
    // Lookup - return false if not there
    bool Lookup(const KEY& cKey, VALUE** ppValue = NULL) const
    {
        COpcEntry* pEntry = Find(cKey);

        if (pEntry == NULL)
        {
            return false;
        }

        if (ppValue != NULL)
        {
            *ppValue = &(pEntry->cValue);
        }

        return true;
    }

    //==========================================================================
    // Lookup - return false if not there
    bool Lookup(const KEY& cKey, VALUE& cValue) const
    {
        COpcEntry* pEntry = Find(cKey);

        if (pEntry == NULL)
        {
            return false;
        }

        cValue = pEntry->cValue;
        return true;
    }

    //==========================================================================
    // Lookup - and add if not there
    VALUE& operator[](const KEY& cKey)
    {
        COpcEntry* pEntry = Find(cKey);

        if (pEntry == NULL)
        {
            pEntry = NewEntry(cKey);
        }

        return pEntry->cValue;
    }

    //==========================================================================
    // SetAt - add a new (key, value) pair
    void SetAt(const KEY& cKey, const VALUE& cValue)
    {
        (*this)[cKey] = cValue;
    }

    //==========================================================================
    // RemoveKey - removing existing (key, ?) pair
    bool RemoveKey(const KEY& cKey)
    {
        UINT uBin = HashKey(cKey)%m_uTableSize;
        COpcEntry* pEntry = m_ppHashTable[uBin];
        COpcEntry* pPrev  = NULL;

        while (pEntry != NULL)
        {
            if (pEntry->cKey == cKey)
            {
                if (pPrev == NULL)
                {
                    m_ppHashTable[uBin] = pEntry->pNext;
                }
                else
                {
                    pPrev->pNext = pEntry->pNext;
                }

                FreeEntry(pEntry);
                return true;
            }

            pPrev  = pEntry;
            pEntry = pEntry->pNext;
        }

        return false;
    }

    //==========================================================================
    // RemoveAll
    void RemoveAll()
    {
        COpcBlock* pBlock = m_pBlocks;
        COpcBlock* pNext  = NULL;

        while (pBlock != NULL)
        {
            pNext = pBlock->pNext;
            delete pBlock;
            pBlock = pNext;
        }

        m_uCount   = 0;
        m_pUnusedEntries = NULL;
        m_pBlocks  = NULL;
        memset(m_ppHashTable, 0, m_uTableSize*sizeof(COpcEntry*));
    }

    //==========================================================================
    // GetStartPosition
    OPC_POS GetStartPosition() const
    {
        UINT uBin = 0;

        while (uBin < m_uTableSize && m_ppHashTable[uBin] == NULL) 
        {
            uBin++;
        }

        if (uBin == m_uTableSize)
        {
            return (OPC_POS)NULL;
        }

        return (OPC_POS)m_ppHashTable[uBin];
    }

    //==========================================================================
    // GetNextAssoc
    void GetNextAssoc(OPC_POS& pos, KEY& cKey) const
    {
        COpcEntry* pEntry = (COpcEntry*)pos;
        OPC_ASSERT(pos != NULL);

        cKey = pEntry->cKey;

        pos = GetNextAssoc(pos);
    }

    //==========================================================================
    // GetNextAssoc
    void GetNextAssoc(OPC_POS& pos, KEY& cKey, VALUE& cValue) const
    {
        COpcEntry* pEntry = (COpcEntry*)pos;
        OPC_ASSERT(pos != NULL);

        cKey = pEntry->cKey;
        cValue = pEntry->cValue;

        pos = GetNextAssoc(pos);
    }

    //==========================================================================
    // GetNextAssoc
    void GetNextAssoc(OPC_POS& pos, KEY& cKey, VALUE*& pValue) const
    {
        COpcEntry* pEntry = (COpcEntry*)pos;
        OPC_ASSERT(pos != NULL);

        cKey = pEntry->cKey;
        pValue = &(pEntry->cValue);

        pos = GetNextAssoc(pos);
    }
    
	//==========================================================================
    // IsValid
    bool IsValid(OPC_POS pos) const
    {
        COpcBlock* pBlock = m_pBlocks;

        while (pBlock != NULL)
        {			
			for (UINT ii = 0; ii < m_uBlockSize; ii++)
            {
				if (pos == (OPC_POS)&(pBlock->pEntries[ii]))
				{
					return true;
				}
            }

            pBlock = pBlock->pNext;
        }

		return false;
    }
    
	//==========================================================================
    // GetPosition
    OPC_POS GetPosition(const KEY& cKey) const
    {
        return (OPC_POS)Find(cKey);
    }

    //==========================================================================
    // GetKey
    const KEY& GetKey(OPC_POS pos) const
    {
        OPC_ASSERT(IsValid(pos));
        return ((COpcEntry*)pos)->cKey;
    }
    
	//==========================================================================
    // GetValue
    VALUE& GetValue(OPC_POS pos)
    {
        OPC_ASSERT(IsValid(pos));
        return ((COpcEntry*)pos)->cValue;
    }

    //==========================================================================
    // GetValue
    const VALUE& GetValue(OPC_POS pos) const
    {
        OPC_ASSERT(IsValid(pos));
        return ((COpcEntry*)pos)->cValue;
    }

    //==========================================================================
    // GetBlockSize
    UINT GetBlockSize() const
    {
        return m_uBlockSize;
    }

    //==========================================================================
    // GetHashTableSize
    UINT GetHashTableSize() const
    {
        return m_uTableSize;
    }

    //==========================================================================
    // InitHashTable
    void InitHashTable(UINT uTableSize)
    {
        COpcEntry* pEntries = NULL;

        for (UINT ii = 0; ii < m_uTableSize; ii++)
        {
            COpcEntry* pEntry = m_ppHashTable[ii];
            COpcEntry* pNext  = NULL;

            while (pEntry != NULL)
            {
                pNext         = pEntry->pNext;
                pEntry->pNext = pEntries;
                pEntries      = pEntry;
                pEntry        = pNext;
            }
        }

        delete [] m_ppHashTable;

        m_uTableSize = uTableSize;
        m_ppHashTable = new COpcEntry*[m_uTableSize];
        memset(m_ppHashTable, 0, m_uTableSize*sizeof(COpcEntry*));

        COpcEntry* pEntry = pEntries;
        COpcEntry* pNext  = NULL;

        while (pEntry != NULL)
        {
            pNext = pEntry->pNext;

            UINT uBin = (pEntry->uHash)%m_uTableSize;
            pEntry->pNext       = m_ppHashTable[uBin];
            m_ppHashTable[uBin] = pEntry;

            pEntry = pNext;
        }
    }

private:

    //==========================================================================
    // Find
    COpcEntry* Find(const KEY& cKey) const
    {
        UINT uBin = HashKey(cKey)%m_uTableSize;
        COpcEntry* pEntry = m_ppHashTable[uBin];

        while (pEntry != NULL)
        {
            if (pEntry->cKey == cKey)
            {
                return pEntry;
            }

            pEntry = pEntry->pNext;
        }

        return NULL;
    }

    //==========================================================================
    // NewEntry
    COpcEntry* NewEntry(const KEY& cKey)
    {
        // optimize hash table size.
        if (m_uTableSize < 1.2*m_uCount)
        {
            InitHashTable(2*m_uCount);
        }

        // create a new block if necessary.
        if (m_pUnusedEntries == NULL)
        {
            COpcBlock* pBlock = new COpcBlock(m_uBlockSize);

            for (UINT ii = 0; ii < m_uBlockSize; ii++)
            {
                pBlock->pEntries[ii].pNext = m_pUnusedEntries;
                m_pUnusedEntries           = &(pBlock->pEntries[ii]);
            }

            pBlock->pNext = m_pBlocks;
            m_pBlocks     = pBlock;
        }

        OPC_ASSERT(m_pUnusedEntries != NULL); 

        // remove entry from unused entry list.
        COpcEntry* pEntry   = m_pUnusedEntries;
        m_pUnusedEntries   = m_pUnusedEntries->pNext;

        // insert entry into hash table.
        pEntry->cKey  = cKey;
        pEntry->uHash = HashKey(cKey);

        UINT uBin           = pEntry->uHash%m_uTableSize;
        pEntry->pNext       = m_ppHashTable[uBin];
        m_ppHashTable[uBin] = pEntry;

        m_uCount++;

        return pEntry;
    }

    //==========================================================================
    // FreeEntry
    void FreeEntry(COpcEntry* pEntry)
    {
        // return to unused entries list
        pEntry->pNext    = m_pUnusedEntries;
        m_pUnusedEntries = pEntry;
        m_uCount--;
        OPC_ASSERT(m_uCount >= 0);  // make sure we don't underflow

        // if no more elements, cleanup completely
        if (m_uCount == 0)
        {
            RemoveAll();
        }
    }

    //==========================================================================
    // GetNextAssoc
    OPC_POS GetNextAssoc(OPC_POS pos) const
    {
        COpcEntry* pEntry = (COpcEntry*)pos;
        OPC_ASSERT(pos != NULL);

        if (pEntry->pNext == NULL)
        {
            UINT uBin = pEntry->uHash%m_uTableSize;

            do
            {
                uBin++;
            }
            while (uBin < m_uTableSize && m_ppHashTable[uBin] == NULL);

            if (uBin == m_uTableSize)
            {
                return (OPC_POS)NULL;
            }

            return (OPC_POS)m_ppHashTable[uBin];
        }

        return (OPC_POS)pEntry->pNext;
    }

    //==========================================================================
    // Members
    COpcEntry** m_ppHashTable;
    UINT        m_uTableSize;
    UINT        m_uCount;
    COpcEntry*  m_pUnusedEntries;
    COpcBlock*  m_pBlocks;
    UINT        m_uBlockSize;
};


//==============================================================================
// FUNCTION: HashKey<KEY>
// PURPOSE:  Default hash key generator.
template<class KEY>
inline UINT HashKey(const KEY& cKey)
{
	return ((UINT)(void*)(DWORD)cKey) >> 4;
}

//==============================================================================
// FUNCTION: HashKey<LPCTSTR>
// PURPOSE:  String hash key generator.
template<> 
inline UINT HashKey<LPCTSTR> (const LPCTSTR& tsKey)
{
    LPCTSTR key = tsKey;
    if (key == NULL) return -1;

	UINT nHash = 0;
	while (*key)
		nHash = (nHash<<5) + nHash + *key++;
	return nHash;
}

//==============================================================================
// FUNCTION: HashKey<COpcString>
// PURPOSE:  String object hash key generator.
template<> 
inline UINT HashKey<COpcString> (const COpcString& cKey)
{
    LPCTSTR key = cKey;
    if (key == NULL) return -1;

	UINT nHash = 0;
	while (*key)
		nHash = (nHash<<5) + nHash + *key++;
	return nHash;
}

//==============================================================================
// TYPE:    COpcStringMap
// PURPOSE: A string to string map.

typedef COpcMap<COpcString,COpcString> COpcStringMap;

//#ifndef OPCUTILS_EXPORTS
template class OPCUTILS_API COpcMap<COpcString,COpcString>;
//#endif

#endif //ndef _COpcMap_H_

