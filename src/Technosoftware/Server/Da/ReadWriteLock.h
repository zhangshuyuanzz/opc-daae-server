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

#ifndef __READWRITELOCK_H_
#define __READWRITELOCK_H_

/**
 * @class	Provides a class implementing a read / write lock mechanism
 *
 * @brief	DOM-IGNORE-BEGIN.
 */

class ReadWriteLock
{
   private:
      HANDLE readMutex_;
      HANDLE writeMutex_;
      HANDLE readerEvent_;
      long   counter_;

   public:

      /**
       * @fn	ReadWriteLock::ReadWriteLock( void );
       *
       * @brief	Creates the Monitor Object.
       */

      ReadWriteLock( void );

      /**
       * @fn	ReadWriteLock::~ReadWriteLock();
       *
       * @brief	Free all resources.
       */

      ~ReadWriteLock();

      /**
       * @fn	BOOL ReadWriteLock::Initialize();
       *
       * @brief	Initialize the Monitor Object.
       *
       * @return	true if it succeeds, false if it fails.
       */

      BOOL Initialize();

      /**
       * @fn	void ReadWriteLock::BeginReading();
       *
       * @brief	Reader wants to enter.
       */

      void BeginReading();

      /**
       * @fn	void ReadWriteLock::EndReading();
       *
       * @brief	Reader leaves critical section.
       */

      void EndReading();

      /**
       * @fn	void ReadWriteLock::BeginWriting();
       *
       * @brief	Writer wants to enter.
       */

      void BeginWriting();

      /**
       * @fn	void ReadWriteLock::EndWriting();
       *
       * @brief	Writer leaves critical section.
       */

      void EndWriting();
};
//DOM-IGNORE-END

#endif // if not defines __READWRITELOCK_H_
