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

#ifndef __OPENARRAY_H_
#define __OPENARRAY_H_

//DOM-IGNORE-BEGIN

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

template <class T>
class OpenArray {

   private:
      long     size;         // highest used index          
      long     totElem;      // number of non NULL elements
      long     allOPCize;    // allocated memory (number of elements, not bytes)
      T       *array;        // the array of elements of class T

   public:
      //!temp!//const static int E_OK;
      //!temp!//const static int E_NOTENOUGHMEMORY;
      //!temp!//const static int E_WRONGINDEX;

  
         ///////////////////////////////////////////////////////////////
         //   Constructor
         ///////////////////////////////////////////////////////////////
     OpenArray() {
         array       = NULL;
         allOPCize   = 0;
         size        = 0;
         totElem     = 0;
      }

         ///////////////////////////////////////////////////////////////
         //   Desstructor
         ///////////////////////////////////////////////////////////////
      ~OpenArray() {
         if (array != NULL) {
            delete [] array;
			array = NULL;
         }

      }



         ///////////////////////////////////////////////////////////////
         // Get the highest allocated element index
         ///////////////////////////////////////////////////////////////
      long Size() {
         return size;
      }



         ///////////////////////////////////////////////////////////////
         //  Returns the 1st free index.
         //  If new element is free then size+1 is returned.
         ///////////////////////////////////////////////////////////////
      long New( void )
      {
            long i;

         for( i=1 ; i < allOPCize ; ++i ) {        // search from 1 because element 0 is never used.
            if( array[i] == NULL ) {               // is free
               return i;
            }
         }
         return allOPCize;
      }





         ///////////////////////////////////////////////////////////////
         //  Appends the given element
         ///////////////////////////////////////////////////////////////
      int AppendElem( T elem )
      {
         return PutElem( New(), elem );
      }





         ///////////////////////////////////////////////////////////////
         //  Get the element at the given index
         ///////////////////////////////////////////////////////////////
      int GetElem( long idx, T *elem ) {
         if( idx < 0 ) {
            *elem = NULL;                          
            return E_INVALIDARG;
         }
         if( idx >= size ) {
            *elem = NULL;
            return E_FAIL;   
         }
         *elem = array[idx];
         if( *elem )          return S_OK;         
         else                 return E_FAIL;       // NULL element is error
      }



         ///////////////////////////////////////////////////////////////
         //  Store the given element at the given index.
         //  NULL as the element frees the specified index.
         //  If the given element index is above the allocated array
         //  sie, then the array size is increased.
         ///////////////////////////////////////////////////////////////
      int PutElem( long idx,     // element index
                   T    elem ) { // NULL deletes the entry
      
            long newallOPCize;
            T *newarray;
            long i;
      
         if( idx < 0 ) {                           // illegal index
            return E_INVALIDARG;
         }
         if( idx >= allOPCize ) {                  // above current size
      
            if (allOPCize == 0) {                  // first allocation
               newallOPCize = 4;
            } else {
               newallOPCize = allOPCize*2;         // standard enlargement
            }
      
                  // Heuristic: new size = new allOPCize * 2
            while( idx >= newallOPCize ) {           
               newallOPCize *= 2;
            }
      
               // Must allocate a new array filling the undefined elements with NULLs
            newarray = new T[newallOPCize];
            if( newarray == NULL ) {
               return E_OUTOFMEMORY;               // error
            }
      
            for( i = allOPCize ; i < newallOPCize ; i++ ) {
               newarray[i] = NULL;                 // zero the allocated array
            }
            if( allOPCize > 0 ) {                  // copy old elements to new array
                                                   // this could be done with memcpy
               for( i = 0 ; i < allOPCize ; i++ ) {
                  newarray[i] = array[i];
               }
               delete [] array;                    // free old array
            }
            array = newarray;                      // switch to the new array
            allOPCize = newallOPCize;
         }
                                     // now handle the element
         if( array[idx] != NULL ) {               // deleting a element
            totElem --;
         }

         array[idx] = elem; 

         if( elem != NULL ) {                     // inserting a element
            totElem ++;
         }
         if( idx >= size ) {                      // adapt highest used index
            size = idx+1;
         }

         return S_OK;
      }



         ///////////////////////////////////////////////////////////////
         //  return the number of defined elements
         ///////////////////////////////////////////////////////////////
      long TotElem()
      {
         return totElem;
      }


         ///////////////////////////////////////////////////////////////
         //  Returns the index of the 1st used item
         ///////////////////////////////////////////////////////////////
      int First( long *idx )
      {
            long i;

         for( i=0 ; i < size ; ++i ) {
            if( array[i] != NULL ) {         // in use
               *idx = i;
               return S_OK;
            }
         }
         *idx = 0;
         return E_FAIL;
      }



         ///////////////////////////////////////////////////////////////
         //  Returns the index of the next used item
         ///////////////////////////////////////////////////////////////
      int Next( long idxFrom, long *idxNext )
      {
            long i;

         i = idxFrom + 1;
         while ( i < size ) {
            if ( array[i] != NULL ) {
               *idxNext = i;
               return S_OK;
            }
            i++;
         }

         *idxNext = 0;
         return E_FAIL;
      }

};
//DOM-IGNORE-END


#endif // __OPENARRAY_H_

