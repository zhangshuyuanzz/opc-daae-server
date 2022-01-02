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

//DOM-IGNORE-BEGIN

//-----------------------------------------------------------------------
// INLCUDE
//-----------------------------------------------------------------------
#include "stdafx.h"

#include <comdef.h>                       // for _variant_t
#include <math.h>                         // for fabs()

#include "DaDeviceItem.h"
#include "variantcompare.h"

//=========================================================================
// Compare operator for VARIANTs.
//
// Since not all possible VARIANT types are supported by class _variant_t
// we use an own impelemtation.
//=========================================================================
bool operator == ( const VARIANT vSrc, const VARIANT vOp )
{
	if (&vOp == &vSrc) {
		return true;
	}

	//
	// Variants not equal if types don't match
	//
	if (V_VT(&vOp) != V_VT(&vSrc)) {
		return false;
	}

	//
	// Check type specific values
	//
	switch (V_VT(&vOp)) {

	  case VT_EMPTY:
	  case VT_NULL:
		  return true;

	  case VT_I2:
		  return V_I2(&vOp) == V_I2(&vSrc);

	  case VT_I4:
		  return V_I4(&vOp) == V_I4(&vSrc);

	  case VT_R4:
		  return V_R4(&vOp) == V_R4(&vSrc);

	  case VT_R8:
		  return V_R8(&vOp) == V_R8(&vSrc);

	  case VT_CY:
		  return memcmp(&(V_CY(&vOp)), &(V_CY(&vSrc)), sizeof(CY)) == 0;

	  case VT_DATE:
		  return V_DATE(&vOp) == V_DATE(&vSrc);

	  case VT_BSTR:
		  return (::SysStringByteLen(V_BSTR(&vOp)) == ::SysStringByteLen(V_BSTR(&vSrc))) &&
			  (memcmp(V_BSTR(&vOp), V_BSTR(&vSrc), ::SysStringByteLen(V_BSTR(&vOp))) == 0);

	  case VT_DISPATCH:
		  return V_DISPATCH(&vOp) == V_DISPATCH(&vSrc);

	  case VT_ERROR:
		  return V_ERROR(&vOp) == V_ERROR(&vSrc);

	  case VT_BOOL:
		  return V_BOOL(&vOp) == V_BOOL(&vSrc);

	  case VT_UNKNOWN:
		  return V_UNKNOWN(&vOp) == V_UNKNOWN(&vSrc);

	  case VT_DECIMAL:
		  return memcmp(&(V_DECIMAL(&vOp)), &(V_DECIMAL(&vSrc)), sizeof(DECIMAL)) == 0;

	  case VT_UI1:
		  return V_UI1(&vOp) == V_UI1(&vSrc);

		  //
		  // Now the values not supported by _variant_t.
		  //
	  case VT_I1:
		  return V_I1(&vOp) == V_I1(&vSrc);

	  case VT_I8:
		  return V_I8(&vOp) == V_I8(&vSrc);

	  case VT_UI2:
		  return V_UI2(&vOp) == V_UI2(&vSrc);

	  case VT_UI4:
		  return V_UI4(&vOp) == V_UI4(&vSrc);

	  case VT_UI8:
		  return V_UI8(&vOp) == V_UI8(&vSrc);

	  case VT_INT:
		  return V_INT(&vOp) == V_INT(&vSrc);

	  case VT_UINT:
		  return V_UINT(&vOp) == V_UINT(&vSrc);

	  default:
		  _com_issue_error(E_INVALIDARG);
		  // fall through
	}
	return false;

} // operator ==


//=========================================================================
// Operator '-' for class _variant_t
// Note :   only the restricted types of items with Analog EU Info 
//          are supported.
//=========================================================================
_variant_t operator-( const VARIANT Minuend, const VARIANT Subtrahend )
{
	//
	// Variants must have the same types
	//
	if (V_VT( &Minuend ) != V_VT( &Subtrahend )) {
		_com_issue_error( E_INVALIDARG );
	}
	//
	_variant_t varRes( Minuend );       // Result Variant
	//
	// Check type specific values
	//
	switch (V_VT( &varRes )) {
	  case VT_I2:    V_I2( &varRes )   -= V_I2( &Subtrahend );    break;
	  case VT_I4:    V_I4( &varRes )   -= V_I4( &Subtrahend );    break;
	  case VT_R4:    V_R4( &varRes )   -= V_R4( &Subtrahend );    break;
	  case VT_R8:    V_R8( &varRes )   -= V_R8( &Subtrahend );    break;
	  case VT_BOOL:  V_BOOL( &varRes ) -= V_BOOL( &Subtrahend );  break;
	  case VT_UI1:   V_UI1( &varRes )  -= V_UI1( &Subtrahend );   break;
	  default:                         // fall through
		  _com_issue_error( E_INVALIDARG ); 
	}
	return varRes;

} // operator -




//=================================================================================
// CompareSimpleVariant
// --------------------
//    Compares two simple variants and checks if the values are identical.
//    Also Percent Deadband range check is supported.
//=================================================================================
static HRESULT CompareSimpleVariant( DaDeviceItem& DItem,
									float fltPercentDeadband,
									const VARIANT& varLast, const VARIANT& varNew, BOOL& fItemValueChanged )
{
	_ASSERTE( ! V_ISARRAY( &varNew ) );    // only simple types
	_ASSERTE( ! V_ISARRAY( &varLast ) );


	fItemValueChanged = FALSE;

	OPCEUTYPE   EUType;
	_variant_t  vEUInfo;

	HRESULT hres = DItem.get_EUData( &EUType, &vEUInfo );
	if (FAILED( hres )) {
		return hres;
	}

	try {

		if (EUType == OPC_ANALOG) {
			// Only if EU is Analog
			if (V_VT( &varLast ) == VT_EMPTY) {
				// Data Type of Last Read Value is
				fItemValueChanged = TRUE;        // VT_EMPTY ==> It is the first update cycle.
				return S_OK;
		 }
			else if (V_VT( &varNew ) == VT_BOOL) {
				if (!(varNew == varLast)) {      
					fItemValueChanged = TRUE;     // No deadband range for boolean types
				}
				return S_OK;
		 }
			else if (fltPercentDeadband == (float)100.0) {
				fItemValueChanged = FALSE;       // 100% deadband means do not report value changes
				return S_OK;
		 }

			double dAnalogEURange;
			hres = DItem.get_AnalogEURange( &dAnalogEURange );
			if (FAILED( hres )) {
				return hres;
		 }

			double      dDeadbandRange = (fltPercentDeadband/100) * ( dAnalogEURange );
			_variant_t  varDiff;
			// Note :
			//    varNew and varLast are values in the requested data type format.
			//    If the data type does not belong to the restricted canonical data types
			//    of items with Analog EU Info then the value difference must be calculated
			//    with values in the canonical data type format.

			switch (V_VT( &varNew )) {
			case VT_I2:                      // values already in a data type format which is valid
			case VT_I4:                      // for Items with Analog EU Info. No conversion is required.
			case VT_R4:
			case VT_R8:
				varDiff = varLast - varNew;
				break;

			case VT_UI1:                     // usnigned value, only the absolute difference is required
				varDiff = (V_UI1( &varNew ) > V_UI1( &varLast )) ? varNew - varLast :
					varLast - varNew;
				break;

			default:                         // Data format requested by the client (e.g. BSTR) cannot be used 
				// to calculate the difference.
				// Make a copy and calculate difference with values in the
				// canonical data type format.
				{
					_variant_t  varNewCanonical;
					_variant_t  varLastCanonical;

					VARTYPE vtCanonical = DItem.get_CanonicalDataType();

					varNewCanonical.ChangeType( vtCanonical, (_variant_t*)&varNew );
					varLastCanonical.ChangeType( vtCanonical, (_variant_t*)&varLast );

					switch (vtCanonical) {
			case VT_BOOL:           // No deadband range for boolean types
				if (!((VARIANT)varNewCanonical == varLastCanonical)) {
					fItemValueChanged = TRUE;
				}
				return S_OK;

			case VT_UI1:            // usnigned value, only the absolute difference is required
				varDiff = (V_UI1( &varNewCanonical ) > V_UI1( &varLastCanonical )) ?
					varNewCanonical - varLastCanonical :
				varLastCanonical - varNewCanonical;
				break;

			default :
				varDiff = varLastCanonical - varNewCanonical;
				break;
					}
				}
				break;
		 }

			varDiff.ChangeType( VT_R8 );        // convert to double, ChangeType() is better than
			// typecast (double) because makes no copy.
			if (fabs( (double)varDiff ) > dDeadbandRange) {
				fItemValueChanged = TRUE;
		 }
		} // EU Type is Analog
		else
		{
			switch (varNew.vt)
			{
			case VT_EMPTY: fItemValueChanged = (varNew.vt != varLast.vt); 
				break;
			case VT_I1:    fItemValueChanged = (varNew.cVal != varLast.cVal);
				break;
			case VT_UI1:   fItemValueChanged = (varNew.bVal != varLast.bVal);
				break;
			case VT_I2:    fItemValueChanged = (varNew.iVal != varLast.iVal);
				break;
			case VT_UI2:   fItemValueChanged = (varNew.uiVal != varLast.uiVal);
				break;
			case VT_I4:    fItemValueChanged = (varNew.lVal != varLast.lVal);
				break;
			case VT_UI4:   fItemValueChanged = (varNew.ulVal != varLast.ulVal);
				break;
			case VT_I8:    fItemValueChanged = (varNew.llVal != varLast.llVal);
				break;
			case VT_UI8:   fItemValueChanged = (varNew.ullVal != varLast.ullVal);
				break;
			case VT_R4:    fItemValueChanged = (varNew.fltVal != varLast.fltVal);
				break;
			case VT_R8:    fItemValueChanged = (varNew.dblVal != varLast.dblVal);
				break;
			case VT_CY:    fItemValueChanged = (varNew.cyVal.int64 != varLast.cyVal.int64);
				break;
			case VT_DATE:  fItemValueChanged = (varNew.date != varLast.date);
				break;
			case VT_BOOL:  fItemValueChanged = (varNew.boolVal != varLast.boolVal);   
				break;	     
			case VT_BSTR:  
				{
					if (varNew.bstrVal != NULL && varLast.bstrVal != NULL)
					{
						fItemValueChanged = (wcscmp(varNew.bstrVal, varLast.bstrVal) != 0);
					}
					else
					{
						fItemValueChanged = (varNew.bstrVal != varLast.bstrVal);
					}
				}
			}			
		}
	}
	catch (_com_error &e) {                   // _variant_t operators '-' and '==' can throw exceptions
		return e.Error();
	}
	return S_OK;

} // CompareSimpleVariant


//=================================================================================
// CompareVariant
// --------------
//    Compares two variants (incl. SAFEARRAYs) and checks if the values
//    are identical. Also Percent Deadband range check is supported.
//=================================================================================
HRESULT CompareVariant( DaDeviceItem& DItem, float fltPercentDeadband,
					   const VARIANT& varLast, const VARIANT& varNew, BOOL& fItemValueChanged )
{
	if (V_VT( &varNew ) != V_VT( &varLast )) {
		fItemValueChanged = TRUE;                 // Not identical if type changed.
		return S_OK;
	}

	fItemValueChanged = FALSE;

	if (V_ISARRAY( &varNew )) {
		// 
		// Item Value is an Array
		//

		// First check if the size of arrays is identical
		long     lLowerBoundLast, lUpperBoundLast, lLowerBoundNew, lUpperBoundNew, i;
		HRESULT  hres;

		hres = SafeArrayGetLBound( V_ARRAY( &varLast ), 1, &lLowerBoundLast );
		if (FAILED( hres )) {
			return hres;
		}
		hres = SafeArrayGetUBound( V_ARRAY( &varLast ), 1, &lUpperBoundLast );
		if (FAILED( hres )) {
			return hres;
		}
		hres = SafeArrayGetLBound( V_ARRAY( &varNew ), 1, &lLowerBoundNew );
		if (FAILED( hres )) {
			return hres;
		}
		hres = SafeArrayGetUBound( V_ARRAY( &varNew ), 1, &lUpperBoundNew );
		if (FAILED( hres )) {
			return hres;
		}

		if ((lLowerBoundNew != lLowerBoundLast) || (lUpperBoundNew != lUpperBoundLast)) {
			fItemValueChanged = TRUE;              // Array size has changed
			return S_OK;
		}

		SAFEARRAY & saNew  = *V_ARRAY( &varNew  );
		SAFEARRAY & saLast = *V_ARRAY( &varLast );

		// lock arrays for direct manipulations
		hres = SafeArrayLock( &saNew );
		if (FAILED( hres )) {
			return hres;
		}
		hres = SafeArrayLock( &saLast );
		if (FAILED( hres )) {
			SafeArrayUnlock( &saNew );
			return hres;
		}
		// Both arrays has the same size and identical data types.
		// Check if values of the array elements has changed.
		//
		// Cannot compare entire data block because elements of VARIANT, BSTR, CY
		// and DECIMAL arrays points to values.
		if ((V_VT( &varNew ) == (VT_ARRAY | VT_VARIANT))   ||
			(V_VT( &varNew ) == (VT_ARRAY | VT_BSTR))      ||
			(V_VT( &varNew ) == (VT_ARRAY | VT_CY))        || 
			(V_VT( &varNew ) == (VT_ARRAY | VT_DECIMAL)) ) {
				// For better performance we don't extract single elements.
				// Instead all elements are compared directly.

				switch (V_VT( &varNew ) & ~VT_ARRAY) {

			case VT_VARIANT   :                 // compare single VARIANTs
				// ------------------------------------------------------------------
				for (i=lLowerBoundNew; i<=lUpperBoundNew; i++) {
					hres = CompareVariant( DItem, fltPercentDeadband, ((VARIANT *)saLast.pvData)[i], ((VARIANT *)saNew.pvData)[i], fItemValueChanged );
					if (fItemValueChanged || FAILED( hres )) {
						break;
					}
				}
				break;


			case VT_BSTR      :                 // compare single BSTRs
				// ------------------------------------------------------------------
				for (i=lLowerBoundNew; i<=lUpperBoundNew; i++) {
					if ((SysStringByteLen( ((BSTR *)saNew.pvData)[i] ) != SysStringByteLen( ((BSTR *)saLast.pvData)[i] )) ||
						(memcmp( ((BSTR *)saNew.pvData)[i], ((BSTR *)saLast.pvData)[i], SysStringByteLen( ((BSTR *)saNew.pvData)[i] )) != 0)) {

							fItemValueChanged = TRUE;
							break;
					}
				}
				break;

			case VT_CY        :                  // compare single CURRENCYs
				// ------------------------------------------------------------------
				for (i=lLowerBoundNew; i<=lUpperBoundNew; i++) {
					if (memcmp(&((CY *)saLast.pvData)[i], &((CY *)saNew.pvData)[i], sizeof (CY)) != 0) {
						fItemValueChanged = TRUE;
						break;
					}
				}
				break;

			case VT_DECIMAL   :                  // compare single DECIMALs
				// ------------------------------------------------------------------
				for (i=lLowerBoundNew; i<=lUpperBoundNew; i++) {
					if (memcmp(&((DECIMAL *)saLast.pvData)[i], &((DECIMAL *)saNew.pvData)[i], sizeof (DECIMAL)) != 0) {
						fItemValueChanged = TRUE;
						break;
					}
				}
				break;

				} // switch complex array types

		}
		else {
			// Array of simple types.
			// The entire memory block can be compared direcly if there are no elemnts with EU info.
			BOOL  fCompareEntireMemBlock;

			switch (V_VT( &varNew  ) & ~VT_ARRAY) {
			   case VT_I2:
			   case VT_UI2:
			   case VT_I4:
			   case VT_UI4:
			   case VT_I8:
			   case VT_UI8:
			   case VT_R4:
			   case VT_R8:
			   case VT_BOOL:                    // valid VARTYPEs for Items with Analog EU Info
			   case VT_UI1:      fCompareEntireMemBlock = FALSE;
				   break;

			   default:          fCompareEntireMemBlock = TRUE;
				   break;
			}

			if (fCompareEntireMemBlock) {

				// Compare memory
				fItemValueChanged = (memcmp(  saNew.pvData, saLast.pvData,
					saNew.cbElements * saNew.rgsabound[0].cElements ) != 0);

		 }
			else {
				// Extract single elements and compare.
				for (i=lLowerBoundNew; i<=lUpperBoundNew; i++) {

					_variant_t  varNewElem;          // Variants with element values to compare
					_variant_t  varLastElem;

					// Set the type of the element
					V_VT( &varNewElem )  = V_VT( &varNew  ) & ~VT_ARRAY;
					V_VT( &varLastElem ) = V_VT( &varLast ) & ~VT_ARRAY;

					// Set the value of the element (get from array)
					hres = SafeArrayGetElement( V_ARRAY( &varLast ), &i, &V_UI1( &varLastElem ) );
					if (FAILED( hres )) {
						break;
					}
					hres = SafeArrayGetElement( V_ARRAY( &varNew ), &i, &V_UI1( &varNewElem ) );
					if (FAILED( hres )) {
						break;
					}
					// Compare the extracted variant value.
					hres = CompareVariant( DItem, fltPercentDeadband, varLastElem, varNewElem, fItemValueChanged );
					if (fItemValueChanged || FAILED( hres )) {
						break;
					}
				}
		 }

		} // array of simple types

		// unlock arrays
		HRESULT hresTmp = SafeArrayUnlock( &saNew );
		if (FAILED( hresTmp )) {
			SafeArrayUnlock( &saLast );
		}
		else {
			hresTmp = SafeArrayUnlock( &saLast );
		}

		if (FAILED( hres )) {
			return hres;                           // Keep error code from compare if failed.   
		}
		return hresTmp;

	}  // Data Type is an Array
	else {
		// 
		// Item Value is not an ARRAY
		//
		return CompareSimpleVariant( DItem, fltPercentDeadband, varLast, varNew, fItemValueChanged );
	}

	return S_OK;

} // CompareVariant     

//DOM-IGNORE-END
