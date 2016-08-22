/**
	@file
	@brief Interface for the CCRegistry class, used to various registry operations.
*/

/*
 * CC PDF Converter: Windows PDF Printer with Creative Commons license support
 * Excel to PDF Converter: Excel PDF printing addin, keeping hyperlinks AND Creative Commons license support
 * Copyright (C) 2007-2010 Guy Hachlili <hguy@cogniview.com>, Cogniview LTD.
 * 
 * This file is part of CC PDF Converter / Excel to PDF Converter
 * 
 * CC PDF Converter and Excel to PDF Converter are free software;
 * you can redistribute them and/or modify them under the terms of the 
 * GNU General Public License as published by the Free Software Foundation;
 * either version 2 of the License, or (at your option) any later version.
 * 
 * CC PDF Converter and Excel to PDF Converter are is distributed in the hope 
 * that they will be useful, but WITHOUT ANY WARRANTY; without even the implied 
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>. * 
 */

#if !defined(__CCREGISTRY_H__)
#define __CCREGISTRY_H__

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "CCTChar.h"
#include <tchar.h>

/**
    @brief Class used for various registry operations, under a single pre-defined registry key.
	
	CCRegistry supports all kinds of registry operations, including basic ones like getting and setting of values
	(both string and numeric), and specific ones like retrieving C-Center plugin DLL paths. Additionally, it
	has a different registry base path depending on the product and debug mode (Cogniview vs. CogniviewD vs. CatIndex).

	It should have been a namespace and not a class as it's a fully static class, has no members and all its
	functions are static.
*/
class CCRegistry
{
public:
	// Definitions
	/// Possible registry error types
	enum RegistryOperation
	{
		OpenKey,
		QueryValueNull,
		QueryValue,
		WrongType,
		SizeZero,
		CreateKey,
		SetValue,
		None
	};


	// Generic registry operations
	/// Retrieves a string registry value from anywhere in the registry.
	static std::tstring	GetString(HKEY hKey, LPCTSTR pPath, LPCTSTR pName, bool* pSuccess = NULL, DWORD dwExtraFlags = 0);
	/// Sets a string value anywhere in the registry.
	static bool			SetString(HKEY hKey, LPCTSTR pPath, LPCTSTR pName, LPCTSTR pValue);

	/// Retrieves a binary registry value from anywhere in the registry.
	static bool			GetBinary(HKEY hKey, LPCTSTR pPath, LPCTSTR pName, unsigned char*& pBuffer, int& nBytes);
	/// Sets a binary value anywhere in the registry.
	static bool			SetBinary(HKEY hKey, LPCTSTR pPath, LPCTSTR pName, unsigned char* pBuffer, int nBytes);

	/// Retrieves a numeric value from the registry
	static DWORD		GetNumber(HKEY hKey, LPCTSTR pPath, LPCTSTR pName, bool* pSuccess = NULL);
	/// Sets a numeric in the registry
	static bool			SetNumber(HKEY hKey, LPCTSTR pPath, LPCTSTR pName, DWORD dwValue);

	
	/// Tests if writing a value in the registry is allowed
	static bool			TestWrite(HKEY hKey, LPCTSTR pPath);

	/**
		@brief Returns the last registry error type
		@return The last error type
	*/
	static RegistryOperation GetLastErrorOperation() {return eLastErrorOp;};
	/**
		@brief Returns the last system error code for the registry operation
		@return The last error code
	*/
	static DWORD		GetLastError() {return dwLastError;};

protected:
	/// Last error type
	static enum RegistryOperation eLastErrorOp;
	/// Last error code
	static DWORD dwLastError;
};

#endif // !defined(__CCREGISTRY_H__)
