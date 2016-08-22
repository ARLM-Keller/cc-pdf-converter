/**
	@file
	@brief Implementation for the CCRegistry class, used to various registry operations.
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

#include "precomp.h"
#include "CCRegistry.h"

/// Maximum length of string value supported
#define STRING_MAX	256

/// Last error's operation
CCRegistry::RegistryOperation CCRegistry::eLastErrorOp = None;
/// Last error code
DWORD CCRegistry::dwLastError = 0;

/**
	@param hFromKey the registry key to look in (can be any valid HKEY value)
	@param pPath path to the string, inside the key
	@param pName the string name
	@param[out] pSuccess If not NULL, will contain a success flag.
	@return The retrieved string, empty if failed. Note that an empty string will also be
	returned if the value is empty; in order to differenciate between those cases, use the pSuccess parameter.
*/
std::tstring CCRegistry::GetString(HKEY hFromKey, LPCTSTR pPath, LPCTSTR pName, bool* pSuccess /* = NULL */, DWORD dwExtraFlags /* = 0 */)
{
	HKEY hKey = NULL;
	// open the subkey
	if ((dwLastError = RegOpenKeyEx(hFromKey, pPath, 0, KEY_READ, &hKey)) != ERROR_SUCCESS)
	{
		eLastErrorOp = OpenKey;
		if (pSuccess != NULL)
			*pSuccess = false;
		return _T("");
	}

	// fetch value's length and type
	DWORD type;
	DWORD nSize = 0;
	if ((dwLastError = RegQueryValueEx(hKey, pName, NULL, &type, NULL, &nSize)) != ERROR_SUCCESS) {
		eLastErrorOp = QueryValueNull;
		RegCloseKey(hKey);
		if (pSuccess != NULL)
			*pSuccess = false;
		return _T("");
	}

	// make sure it's a string
	if (type != REG_SZ) {
		eLastErrorOp = WrongType;
		dwLastError = type;
		RegCloseKey(hKey);
		if (pSuccess != NULL)
			*pSuccess = false;
		return _T("");
	}

	if (nSize == 0)
	{
		eLastErrorOp = SizeZero;
		dwLastError = 0;
		RegCloseKey(hKey);
		if (pSuccess != NULL)
			*pSuccess = false;
		return _T("");
	}

	TCHAR* p = new TCHAR[(nSize / sizeof(TCHAR)) + 1];

	// get the value itself
	if ((dwLastError = RegQueryValueEx(hKey, pName, NULL, &type, (LPBYTE)p, &nSize)) != ERROR_SUCCESS) {
		eLastErrorOp = QueryValue;
		delete [] p;
		RegCloseKey(hKey);
		if (pSuccess != NULL)
			*pSuccess = false;
		return _T("");
	}

	// close key and cleanup
	std::tstring sRet = p;
	delete [] p;
	RegCloseKey(hKey);
	if (pSuccess != NULL)
		*pSuccess = true;
	return sRet;
}

/**
	@param hKey the registry key to look in (can be any valid HKEY value)
	@param pPath path to the string, inside the key
	@param pName the string name
	@param pValue the string to set
	@return true if successfully set, false otherwise
*/
bool CCRegistry::SetString(HKEY hKey, LPCTSTR pPath, LPCTSTR pName, LPCTSTR pValue)
{
	// Open the key (create it if necessary)
	HKEY hSubKey = NULL;
	if ((dwLastError = RegCreateKeyEx(hKey, pPath, 0, _T(""), REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hSubKey, NULL)) != ERROR_SUCCESS)
	{
		eLastErrorOp = CreateKey;
		return false;
	}

	// Set the value
	if ((dwLastError = RegSetValueEx(hSubKey, pName, 0, REG_SZ, (LPBYTE)pValue, (DWORD)(_tcslen(pValue) * sizeof(TCHAR)))) != ERROR_SUCCESS)
	{
		eLastErrorOp = SetValue;
		RegCloseKey(hSubKey);
		return false;
	}

	// Close the key
	RegCloseKey(hSubKey);
	return true;
}

/**
	@param hKey the registry key to look in (can be any valid HKEY value)
	@param pPath path to the string, inside the key
	@return true if successfully set, false otherwise
*/
bool CCRegistry::TestWrite(HKEY hKey, LPCTSTR pPath)
{
	// Open the key (create it if necessary)
	HKEY hSubKey = NULL;
	if ((dwLastError = RegCreateKeyEx(hKey, pPath, 0, _T(""), REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hSubKey, NULL)) != ERROR_SUCCESS)
	{
		eLastErrorOp = CreateKey;
		return false;
	}

	RegCloseKey(hSubKey);
	return true;
}


/**
	@param hFromKey the registry key to look in (can be any valid HKEY value)
	@param pPath path to the number, under the base key
	@param pName name of the number
	@param dwValue number value
	@return false on error, true on success
*/
bool CCRegistry::SetNumber(HKEY hFromKey, LPCTSTR pPath, LPCTSTR pName, DWORD dwValue)
{
	// Open the key (create if necessary)
	HKEY hKey = NULL;
	if ((dwLastError = RegCreateKeyEx(hFromKey, pPath, 0, _T(""), REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL)) != ERROR_SUCCESS)
	{
		eLastErrorOp = CreateKey;
		return false;
	}

	// Set the value
	if ((dwLastError = RegSetValueEx(hKey, pName, 0, REG_DWORD, (LPBYTE)&dwValue, sizeof(dwValue))) != ERROR_SUCCESS)
	{
		eLastErrorOp = SetValue;
		RegCloseKey(hKey);
		return false;
	}

	// Close the key
	RegCloseKey(hKey);
	return true;
}

/**
	@param hFromKey the registry key to look in (can be any valid HKEY value)
	@param pPath path to the number, under the base key
	@param pName name of the number
	@param[out] pSuccess If not NULL, will contain a success flag.
	@return The numeric value from the registry. Note that 0 will also be
	returned if the value is empty; in order to differenciate between those cases, use the pSuccess parameter.
*/
DWORD CCRegistry::GetNumber(HKEY hFromKey, LPCTSTR pPath, LPCTSTR pName, bool* pSuccess /* = NULL */)
{
	HKEY hKey = NULL;
	// Open the key (never create!)
	if ((dwLastError = RegOpenKeyEx(hFromKey, pPath, 0, KEY_READ, &hKey)) != ERROR_SUCCESS)
	{
		eLastErrorOp = OpenKey;
		if (pSuccess != NULL)
			*pSuccess = false;
		return -1;
	}

	// Make sure it's the right type
	DWORD type;
	DWORD nSize = 0;
	if (RegQueryValueEx(hKey, pName, NULL, &type, NULL, &nSize) != ERROR_SUCCESS) 
	{
		// Nope, error.
		RegCloseKey(hKey);
		if (pSuccess != NULL)
			*pSuccess = false;
		return -1;
	}

	// Make sure it's the right size
	DWORD dwRet;
	if ((type != REG_DWORD) || (nSize != sizeof(dwRet))) 
	{
		// Nope
		RegCloseKey(hKey);
		if (pSuccess != NULL)
			*pSuccess = false;
		return -1;
	}

	// Get the value
	if (RegQueryValueEx(hKey, pName, NULL, &type, (LPBYTE)&dwRet, &nSize) != ERROR_SUCCESS) 
	{
		// Problem:
		RegCloseKey(hKey);
		if (pSuccess != NULL)
			*pSuccess = false;
		return -1;
	}

	// Close up
	RegCloseKey(hKey);
	if (pSuccess != NULL)
		*pSuccess = true;
	return dwRet;
}

/**
	@param hKey the registry key to look in (can be any valid HKEY value)
	@param pPath path to the string, inside the key
	@param pName the string name
	@param pBuffer[out] reference to a variable that receives the pointer to a new byte array containing the binary data
	@param nBytes[out] number of bytes that were allocated and populated in pBuffer
	@return true if successful (in which case pBuffer should be delete []'ed), false otherwise
*/
bool CCRegistry::GetBinary(HKEY hFromKey, LPCTSTR pPath, LPCTSTR pName, unsigned char*& pBuffer, int& nBytes)
{
	HKEY hKey = NULL;
	// open the subkey
	if ((dwLastError = RegOpenKeyEx(hFromKey, pPath, 0, KEY_READ, &hKey)) != ERROR_SUCCESS)
	{
		eLastErrorOp = OpenKey;
		return false;
	}

	// fetch value's length and type
	DWORD type;
	DWORD nSize = 0;
	if ((dwLastError = RegQueryValueEx(hKey, pName, NULL, &type, NULL, &nSize)) != ERROR_SUCCESS)
	{
		eLastErrorOp = QueryValueNull;
		RegCloseKey(hKey);
		return false;
	}

	// make sure it's a binary
	if (type != REG_BINARY)
	{
		eLastErrorOp = WrongType;
		dwLastError = type;
		RegCloseKey(hKey);
		return false;
	}

	if (nSize == 0)
	{
		eLastErrorOp = SizeZero;
		dwLastError = 0;
		RegCloseKey(hKey);
		return false;
	}

	unsigned char* pNewBuf = new unsigned char[nSize];

	// get the value itself
	if ((dwLastError = RegQueryValueEx(hKey, pName, NULL, &type, pNewBuf, &nSize)) != ERROR_SUCCESS) 
	{
		eLastErrorOp = QueryValue;
		delete [] pNewBuf;
		RegCloseKey(hKey);
		return false;
	}

	// close key and cleanup
	RegCloseKey(hKey);

	pBuffer = pNewBuf;
	nBytes = nSize;
	return true;
}

/**
	@param hKey the registry key to look in (can be any valid HKEY value)
	@param pPath path to the string, inside the key
	@param pName the string name
	@param pBuffer pointer to the byte array to store
	@param nBytes number of bytes in pBuffer to store
	@return true if successful, false otherwise
*/
bool CCRegistry::SetBinary(HKEY hKey, LPCTSTR pPath, LPCTSTR pName, unsigned char* pBuffer, int nBytes)
{
	// Open the key (create it if necessary)
	HKEY hSubKey = NULL;
	if ((dwLastError = RegCreateKeyEx(hKey, pPath, 0, _T(""), REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hSubKey, NULL)) != ERROR_SUCCESS)
	{
		eLastErrorOp = CreateKey;
		return false;
	}

	// Set the value
	if ((dwLastError = RegSetValueEx(hSubKey, pName, 0, REG_BINARY, pBuffer, nBytes)) != ERROR_SUCCESS)
	{
		eLastErrorOp = SetValue;
		RegCloseKey(hKey);
		return false;
	}

	// Close the key
	RegCloseKey(hSubKey);
	return true;
}
