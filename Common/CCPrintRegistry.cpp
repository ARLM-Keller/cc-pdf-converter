/**
	@file
	@brief 
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
#include "CCPrintRegistry.h"

#include <winspool.h>
#include <tchar.h>

/**
	@param hPrinter Handle to the printer
	@param lpSetting Name of the key to get data from
	@param dwDefault Default value of data (used if the key isn't found)
	@return The value in the registry, or the default if not found (or is of the wrong type)
*/
DWORD CCPrintRegistry::GetRegistryDWORD(HANDLE hPrinter, LPCTSTR lpSetting, DWORD dwDefault)
{
	// Get the printer key:
	DWORD dwType, dwValue, dwSize = sizeof(DWORD);
	if (GetPrinterData(hPrinter, (LPTSTR)lpSetting, &dwType, (LPBYTE)&dwValue, sizeof(DWORD), &dwSize) != ERROR_SUCCESS)
		return dwDefault;

	if (dwType != REG_DWORD)
		return dwDefault;

	return dwValue;
}

/**
	@param hPrinter Handle to the printer
	@param lpSetting Name of the key to write data to
	@param dwValue The value to write
	@return true if written successfully, false if failed
*/
bool CCPrintRegistry::SetRegistryDWORD(HANDLE hPrinter, LPCTSTR lpSetting, DWORD dwValue)
{
	// Get the printer key:
	HRESULT hr = SetPrinterData(hPrinter, (LPTSTR)lpSetting, REG_DWORD, (LPBYTE)&dwValue, sizeof(dwValue));
	return hr == ERROR_SUCCESS;
}

/**
	@param hPrinter Handle to the printer
	@param lpSetting Name of the key to get data from
	@param lpDefault Default value of data (used if the key isn't found)
	@return The value in the registry, or the default if not found (or is of the wrong type)
*/
std::tstring CCPrintRegistry::GetRegistryString(HANDLE hPrinter, LPCTSTR lpSetting, LPCTSTR lpDefault)
{
	// Get the printer key:
	DWORD dwType, dwSize;
	LONG lRes = GetPrinterData(hPrinter, (LPTSTR)lpSetting, &dwType, NULL, 0, &dwSize);
	if (((lRes != ERROR_MORE_DATA) && (lRes != ERROR_SUCCESS)) || (dwType != REG_SZ))
		return lpDefault;

	if (dwSize == 0)
		return _T("");

	std::tstring sRet(lpDefault);
	TCHAR* lpData = new TCHAR[(dwSize / sizeof(TCHAR)) + 1];
	if (GetPrinterData(hPrinter, (LPTSTR)lpSetting, NULL, (LPBYTE)lpData, dwSize, &dwSize) == ERROR_SUCCESS)
		sRet = lpData;
	delete [] lpData;

	return sRet;
}

/**
	@param hPrinter Handle to the printer
	@param lpSetting Name of the key to write data to
	@param sValue The value to write
	@return true if written successfully, false if failed
*/
bool CCPrintRegistry::SetRegistryString(HANDLE hPrinter, LPCTSTR lpSetting, const std::tstring& sValue)
{
	// Set the value
	return SetPrinterData(hPrinter, (LPTSTR)lpSetting, REG_SZ, (LPBYTE)sValue.c_str(), (WORD) ((sValue.size() + 1) * sizeof(TCHAR))) == ERROR_SUCCESS;
}

/**
	@param hPrinter Handle to the printer
	@param lpSetting Name of the key to remove
	@return true if removed successfully, false if failed
*/
bool CCPrintRegistry::EraseRegistryValue(HANDLE hPrinter, LPCTSTR lpSetting)
{
	return (DeletePrinterData(hPrinter, (LPTSTR)lpSetting) == ERROR_SUCCESS);
}

/**
	@param hPrinter Handle to the printer
	@param lValues[out] List to fill with the value names
	@param lpPrefix Value name prefix; if not NULL, will only return the value names that start with the prefix
	@return true if all went well, false if something failed
*/
bool CCPrintRegistry::EnumRegistryValues(HANDLE hPrinter, STRLIST& lValues, LPCTSTR lpPrefix /* = NULL */)
{
	// First get the size of buffer we'll need:
	DWORD dwValueSize = 0, dwIndex = 0, dw, dw2, dwRet;
	if (EnumPrinterData(hPrinter, 0, NULL, 0, &dwValueSize, NULL, NULL, 0, &dw) != ERROR_SUCCESS)
		return false;

	// No values, all's well
	if (dwValueSize == 0)
		return true;

	// Create buffer
	TCHAR* pValueName = (TCHAR*)new char[dwValueSize];
	do
	{
		// Get next value
		dwRet = EnumPrinterData(hPrinter, dwIndex, pValueName, dwValueSize, &dw, NULL, NULL, 0, &dw2);
		if ((dwRet != ERROR_MORE_DATA) && (dwRet != ERROR_SUCCESS))
			break;
		dwIndex++;
		if (_tcslen(pValueName) == 0)
			// No value name, continue
			continue;
		if (lpPrefix != NULL)
		{
			// Check prefix
			if (_tcsncmp(lpPrefix, pValueName, _tcslen(lpPrefix)) != 0)
				// Not the same, move on
				continue;
		}
		// OK, add value name
		lValues.push_back(pValueName);
	} while (true);

	delete [] pValueName;

	return true;
}
