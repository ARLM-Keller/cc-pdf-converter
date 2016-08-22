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

#ifndef _CCPRINTREGISTRY_H_
#define _CCPRINTREGISTRY_H_

#include "CCTChar.h"
#include <list>

/// List of strings
typedef std::list<std::tstring> STRLIST;

/**
    @brief Functions to handle reading/writing to the printer registry settings
*/
namespace CCPrintRegistry
{

/// Reads DWORD data from the printer's registry settings
DWORD		GetRegistryDWORD(HANDLE hPrinter, LPCTSTR lpSetting, DWORD dwDefault);
/**
	@brief Reads interger data from the printer's registry settings
	@param hPrinter Handle to the printer
	@param lpSetting Name of the key to read from
	@param nDefault Default value of the integer if not found in the registry
	@return The value of the data in the registry, or the default value if not found there
*/
inline int	GetRegistryInt(HANDLE hPrinter, LPCTSTR lpSetting, int nDefault) {return (int)GetRegistryDWORD(hPrinter, lpSetting, (DWORD)nDefault);}
/// Reads string data from the printer's registry settings
std::tstring GetRegistryString(HANDLE hPrinter, LPCTSTR lpSetting, LPCTSTR lpDefault);
/**
	@brief Reads boolean data from the printer's registry settings
	@param hPrinter Handle to the printer
	@param lpSetting Name of the key to read from
	@param bDefault Default value of the flag if not found in the registry
	@return The value of the data in the registry, or the default value if not found there
*/
inline bool	GetRegistryBool(HANDLE hPrinter, LPCTSTR lpSetting, bool bDefault) {return (GetRegistryDWORD(hPrinter, lpSetting, bDefault ? 1 : 0) == 1);};

/// Writes DWORD data to the printer's registry settings
bool		SetRegistryDWORD(HANDLE hPrinter, LPCTSTR lpSetting, DWORD dwValue);
/**
	@brief Writes integer data to the printer's registry settings
	@param hPrinter Handle to the printer
	@param lpSetting Name of the key to write to
	@param nValue Value to set
*/
inline bool	SetRegistryInt(HANDLE hPrinter, LPCTSTR lpSetting, int nValue) {return SetRegistryDWORD(hPrinter, lpSetting, (DWORD)nValue);};
/**
	@brief Writes boolean data to the printer's registry settings
	@param hPrinter Handle to the printer
	@param lpSetting Name of the key to write to
	@param bValue Value to set
*/
inline bool	SetRegistryBool(HANDLE hPrinter, LPCTSTR lpSetting, bool bValue) {return SetRegistryDWORD(hPrinter, lpSetting, bValue ? 1 : 0);};
/// Writes string data to the printer's registry settings
bool		SetRegistryString(HANDLE hPrinter, LPCTSTR lpSetting, const std::tstring& sValue);

/// Deletes a data value from the printer's registry settings
bool		EraseRegistryValue(HANDLE hPrinter, LPCTSTR lpSetting);

/// Get list of value names of a key (optionally, only those that have a specific prefix)
bool		EnumRegistryValues(HANDLE hPrinter, STRLIST& lValues, LPCTSTR lpPrefix = NULL);

};

#endif   //#define _CCPRINTREGISTRY_H_
