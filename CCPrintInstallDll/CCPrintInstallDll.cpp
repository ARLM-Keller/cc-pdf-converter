/**
	@file
	@brief Entry and exported functions for the CCPrintInstallDll module
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

#include "stdafx.h"
#include "CCPrintInstallFunctions.h"

/// Global handle for the module
HANDLE ghModule = NULL;

/**
	@brief DLL initialization/destruction function
	@param hModule Handle to the module
	@param ul_reason_for_call Reason this function is called
	@param lpReserved Ignored
	@return TRUE
*/
BOOL APIENTRY DllMain(HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved)
{
	switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
			// Initial module load for this process; keep the handle
			ghModule = hModule;
			break;
	}
    return TRUE;
}

/**
	@brief Call to install the CC PDF printer driver
	@param hwnd Handle to the parent window
	@param lpIValue Ignored
	@param lpszValue Ignored
	@return 0 if successfully installed, 1 if failed
*/
LONG APIENTRY InstallPrinter(HWND hwnd, LPLONG lpIValue, LPTSTR installPath)
{
	CCPDFPrinterInstaller installer(hwnd);
	return installer.DoInstall((HMODULE)ghModule, installPath) ? 0 : 1;
}

/**
	@brief Call to remove the CC PDF printer driver
	@param hwnd Handle to the parent window
	@param lpIValue Ignored
	@param lpszValue Ignored
	@return 0 if successfully removed, 1 if failed
*/
LONG APIENTRY RemovePrinter(HWND hwnd, LPLONG lpIValue, LPSTR lpszValue)
{
	CCPDFPrinterInstaller installer(hwnd);
	return installer.DoRemove() ? 0 : 1;
}

/**
	@brief Call to check if the CC PDF Printer is installed
	@param hwnd Handle to the parent window
	@param lpIValue Ignored
	@param lpszValue Ignored
	@return 0 if the printer is installed, 1 if not
*/
LONG APIENTRY IsInstalled(HWND hwnd, LPLONG lpIValue, LPSTR lpszValue)
{
	CCPDFPrinterInstaller installer(hwnd);
	return installer.IsPrinterInstalled() ? 0 : 1;
}
