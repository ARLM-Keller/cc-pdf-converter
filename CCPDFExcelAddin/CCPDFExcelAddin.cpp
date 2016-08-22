/**
	@file
	@brief CCPDFExcelAddin.cpp : Implementation of DLL Exports.
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
#include "resource.h"
#include <initguid.h>
#ifdef CC_PDF_CONVERTER
#include "CCPDFExcelAddin.h"
#elif EXCEL_TO_PDF
#include "XL2PDFExcelAddin.h"
#else
#error "Please define one of the printer types"
#endif
#include "dlldatax.h"

#ifdef CC_PDF_CONVERTER
#include "CCPDFExcelAddin_i.c"
#elif EXCEL_TO_PDF
#include "XL2PDFExcelAddin_i.c"
#else
#error "Please define one of the printer types"
#endif
#include "CCPDFExcelAddinObj.h"

#ifdef _MERGE_PROXYSTUB
/// Handle to proxy DLL
extern "C" HINSTANCE hProxyDll;
#endif

/// Module object
CComModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
#ifdef CC_PDF_CONVERTER
OBJECT_ENTRY(CLSID_CCPDFExcelAddinObj, CCCPDFExcelAddinObj)
#elif EXCEL_TO_PDF
OBJECT_ENTRY(CLSID_XL2PDFExcelAddinObj, CCCPDFExcelAddinObj)
#else
#error "Please define one of the printer types"
#endif
END_OBJECT_MAP()

/////////////////////////////////////////////////////////////////////////////
// DLL Entry Point

/**
	@param hInstance Handle to the DLL instance
	@param dwReason Reason of call
	@param lpReserved Not used
	@return TRUE
*/
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    lpReserved;
#ifdef _MERGE_PROXYSTUB
    if (!PrxDllMain(hInstance, dwReason, lpReserved))
        return FALSE;
#endif
    if (dwReason == DLL_PROCESS_ATTACH)
    {
#ifdef CC_PDF_CONVERTER
        _Module.Init(ObjectMap, hInstance, &LIBID_CCPDFEXCELADDINLib);
#elif EXCEL_TO_PDF
        _Module.Init(ObjectMap, hInstance, &LIBID_XL2PDFEXCELADDINLib);
#else
#error "Please define one of the printer types"
#endif
        DisableThreadLibraryCalls(hInstance);
    }
    else if (dwReason == DLL_PROCESS_DETACH)
        _Module.Term();
    return TRUE;    // ok
}

/**
	@brief Used to determine whether the DLL can be unloaded by OLE
	@return S_OK if can be unloaded, S_FALSE otherwise
*/
STDAPI DllCanUnloadNow(void)
{
#ifdef _MERGE_PROXYSTUB
    if (PrxDllCanUnloadNow() != S_OK)
        return S_FALSE;
#endif
    return (_Module.GetLockCount()==0) ? S_OK : S_FALSE;
}

/**
	@brief Returns a class factory to create an object of the requested type
	@param rclsid The requested CLSID
	@param riid The requested IID
	@param ppv Pointer to the factory interface
	@return S_OK for success, error code if failed
*/
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
#ifdef _MERGE_PROXYSTUB
    if (PrxDllGetClassObject(rclsid, riid, ppv) == S_OK)
        return S_OK;
#endif
    return _Module.GetClassObject(rclsid, riid, ppv);
}

/**
	@brief Adds entries to the system registry
	@return S_OK for success, error code if failed
*/
STDAPI DllRegisterServer(void)
{
#ifdef _MERGE_PROXYSTUB
    HRESULT hRes = PrxDllRegisterServer();
    if (FAILED(hRes))
        return hRes;
#endif
    // registers object, typelib and all interfaces in typelib
    return _Module.RegisterServer(TRUE);
}

/**
	@brief Removes entries from the system registry
	@return S_OK for success, error code if failed
*/
STDAPI DllUnregisterServer(void)
{
#ifdef _MERGE_PROXYSTUB
    PrxDllUnregisterServer();
#endif
    return _Module.UnregisterServer(TRUE);
}


