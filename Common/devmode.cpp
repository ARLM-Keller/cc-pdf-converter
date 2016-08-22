/**
	@file
	@brief Implementation of Devmode functions shared with OEM UI and OEM rendering modules.
			Based on:
			Devmode.cpp
			Printer Driver Plugin Sample
			by Microsoft Corporation
		
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
#include "debug.h"
#include "devmode.h"
#include "CCCommon.h"

#include "CCPrintLicenseInfo.h"
#include "CCPrintRegistry.h"
#include "Helpers.h"

typedef struct tagOEMDEV_NoExtra
{
	/// Extra data header: must be included according to the DDK
    OEM_DMEXTRAHEADER   dmOEMExtra;
	/// Filename of the document to be created
	WCHAR				cFilename[MAX_PATH + 1];
	/// TRUE to open the document after creation, FALSE not to
	BOOL				bAutoOpen;
	/// TRUE to add license information to the document's XMP information, FALSE not to
	BOOL				bSetProperties;
	/// The location of the license stamp
	LicenseLocationInfo location;
	/// The license information
	LicenseInfo			info;
} OEMDEV_NoExtra, *POEMDEV_NoExtra;

typedef struct tagOEMDEV_NoTrans
{
	/// Extra data header: must be included according to the DDK
    OEM_DMEXTRAHEADER   dmOEMExtra;
	/// Filename of the document to be created
	WCHAR				cFilename[MAX_PATH + 1];
	/// TRUE to open the document after creation, FALSE not to
	BOOL				bAutoOpen;
	/// TRUE to add license information to the document's XMP information, FALSE not to
	BOOL				bSetProperties;
	/// The location of the license stamp
	LicenseLocationInfo location;
	/// The license information
	LicenseInfo			info;
	/// Internal runtime information structure
	void*				pExtra;
} OEMDEV_NoTrans, *POEMDEV_NoTrans;

typedef struct tagOEMDEV_Runtime
{
	/// Extra data header: must be included according to the DDK
    OEM_DMEXTRAHEADER   dmOEMExtra;
	/// Filename of the document to be created
	WCHAR				cFilename[MAX_PATH + 1];
	/// TRUE to open the document after creation, FALSE not to
	BOOL				bAutoOpen;
	/// TRUE to add license information to the document's XMP information, FALSE not to
	BOOL				bSetProperties;
	/// The location of the license stamp
	LicenseLocationInfo location;
	/// The license information
	LicenseInfo			info;
	/// Runtime glyph translation
	class GlyphTranslator* pTranslator;
	/// Internal runtime information structure
	void*				pExtra;

} OEMDEV_Runtime, *POEMDEV_Runtime;

typedef struct tagOEMDEV_NoURLs
{
	/// Extra data header: must be included according to the DDK
    OEM_DMEXTRAHEADER   dmOEMExtra;
	/// Filename of the document to be created
	WCHAR				cFilename[MAX_PATH + 1];
	/// TRUE to open the document after creation, FALSE not to
	BOOL				bAutoOpen;
	/// TRUE to add license information to the document's XMP information, FALSE not to
	BOOL				bSetProperties;
	/// The location of the license stamp
	LicenseLocationInfo location;
	/// The license information
	LicenseInfo			info;

} OEMDEV_NoURLs, *POEMDEV_NoURLs;


/**
	@param bFirstPage true if the requested information is for the first page
	@param szPageSize Size of the page
	@param szLicenseSize Size of the stamp
	@return Location of stamp on the page
*/
POINT LicenseLocationInfo::LocationForPage(bool bFirstPage, SIZE szPageSize, SIZE szLicenseSize) const
{
	LicenseLocation eLocation;
	POINT ptUser;

	// Should we use the first page information?
	if (bFirstPage || (eOtherPages == LLOther))
	{
		// Yeah
		eLocation = eFirstPage;
		ptUser = ptFirstPage;
	}
	else
	{
		// Nope, it's for the other pages
		eLocation = eOtherPages;
		ptUser = ptOtherPages;
	}
	
	// Initialize point
	POINT ptRet;
	ptRet.x = ptRet.y = -1;
	
	// Calculate according to the stamp location information:
	switch (eLocation)
	{
		case LLTop:
			ptRet.y = 0;
			ptRet.x = (szPageSize.cx - szLicenseSize.cx) / 2;
			break;
		case LLBottom:
			ptRet.y = szPageSize.cy - szLicenseSize.cy;
			ptRet.x = (szPageSize.cx - szLicenseSize.cx) / 2;
			break;
		case LLUserDefined:
			ptRet.x = (szPageSize.cx - szLicenseSize.cx) * ptUser.x / 100;
			ptRet.y = (szPageSize.cy - szLicenseSize.cy) * ptUser.y / 100;

			break;
		default:
			break;
	}

	return ptRet;
}

/**
	@param pOEMDevOut Pointer to the structure to write data into
	@param hPrinter Handle to the printer
*/
void InitToDefault(POEMDEV pOEMDevOut, HANDLE hPrinter)
{
    pOEMDevOut->dmOEMExtra.dwSize       = sizeof(OEMDEV);
    pOEMDevOut->dmOEMExtra.dwSignature  = CCPRINT_SIGNATURE;
    pOEMDevOut->dmOEMExtra.dwVersion    = CCPRINT_VERSION;
	pOEMDevOut->cFilename[0]			= '\0';
	ReadOEMDevFromRegistry(pOEMDevOut, hPrinter);
}

/**
	@param dwMode Initialization mode information (see IPrintOemPS::DevMode function in the DDK)
	@param pOemDMParam Pointer to an OEMDMPARAM structure containing printer data
	@return S_OK if successful, E_FAIL if failed
*/
HRESULT hrOEMDevMode(DWORD dwMode, POEMDMPARAM pOemDMParam)
{
    POEMDEV pOEMDevIn;
    POEMDEV pOEMDevOut;

    // Verify parameters.
    if( (NULL == pOemDMParam)
        ||
        ( (OEMDM_SIZE != dwMode)
          &&
          (OEMDM_DEFAULT != dwMode)
          &&
          (OEMDM_CONVERT != dwMode)
          &&
          (OEMDM_MERGE != dwMode)
        )
      )
    {
        ERR(ERRORTEXT("DevMode() ERROR_INVALID_PARAMETER.\r\n"));
        VERBOSE(DLLTEXT("\tdwMode = %d, pOemDMParam = %#lx.\r\n"), dwMode, pOemDMParam);

        SetLastError(ERROR_INVALID_PARAMETER);
        return E_FAIL;
    }

    // Cast generic (i.e. PVOID) to OEM private devomode pointer type.
    pOEMDevIn = (POEMDEV) pOemDMParam->pOEMDMIn;
    pOEMDevOut = (POEMDEV) pOemDMParam->pOEMDMOut;

    switch(dwMode)
    {
        case OEMDM_SIZE:
			// Just needs the size of the structure
            pOemDMParam->cbBufSize = sizeof(OEMDEV);
            break;

        case OEMDM_DEFAULT:
			// Fill in with default parameters
			InitToDefault(pOEMDevOut, pOemDMParam->hPrinter);
            VERBOSE(DLLTEXT("pOEMDevOut after setting default values:\r\n"));
            Dump(pOEMDevOut);
            break;

        case OEMDM_CONVERT:
			// Convert data from old version
            ConvertOEMDevmode(pOEMDevIn, pOEMDevOut, pOemDMParam->hPrinter);
            break;

        case OEMDM_MERGE:
			// Merge data from old and new structures
            ConvertOEMDevmode(pOEMDevIn, pOEMDevOut, pOemDMParam->hPrinter);
            MakeOEMDevmodeValid(pOEMDevOut);
            break;
    }
	
	// Write it up
    Dump(pOemDMParam);

    return S_OK;
}

/**
	@param pOEMDevIn Pointer to the structure to read data from
	@param pOEMDevOut Pointer to the structure to write data into
	@param hPrinter Handle to the printer
	@return TRUE if all went well, FALSE if there's a problem
*/
BOOL ConvertOEMDevmode(PCOEMDEV pOEMDevIn, POEMDEV pOEMDevOut, HANDLE hPrinter)
{
    if( (NULL == pOEMDevIn)
        ||
        (NULL == pOEMDevOut)
      )
    {
        ERR(ERRORTEXT("ConvertOEMDevmode() invalid parameters.\r\n"));
        return FALSE;
    }

    // Check OEM Signature, if it doesn't match ours,
    // then just assume DMIn is bad and use defaults.
    if(pOEMDevIn->dmOEMExtra.dwSignature == pOEMDevOut->dmOEMExtra.dwSignature)
    {
        VERBOSE(DLLTEXT("Converting private OEM Devmode.\r\n"));
        VERBOSE(DLLTEXT("pOEMDevIn:\r\n"));
        Dump(pOEMDevIn);

        // Copy the old structure in to the new using which ever size is the smaller.
        // Devmode maybe from newer Devmode (not likely since there is only one), or
        // Devmode maybe a newer Devmode, in which case it maybe larger,
        // but the first part of the structure should be the same.

        // DESIGN ASSUMPTION: the private DEVMODE structure only gets added to;
        // the fields that are in the DEVMODE never change only new fields get added to the end.
		DWORD dwOut = pOEMDevOut->dmOEMExtra.dwSize;

		switch (pOEMDevIn->dmOEMExtra.dwVersion)
		{
			case CCPRINT_VERSION_NO_EXTRA:
			case CCPRINT_VERSION_NO_TRANS:
			case CCPRINT_VERSION_RUNTIME:
			case CCPRINT_VERSION_NO_URLS:
				if (pOEMDevIn->dmOEMExtra.dwSize < sizeof(OEMDEV_NoURLs))
					return FALSE;
				memcpy(pOEMDevOut, pOEMDevIn, sizeof(OEMDEV_NoURLs));
				break;
			case CCPRINT_VERSION:
				if (pOEMDevIn->dmOEMExtra.dwSize < sizeof(OEMDEV))
					return FALSE;
				memcpy(pOEMDevOut, pOEMDevIn, sizeof(OEMDEV));
				break;
			default:
				WARNING(DLLTEXT("Unknown DEVMODE version, pOEMDMIn ignored.\r\n"));
				InitToDefault(pOEMDevOut, hPrinter);
				break;
		};

		pOEMDevOut->dmOEMExtra.dwSize = dwOut;
    }
    else
    {
        WARNING(DLLTEXT("Unknown DEVMODE signature, pOEMDMIn ignored.\r\n"));

		InitToDefault(pOEMDevOut, hPrinter);
    }

    return TRUE;
}


/**
	@param pOEMDevmode Pointer to the printer data structure to initialize
	@return TRUE if initialized successfully, FALSE if failed
*/
BOOL MakeOEMDevmodeValid(POEMDEV pOEMDevmode)
{
    if(NULL == pOEMDevmode)
    {
		// Bad parameter!
        return FALSE;
    }

    // ASSUMPTION: pOEMDevmode is large enough to contain OEMDEV structure.

    // Make sure that dmOEMExtra indicates the current OEMDEV structure.
    pOEMDevmode->dmOEMExtra.dwSize       = sizeof(OEMDEV);
    pOEMDevmode->dmOEMExtra.dwSignature  = CCPRINT_SIGNATURE;
    pOEMDevmode->dmOEMExtra.dwVersion    = CCPRINT_VERSION;

	int i;
	for (i=0;i<MAX_PATH+1;i++)
	{
		if (pOEMDevmode->cFilename[i] == '\0')
			break;
	}
	if (i == MAX_PATH + 1)
		pOEMDevmode->cFilename[0] = '\0';

    return TRUE;
}

/**
	@param pOEMDevmode Pointer to the printer data structure to dump
*/
void Dump(PCOEMDEV pOEMDevmode)
{
    if( (NULL != pOEMDevmode)
        &&
        (pOEMDevmode->dmOEMExtra.dwSize >= sizeof(OEMDEV_NoURLs))
        &&
        (CCPRINT_SIGNATURE == pOEMDevmode->dmOEMExtra.dwSignature)
      )
    {
        VERBOSE(__TEXT("\tdmOEMExtra.dwSize      = %d\r\n"), pOEMDevmode->dmOEMExtra.dwSize);
        VERBOSE(__TEXT("\tdmOEMExtra.dwSignature = %#x\r\n"), pOEMDevmode->dmOEMExtra.dwSignature);
        VERBOSE(__TEXT("\tdmOEMExtra.dwVersion   = %#x\r\n"), pOEMDevmode->dmOEMExtra.dwVersion);
        VERBOSE(__TEXT("\tOutputFile             = %.*s\r\n"), MAX_PATH, pOEMDevmode->cFilename);
        VERBOSE(__TEXT("\tAutoOpen               = %s\r\n"), pOEMDevmode->bAutoOpen ? __TEXT("Yes") : __TEXT("No"));
        VERBOSE(__TEXT("\tWriteProperties        = %s\r\n"), pOEMDevmode->bSetProperties ? __TEXT("Yes") : __TEXT("No"));
        VERBOSE(__TEXT("\tLocation               = %u\r\n"), (UINT)pOEMDevmode->location.eFirstPage);
        VERBOSE(__TEXT("\tLocation(2)            = %u\r\n"), (UINT)pOEMDevmode->location.eOtherPages);
        VERBOSE(__TEXT("\tLicense                = %u\r\n"), (UINT)pOEMDevmode->info.m_eLicense);
		if ((pOEMDevmode->dmOEMExtra.dwVersion > CCPRINT_VERSION_NO_URLS) && (pOEMDevmode->dmOEMExtra.dwSize >= sizeof(OEMDEV)))
		{
	        VERBOSE(__TEXT("\tAutoURLs                = %u\r\n"), (UINT)pOEMDevmode->bAutoURLs);
		}
    }
    else
    {
        ERR(ERRORTEXT("Dump(POEMDEV) unknown private OEM DEVMODE.\r\n"));
    }
}

/**
	@param pDev Pointer to the plugin data
	@param hPrinter Handle to the printer
*/
void ReadOEMDevFromRegistry(POEMDEV pDev, HANDLE hPrinter)
{
	// Read the auto open and XMP settings
	pDev->bAutoOpen = CCPrintRegistry::GetRegistryBool(hPrinter, (LPTSTR)SETTINGS_AUTOOPEN, CanOpenPDFFiles());
	pDev->bSetProperties = CCPrintRegistry::GetRegistryBool(hPrinter, (LPTSTR)SETTINGS_WRITEPROPERTIES, true);
	pDev->bAutoURLs = CCPrintRegistry::GetRegistryBool(hPrinter, (LPTSTR)SETTINGS_AUTOURLS, true);
	pDev->bCreateAsTemp = CCPrintRegistry::GetRegistryBool(hPrinter, (LPTSTR)SETTINGS_CREATEASTEMP, false);

	// Read the license and stamp locations
	CCPrintLicenseInfo::ReadFromRegistry(hPrinter, pDev->info);
	CCPrintLicenseInfo::ReadFromRegistry(hPrinter, pDev->location);
}

/**
	@param pDev Pointer to the plugin data
	@param hPrinter Handle to the printer
	@return true if written successfully, false if failed
*/
bool WriteOEMDevToRegistry(const POEMDEV pDev, HANDLE hPrinter)
{
	// Write the auto open and XMP settings
	bool bRet = true;
	bRet &= CCPrintRegistry::SetRegistryBool(hPrinter, (LPTSTR)SETTINGS_AUTOOPEN, pDev->bAutoOpen ? true : false);
	bRet &= CCPrintRegistry::SetRegistryBool(hPrinter, (LPTSTR)SETTINGS_WRITEPROPERTIES, pDev->bSetProperties ? true : false);
	bRet &= CCPrintRegistry::SetRegistryBool(hPrinter, (LPTSTR)SETTINGS_AUTOURLS, pDev->bAutoURLs ? true : false);
	bRet &= CCPrintRegistry::SetRegistryBool(hPrinter, (LPTSTR)SETTINGS_CREATEASTEMP, pDev->bCreateAsTemp ? true : false);

	// Write the license and stamp locations
	bRet &= CCPrintLicenseInfo::WriteToRegistry(hPrinter, pDev->info);
	bRet &= CCPrintLicenseInfo::WriteToRegistry(hPrinter, pDev->location);
	return bRet;
}

