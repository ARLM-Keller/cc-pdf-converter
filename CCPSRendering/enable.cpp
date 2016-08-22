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
#include "debug.h"
#include "oemps.h"
#include "GlyphTranslator.h"


////////////////////////////////////////////////////////
//      Internal Constants
////////////////////////////////////////////////////////

/// Hook functions (must mach the order of ENUMHOOKS enum
static const DRVFN OEMHookFuncs[] =
{
    { INDEX_DrvStartPage,                   (PFN) OEMStartPage                  },
    { INDEX_DrvSendPage,                    (PFN) OEMSendPage                   },
    { INDEX_DrvStartDoc,                    (PFN) OEMStartDoc                   },
    { INDEX_DrvEndDoc,                      (PFN) OEMEndDoc                     },
	{ INDEX_DrvEscape,						(PFN) OEMEscape						},
	{ INDEX_DrvTextOut,						(PFN) OEMTextOut					},
};




/**
	@param pdevobj Pointer to the DEVOBJ structure
	@param pPrinterName Pointer to the printer name
	@param cPatterns Number of surface patterns in phsurtPatterns
	@param phsurfPatterns Pointer to surface patters handles array
	@param cjGdiInfo Size of pGdiInfo structure
	@param pGdiInfo Pointer to GDIINFO structure
	@param cjDevInfo Size of pDevInfo structure
	@param pDevInfo Pointer to a DEVINFO structure
	@param pded Pointer to a DRVENABLEDATA structure to fill with hooking function pointers
	@return Pointer to the private PDEV structure, if successful, or NULL if failed
*/
PDEVOEM APIENTRY OEMEnablePDEV(PDEVOBJ pdevobj, PWSTR pPrinterName, ULONG cPatterns, HSURF *phsurfPatterns, ULONG cjGdiInfo, GDIINFO *pGdiInfo, ULONG cjDevInfo, DEVINFO *pDevInfo, DRVENABLEDATA *pded)
{
    POEMPDEV    poempdev;
    INT         i, j;
    DWORD       dwDDIIndex;
    PDRVFN      pdrvfn;

    VERBOSE(DLLTEXT("OEMEnablePDEV() entry.\r\n"));

    //
    // Allocate the OEMDev
    //
    poempdev = new OEMPDEV;
    if (NULL == poempdev)
    {
        return NULL;
    }

    //
    // Fill in OEMDEV as you need
    //
	poempdev->nPage = 0;
	poempdev->pLinks = NULL;
	poempdev->pTranslator = NULL;
	POEMDEV pDevMode = (POEMDEV)pdevobj->pOEMDM;
	poempdev->bNeedText = pDevMode->bAutoURLs ? true : false;

    //
    // Fill in OEMDEV
    //
    for (i = 0; i < MAX_DDI_HOOKS; i++)
    {
        //
        // search through Unidrv's hooks and locate the function ptr
        //
        dwDDIIndex = OEMHookFuncs[i].iFunc;
        for (j = pded->c, pdrvfn = pded->pdrvfn; j > 0; j--, pdrvfn++)
        {
            if (dwDDIIndex == pdrvfn->iFunc)
            {
                poempdev->pfnPS[i] = pdrvfn->pfn;
                break;
            }
        }
        if (j == 0)
        {
            //
            // didn't find the Unidrv hook. Should happen only with DrvRealizeBrush
            //
            poempdev->pfnPS[i] = NULL;
        }

    }

    return (POEMPDEV) poempdev;
}

/**
	@param pdevobj Pointer to the DEVOBJ structure
*/
VOID APIENTRY OEMDisablePDEV(PDEVOBJ pdevobj)
{
    VERBOSE(DLLTEXT("OEMDisablePDEV() entry.\r\n"));


    //
    // Free memory for OEMPDEV and any memory block that hangs off OEMPDEV.
    //
    assert(NULL != pdevobj->pdevOEM);
    POEMPDEV poempdev = (POEMPDEV)pdevobj->pdevOEM;
	while (poempdev->pLinks != NULL)
	{
		InnerEscapeLinkData* pLink = poempdev->pLinks;
		poempdev->pLinks = pLink->pNext;
		delete pLink;
	}
	if (poempdev->pTranslator != NULL)
	{
		delete poempdev->pTranslator;
		poempdev->pTranslator = NULL;
	}
    delete pdevobj->pdevOEM;
}

/**
	@param pdevobjOld Pointer to the original DEVOBJ structure
	@param pdevobjNew Pointer to the new DEVOBJ structure
	@return 
*/
BOOL APIENTRY OEMResetPDEV(PDEVOBJ pdevobjOld, PDEVOBJ pdevobjNew)
{
    VERBOSE(DLLTEXT("OEMResetPDEV() entry.\r\n"));


    //
    // If you want to carry over anything from old pdev to new pdev, do it here.
    //
    POEMPDEV poempdevOld = (POEMPDEV)pdevobjOld->pdevOEM;
    POEMPDEV poempdevNew = (POEMPDEV)pdevobjNew->pdevOEM;
	if (poempdevNew->pTranslator == NULL)
	{
		poempdevNew->pTranslator = poempdevOld->pTranslator;
		poempdevOld->pTranslator = NULL;
	}

    return TRUE;
}

/**
*/
VOID APIENTRY OEMDisableDriver()
{
    VERBOSE(DLLTEXT("OEMDisableDriver() entry.\r\n"));
}

/**
	@param dwOEMintfVersion Version of driver (should be PRINTER_OEMINTF_VERSION)
	@param dwSize Size of DEVENABLEDATA structure in pded
	@param pded Pointer to a DEVENABLEDATA structure
	@return TRUE
*/
BOOL APIENTRY OEMEnableDriver(DWORD dwOEMintfVersion, DWORD dwSize, PDRVENABLEDATA pded)
{
    VERBOSE(DLLTEXT("OEMEnableDriver() entry.\r\n"));

    // List DDI functions that are hooked.
    pded->iDriverVersion =  PRINTER_OEMINTF_VERSION;
    pded->c = sizeof(OEMHookFuncs) / sizeof(DRVFN);
    pded->pdrvfn = (DRVFN *) OEMHookFuncs;

    return TRUE;
}
