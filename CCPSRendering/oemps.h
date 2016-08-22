/**
	@file
	@brief Public definitions for printer driver rendering hooks
			Based on oemps.h
			Printer Driver Rendering Plugin Sample
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

#ifndef _OEMPS_H
#define _OEMPS_H

#include "DEVMODE.H"
#include "CCPrintData.h"
#include "TextPart.h"

/**
	Escape code for adding a link to the current page.
	To use, send the \ref EscapeLinkData structure as the data
*/
#define ESCAPE_LINK_DATA		0x667711aa
/// Escape code: disable Auto URL link (for this print job only)
#define ESCAPE_DISABLE_AUTO_URL	0x667711ab

////////////////////////////////////////////////////////
//      OEM Defines
////////////////////////////////////////////////////////


///////////////////////////////////////////////////////
// Warning: the following enum order must match the 
//          order in OEMHookFuncs[].
///////////////////////////////////////////////////////

/// Enumeration of hooked function
typedef enum tag_Hooks {
    UD_DrvStartPage,
    UD_DrvSendPage,
    UD_DrvStartDoc,
    UD_DrvEndDoc,
	UD_DrvEscape,
	UD_DrvTextOut,

    MAX_DDI_HOOKS,

} ENUMHOOKS;

/**
    @brief Structure holding link data for a single link on the page
*/
struct EscapeLinkData
{
	/// The left border of the link
	long left;
	/// The top border of the link
	long top;
	/// The right border of the link
	long right;
	/// The bottom border of the link
	long bottom;
	/// Offset of the link's tooltip from /ref url (0 for no tooltip) (future use only)
	size_t lTitleOffset;
	/// Null-terminated URL for the link; if there's a tooltip string, it follow immediately
	char url[1];
};

/**
    @brief Class holding a single link's data for writing into the page
*/
struct InnerEscapeLinkData
{
	/**
		@brief Default constructor
	*/
	InnerEscapeLinkData() : pNext(NULL), pData(NULL) {};
	/**
		@brief Constructor from escape sequence data
		@param pInData Pointer to an EscapeLinkData structure
		@param nSize Size of the structure (including the url part)
		@param pInNext Pointer to the next link data object
	*/
	InnerEscapeLinkData(const char* pInData, int nSize, InnerEscapeLinkData* pInNext) : pNext(pInNext) {pData = (EscapeLinkData*)new char[nSize]; memcpy(pData, pInData, nSize);};
	/**
		@brief Constructor from link-file data
		@param rect Location of the link on the page
		@param pURL URL of the link
		@param pInNext Pointer to the next link data object
		@param pTitle Tooltip for the link (future use only)
	*/
	InnerEscapeLinkData(const RECTL& rect, const char* pURL, InnerEscapeLinkData* pInNext, const char* pTitle = NULL) : pNext(pInNext)
	{
		// Calculate the structure's size
		size_t nSize = sizeof(EscapeLinkData) + strlen(pURL);
		if ((pTitle != NULL) && (strlen(pTitle) > 0))
			nSize += strlen(pTitle) + 1;

		// Create structure
		pData = (EscapeLinkData*)new char[nSize];
		// Populate it
		pData->left = rect.left;
		pData->right = rect.right;
		pData->top = rect.top;
		pData->bottom = rect.bottom;
		strcpy_s(pData->url, strlen(pURL)+1, pURL);
		if (pTitle == NULL)
			pData->lTitleOffset = 0;
		else
		{
			pData->lTitleOffset = strlen(pURL) + 1;
			strcpy_s(pData->url + pData->lTitleOffset, strlen(pTitle), pTitle);
		}
	}
	/**
		@brief Destructor
	*/
	~InnerEscapeLinkData() {delete [] pData;};
	/// Pointer to the next link data object
	InnerEscapeLinkData* pNext;
	/// Pointer to the link data
	EscapeLinkData* pData;
};

/// Internal printing data object
typedef struct _OEMPDEV {
    //
    // define whatever needed, such as working buffers, tracking information,
    // etc.
    //
    // This test DLL hooks out every drawing DDI. So it needs to remember
    // PS's hook function pointer so it call back.
    //
    PFN     pfnPS[MAX_DDI_HOOKS];

    //
    // define whatever needed, such as working buffers, tracking information,
    // etc.
    //

	/// Reserved
    DWORD					dwReserved[1];
	/// Pointer to the PDEV object
	class IOemPS*			pOemPS;
	/// Pointer to the driver object
    struct IPrintOemDriverPS*  pOEMHelp;
	/// Current page
	UINT					nPage;
	/// Runtime glyph translation
	class GlyphTranslator*	pTranslator;
	/// Internal runtime information structure
	InnerEscapeLinkData*	pLinks;
	/// Text keeping flag: set true to remember the printed text with its location
	bool					bNeedText;
	/// Set to true if loaded data from a link INI file
	bool					bLoadedData;
	/// Current page text data
	TextArea				oText;
	/// link INI file data
	CCPrintData				dataLinks;
	/// Actual printing flag: true if data was actually printed
	bool					bUsedPrintData;

} OEMPDEV, *POEMPDEV;

#endif
