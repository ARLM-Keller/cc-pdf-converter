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
#include "oemps.h"
#include "resource.h"
#include <math.h>
#include <PRCOMOEM.H>
#include "CCTChar.h"
#include "CCPrintRegistry.h"
#include "GlyphTranslator.h"
#include "CCPrintData.h"

#include "intrface.h"
#include "PngImage.h"
#include "SQLiteDB.h"


/// Instance of module (defined at dllentry.cpp)
extern HINSTANCE ghInstance;

/* Helper functions */

/**
	@brief This function retrieves the resource ID of the image associated with the license
	@param info The license information
	@return Resource ID for the image
*/
UINT GetLicenseImage(const LicenseInfo& info)
{
	switch (info.m_eLicense)
	{
		case LicenseInfo::LTCC:
			if (info.m_eModification != LicenseInfo::MTUnknown)
			{
				UINT uBase = IDPNG_BY_NC + (int)info.m_eModification;
				if (info.m_bCommercialUse)
					uBase += 3;
				return uBase;
			}
			break;
		case LicenseInfo::LTSampling:
			if (info.m_eSampling == LicenseInfo::STUnknown)
				break;
			return IDPNG_SAMPLING + (int)info.m_eSampling;
		case LicenseInfo::LTDevelopingNations:
		case LicenseInfo::LTPublicDomain:
			return IDPNG_SOMERIGHTS;
		default:
			break;
	}
	return 0;
}

/**
	@brief This function copies an image from the source surface to the destination surface
	@param pso The destination surface
	@param pSrc The source surface (the image must be attached to it already)
	@param rectTarget The location to draw the image at
	@return TRUE if drawn successfully, FALSE if failed for any reason
*/
BOOL DrawImage(SURFOBJ* pso, SURFOBJ* pSrc, RECTL& rectTarget)
{
	// Create the translate object (empty)
	XLATEOBJ xlate;
	memset(&xlate, 0, sizeof(xlate));

	// Create the color adjustment structure
	COLORADJUSTMENT clr;
	memset(&clr, 0, sizeof(clr));

	// Create the clip region
	CLIPOBJ clip;
	clip.iUniq = 0;
	clip.rclBounds = rectTarget;
	clip.iDComplexity = DC_RECT;
	clip.iFComplexity = FC_RECT;
	clip.iMode = TC_RECTANGLES;
	clip.fjOptions = 0;

	// Brush location
	POINTL pt;
	pt.x = pt.y = 0;

	// Create source location
	RECTL rectSource;
	rectSource.top = rectSource.left = 0;
	rectSource.right = pSrc->sizlBitmap.cx;
	rectSource.bottom = pSrc->sizlBitmap.cy;

	// Do the work...
	BOOL bRet = EngStretchBlt(pso, pSrc, NULL, &clip, &xlate, &clr, &pt, &rectTarget, &rectSource, &pt, COLORONCOLOR);
	
	return bRet;
}

/**
	@brief This function draws a loaded PNG image onto the received surface in the specified location
	@param pso The surface to draw upon
	@param png The image to draw
	@param rectTarget The drawing location
	@return TRUE if drawn successfully, FALSE if failed
*/
BOOL DrawImage(SURFOBJ* pso, const PngImage& png, RECTL& rectTarget)
{
	SIZEL szBitmap;
	szBitmap.cx = png.GetWidth();
	szBitmap.cy = png.GetHeight();

	UINT uFormat;
	switch (png.GetBitsPerPixel())
	{
		case 1:
			uFormat = BMF_1BPP;
			break;
		case 4:
			uFormat = BMF_4BPP;
			break;
		case 8:
			uFormat = BMF_8BPP;
			break;
		case 16:
			uFormat = BMF_16BPP;
			break;
		case 24:
			uFormat = BMF_24BPP;
			break;
		case 32:
			uFormat = BMF_32BPP;
			break;
	}

	HBITMAP hEngBmp = EngCreateBitmap(szBitmap, png.GetWidthInBytes(), uFormat, BMF_TOPDOWN|BMF_USERMEM, (void*)png.GetBits());
	if (hEngBmp == NULL)
		return FALSE;

	SURFOBJ* pSrc = EngLockSurface((HSURF)hEngBmp);
	BOOL bRet = DrawImage(pso, pSrc, rectTarget);
	EngUnlockSurface(pSrc);
	EngDeleteSurface((HSURF)hEngBmp);

	return bRet;
}

/**
	@brief This function draws a loaded BITMAP image onto the received surface in the specified location
	@param pso The surface to draw upon
	@param hBmp The bitmap to draw
	@param rectTarget The drawing location
	@return TRUE if drawn successfully, FALSE if failed
*/
BOOL DrawImage(SURFOBJ* pso, HBITMAP hBmp, RECTL& rectTarget)
{
	BITMAP bmp;
	::GetObject(hBmp, sizeof(bmp), &bmp);
	SIZEL szBitmap;
	szBitmap.cx = bmp.bmWidth;
	szBitmap.cy = bmp.bmHeight;
	UINT uFormat;
	switch (bmp.bmBitsPixel)
	{
		case 1:
			uFormat = BMF_1BPP;
			break;
		case 4:
			uFormat = BMF_4BPP;
			break;
		case 8:
			uFormat = BMF_8BPP;
			break;
		case 16:
			uFormat = BMF_16BPP;
			break;
		case 24:
			uFormat = BMF_24BPP;
			break;
		case 32:
			uFormat = BMF_32BPP;
			break;
	}

	HBITMAP hEngBmp = EngCreateBitmap(szBitmap, bmp.bmWidthBytes, uFormat, BMF_USERMEM, (void*)bmp.bmBits);
	if (hEngBmp == NULL)
		return FALSE;

	SURFOBJ* pSrc = EngLockSurface((HSURF)hEngBmp);
	BOOL bRet = DrawImage(pso, pSrc, rectTarget);
	EngUnlockSurface(pSrc);
	EngDeleteSurface((HSURF)hEngBmp);

	return bRet;
}

/// PDFMark URL link box definition
#define URLBOX	"\n[ /Rect [%d %d %d %d]\n\
	/Action << /Subtype /URI /URI (%s) >>\n\
	/Border [0 0 2]\n\
	/Color [.7 0 0]\n\
	/Subtype /Link\n\
	/Title (%s)\n\
	/ANN pdfmark\n"

/// PDFMark hyperlink function definition
#define HYPERLINK_FUNC "/cc_hyperlink { dup show stringwidth pop neg\n\
   gsave 0 currentfont dup\n\
      /FontInfo get /UnderlineThickness get exch\n\
      /FontMatrix get dtransform setlinewidth 0 currentfont dup\n\
      /FontInfo get /UnderlinePosition get exch\n\
      /FontMatrix get dtransform rmoveto rlineto stroke\n\
   grestore\n\
} def\n"

/// PDFMark hyperlink start definition
#define HYPERLINK_START	"\ncurrentpoint\n\
  %d sub\n\
  currentcolor 0 0 128 setrgbcolor\
  ("

/// PDFMark hyperlink close definition
#define HYPERLINK_END ") cc_hyperlink\n\
  setcolor\n\
  currentpoint\n\
  [ /Rect 6 -4 roll 4 array astore\n\
    /Action << /Subtype /URI /URI (%s) >>\n\
    /Border [0 0 2]\n\
    /Color [.7 0 0]\n\
    /Subtype /Link\n\
    /ANN pdfmark\n"

/// PDFMark internal document link box definition
#define JUMPBOX "\n[ /Rect [%d %d %d %d]\n\
	/Border [0 0 2]\n\
	/Color [.7 0 0]\n\
	/Dest /%s\n\
	/Title (%s)\n\
	/Subtype /Link\n\
	/ANN pdfmark\n"

/// PDFMark internal document destination definition
#define JUMPDEST "\n[ /Dest /%s\n\
	/View [/Fit]\n\
	/DEST pdfmark\n"

/// PDFMark internal document link to unnamed destination
#define JUMPINTERNAL_TITLE "\n[ /Rect [%d %d %d %d]\n\
	/Border [0 0 2]\n\
	/Color [.7 0 0]\n\
	/Page %d\n\
	/View [/XYZ %d %d 0]\n\
	/Title (%s)\n\
	/Subtype /Link\n\
	/ANN pdfmark\n"

/// PDFMark internal document link to unnamed destination
#define JUMPINTERNAL "\n[ /Rect [%d %d %d %d]\n\
	/Border [0 0 2]\n\
	/Color [.7 0 0]\n\
	/Page %d\n\
	/View [/XYZ %d %d 0]\n\
	/Subtype /Link\n\
	/ANN pdfmark\n"

/// Postscript text writing (with font) start definition
#define PS_TEXT_START "/Times-Roman findfont [%d 0 0 -%d 0 0 ] makefont setfont\n\
%d %d moveto\n\
("
/// Postscript text (with font) writing end definition
#define PS_TEXT_END ") show\n"

/// Postscript text writing start definition
#define PS_JUSTTEXT_START "("
/// Postscript text writing end definition
#define PS_JUSTTEXT_END ") show \n"

/// Postscript text writing (centered) helper function definition
#define PS_TEXT_END_CENTER ")\n\
dup stringwidth pop\n\
%d exch sub\n\
2 div\n\
0 rmoveto\n\
show\n"

/// Postscript text writing (centered) start definition
#define PS_CALC_CENTER_START "/Times-Roman findfont [%d 0 0 -%d 0 0 ] makefont setfont\n\
%d %d moveto\n\
("

/// Postscript text writing (centered) end definition
#define PS_CALC_CENTER_END ") stringwidth pop\n\
%d exch sub\n\
2 div\n\
0 rmoveto\n"

/// Postscript circle definition
#define PS_CIRCLE "newpath %d %d %d 0 360 arc fill closepath\n"

/// PostScript image start definition
#define PS_IMAGE_START "gsave\n\
%d %d translate\n\
%d %d scale\n\
%d %d %d [%d 0 0 -%d 0 %d] {<\n"

/// PostScript image end definition
#define PS_IMAGE_END "\n>} image\n\
grestore\n";

/// 'Created by' text
#define CREATEDBY_TEXT "The document was created by "
#ifdef CC_PDF_CONVERTER
/// 'Created by' link text
#define CREATEDBY_LINK_TEXT "CC PDF Converter"
/// 'Created by' link
#define CREATEDBY_LINK "http://www.cogniview.com/cc-pdf-converter.php"
#elif EXCEL_TO_PDF
/// 'Created by' link text
#define CREATEDBY_LINK_TEXT "Excel to PDF Converter"
/// 'Created by' link
#define CREATEDBY_LINK "http://www.cogniview.com/excel-to-pdf-converter.php"
#else
#error "One of the printer types must be defined"
#endif

/// Postscript public domain license information data: (note - special fix in ghostscript to handle this!)
#define PS_NO_LICENSE_INFO "\n[ /Rights (False)\n\
	/RightsURL ("

/// Postscript public domain license information data: (note - special fix in ghostscript to handle this!)
#define PS_LICENSE_INFO_START "\n[ /Rights (True)\n\
	/RightsURL ("
#define PS_LICENSE_INFO_CONTINUE ")\n\
	/RightsStatement (This work is licensed under a "
#define PS_LICENSE_INFO_END ")\n\
	/DOCINFO pdfmark\n"

/**
	@brief This function escapes text to be able to write it in PostScript 
	@param lpString The string to fill with the escaped text
	@param lpText The original text
	@param dwLen Length of original text
*/
void AddPSText(LPSTR lpString, LPCSTR lpText, std::tstring::size_type dwLen)
{
	std::tstring::size_type nLoc = strlen(lpString);
	for (std::tstring::size_type i=0;i<dwLen;i++)
	{
		switch (lpText[i])
		{
			case '\n':
			case '\r':
			case '\t':
			case '\b':
			case '\f':
			case '\\':
			case '(':
			case ')':
				lpString[nLoc++] = '\\';
				break;
		}
		lpString[nLoc++] = lpText[i];
	}
	lpString[nLoc] = '\0';
}

/**
	@brief This function escapes text to be able to write it in PostScript 
	@param lpString The string to fill with the escaped text
	@param s The original string to write
*/
void AddPSText(LPSTR lpString, const std::string& s)
{
	AddPSText(lpString, s.c_str(), s.size());
}

/**
	@brief This function writes text to the PostScript file
	@param pdevobj Pointer to the device object representing the PostScript printer
	@param pDevOEM Pointer to the CC PDF Converter render plugin object
	@param lpText The text to write
*/
void PrintPS(PDEVOBJ pdevobj, POEMPDEV pDevOEM, LPCSTR lpText)
{
	DWORD dwResult;
	std::tstring::size_type	dwLen = strlen(lpText);
	pDevOEM->pOEMHelp->DrvWriteSpoolBuf(pdevobj, (void*)lpText, (DWORD)dwLen, &dwResult);
}

/**
	@brief This function creates a circle on the specified location in the PostScript file
	@param pdevobj Pointer to the device object representing the PostScript printer
	@param pDevOEM Pointer to the CC PDF Converter render plugin object
	@param nX X location of the center of the circle
	@param nY Y location of the center of the circle
	@param nRadius Radius of the circle
*/
void PrintCircle(PDEVOBJ pdevobj, POEMPDEV pDevOEM, int nX, int nY, int nRadius)
{
	char cCircle[128];
	sprintf_s(cCircle, _S(cCircle) , PS_CIRCLE, nX, nY, nRadius);
	PrintPS(pdevobj, pDevOEM, cCircle);
}

/**
	@brief This function writes a centered text string into the PostScript file
	@param pdevobj Pointer to the device object representing the PostScript printer
	@param pDevOEM Pointer to the CC PDF Converter render plugin object
	@param nFontSize Size of the font to print at
	@param nX X Location of the start of the string's bound box
	@param nY Y location of the string's bound box
	@param nWidth Width of the string's bound box
	@param lpText The text to write
*/
void CenterText(PDEVOBJ pdevobj, POEMPDEV pDevOEM, int nFontSize, int nX, int nY, int nWidth, LPCSTR lpText)
{
	char cStr[1024];
	sprintf_s(cStr, _S(cStr), PS_CALC_CENTER_START, nFontSize, nFontSize, nX, nY + nFontSize);
	AddPSText(cStr, lpText, (DWORD)strlen(lpText));
	char cEnd[1024];
	sprintf_s(cEnd, _S(cEnd), PS_CALC_CENTER_END, nWidth);
	strcat_s(cStr, _S(cStr), cEnd);
	PrintPS(pdevobj, pDevOEM, cStr);
}

/**
	@brief This function writes an inner-document link box into the PostScript file
	@param pdevobj Pointer to the device object representing the PostScript printer
	@param pDevOEM Pointer to the CC PDF Converter render plugin object
	@param lpDestination Name of the link's destination
	@param rectTarget Location of the link's box
	@param lpTitle Title of the location (will be displayed in a tooltip in Acrobat)
*/
void PrintJumpLink(PDEVOBJ pdevobj, POEMPDEV pDevOEM, LPCSTR lpDestination, const RECTL& rectTarget, LPCSTR lpTitle = NULL)
{
	if (lpTitle == NULL)
		lpTitle = lpDestination;
	char cLink[1024];
	sprintf_s(cLink, _S(cLink), JUMPBOX, rectTarget.left, rectTarget.top, rectTarget.right, rectTarget.bottom, lpDestination, lpTitle);
	PrintPS(pdevobj, pDevOEM, cLink);
}

/**
	@brief This function writes an unnamed inner-document link box into the PostScript file
	@param pdevobj Pointer to the device object representing the PostScript printer
	@param pDevOEM Pointer to the CC PDF Converter render plugin object
	@param lpDestination Name of the link's destination
	@param rectTarget Location of the link's box
	@param lpTitle Title of the location (will be displayed in a tooltip in Acrobat)
*/
void PrintInternalLink(PDEVOBJ pdevobj, POEMPDEV pDevOEM, const RECTL& rectTarget, long lPage, long lX, long lY, LPCSTR lpTitle = NULL)
{
	char cLink[1024];
	if (lpTitle == NULL)
		sprintf_s(cLink, _S(cLink), JUMPINTERNAL, rectTarget.left, rectTarget.top, rectTarget.right, rectTarget.bottom, lPage, lX, lY);
	else
		sprintf_s(cLink, _S(cLink), JUMPINTERNAL_TITLE, rectTarget.left, rectTarget.top, rectTarget.right, rectTarget.bottom, lPage, lX, lY, lpTitle);
	PrintPS(pdevobj, pDevOEM, cLink);
}

/**
	@brief This function writes a destination mark into the PostScript file (target for inner-document links)
	@param pdevobj Pointer to the device object representing the PostScript printer
	@param pDevOEM Pointer to the CC PDF Converter render plugin object
	@param lpDestination Name of the destination (must be NULL terminated)
*/
void PrintJumpDestination(PDEVOBJ pdevobj, POEMPDEV pDevOEM, LPCSTR lpDestination)
{
	char cLink[1024];
	sprintf_s(cLink, _S(cLink), JUMPDEST, lpDestination);
	PrintPS(pdevobj, pDevOEM, cLink);
}

/**
	@brief This function writes a URL link box into the PostScript file
	@param pdevobj Pointer to the device object representing the PostScript printer
	@param pDevOEM Pointer to the CC PDF Converter render plugin object
	@param lpURL The target URL (must be NULL terminated)
	@param rectTarget Location of the link's box
	@param lpTitle The tooltip to display when the mouse is over the link (defaults to the destination URL)
*/
void PrintURLLink(PDEVOBJ pdevobj, POEMPDEV pDevOEM, LPCSTR lpURL, const RECTL& rectTarget, LPCSTR lpTitle = NULL)
{
	char cLink[1024];
	if (lpTitle == NULL)
		lpTitle = lpURL;
	sprintf_s(cLink, _S(cLink), URLBOX, rectTarget.left, rectTarget.top, rectTarget.right, rectTarget.bottom, lpURL, lpTitle);
	DWORD dwResult;
	std::tstring::size_type dwLen = strlen(cLink);
	pDevOEM->pOEMHelp->DrvWriteSpoolBuf(pdevobj, cLink, (DWORD)dwLen, &dwResult);
}

/**
	@brief This function writes a Hyperlinked text into the PostScript file
	@param pdevobj Pointer to the device object representing the PostScript printer
	@param pDevOEM Pointer to the CC PDF Converter render plugin object
	@param nFontSize Size of the font to write in
	@param lpText Text of the link
	@param dwLen Length of the text
	@param lpURL The URL of the link (must be NULL terminated)
*/
void PrintHyperlink(PDEVOBJ pdevobj, POEMPDEV pDevOEM, int nFontSize, LPCSTR lpText, std::tstring::size_type dwLen, LPCSTR lpURL)
{
	char cStr[2048];
	sprintf_s(cStr, _S(cStr), HYPERLINK_START, nFontSize);
	AddPSText(cStr, lpText, dwLen);
	char cEnd[1024];
	sprintf_s(cEnd, _S(cEnd), HYPERLINK_END, lpURL);
	strcat_s(cStr, _S(cStr), cEnd);
	PrintPS(pdevobj, pDevOEM, cStr);
}

/**
	@brief This function prepares a text string so it can be written into PostScript
	@param lpString The string to fill
	@param nFontSize Size of the font to use
	@param nX X location of the start of the text
	@param nY Y location of the start of the text
	@param lpText The text to write
	@param dwLen Length of the text string
	@param nWidth Width of the text's bounding box (-1 so it won't be centered)
	@return The length of the new string

	Note that not specifying nWidth (or setting it to -1) will write the text as it is, and specifying nWidth
	will center the text inside a box starting at nX and being nWidth wide
*/
size_t PrepareWriteString(LPSTR lpString, int nStringSize, int nFontSize, int nX, int nY, LPCSTR lpText, std::tstring::size_type dwLen, int nWidth = -1)
{
	sprintf_s(lpString, nStringSize, PS_TEXT_START, nFontSize, nFontSize, nX, nY);
	AddPSText(lpString, lpText, dwLen);
	if (nWidth == -1)
		strcat_s(lpString, nStringSize, PS_TEXT_END);
	else
	{
		char c[128];
		sprintf_s(c, _S(c), PS_TEXT_END_CENTER, nWidth);
		strcat_s(lpString, nStringSize, c);
	}
	return strlen(lpString);
}

/**
	@brief This function writes text into the postscript file (without a specific location)
	@param pdevobj Pointer to the device object representing the PostScript printer
	@param pDevOEM Pointer to the CC PDF Converter render plugin object
	@param lpText The text to write
*/
void PrintText(PDEVOBJ pdevobj, POEMPDEV pDevOEM, LPCSTR lpText)
{
	char cStr[1024];
	sprintf_s(cStr, _S(cStr), PS_JUSTTEXT_START);
	AddPSText(cStr, lpText, strlen(lpText));
	strcat_s(cStr, _S(cStr), PS_JUSTTEXT_END);
	PrintPS(pdevobj, pDevOEM, cStr);
}

/**
	@brief This function writes text into the postscript file in a specific location, breaking into lines if necessary
	@param pdevobj Pointer to the device object representing the PostScript printer
	@param pDevOEM Pointer to the CC PDF Converter render plugin object
	@param nFontSize Size of the font to write with
	@param nX The X location of the bound box
	@param nY The Y loation of the bound box
	@param nWidth The width of the bound box
	@param nLineHeight Height of each text line
	@param lpText The text to write
	@param bCenter true to center the text, false (default) to align to the left
	@return The height of the printed text
*/
int PrintText(PDEVOBJ pdevobj, POEMPDEV pDevOEM, int nFontSize, int nX, int nY, int nWidth, int nLineHeight, LPCSTR lpText, bool bCenter = false)
{
	char cStr[1024];
	// Calculate how much text we can put in the width:
	DWORD dwResult;
	std::tstring::size_type dwWriteLen;
	std::tstring::size_type dwLen = strlen(lpText);
	int nRet = 0;
	while ((dwLen * nFontSize / 2) > nWidth)
	{
		// Find the last space we can break on:
		DWORD dwBreak = max(1, nWidth / (nFontSize / 2));
		while ((dwBreak > 1) && (lpText[dwBreak] != ' '))
			dwBreak--;
		// Write what we have
		nRet += nLineHeight;
		dwWriteLen = PrepareWriteString(cStr, sizeof(cStr), nFontSize, nX, nY + nRet, lpText, dwBreak, bCenter ? nWidth : -1);
		pDevOEM->pOEMHelp->DrvWriteSpoolBuf(pdevobj, cStr, (DWORD)dwWriteLen, &dwResult);
		lpText += dwBreak + 1;
		dwLen -= dwBreak - 1;
	}
	
	if (dwLen > 0)
	{
		nRet += nLineHeight;
		dwWriteLen = PrepareWriteString(cStr, sizeof(cStr), nFontSize, nX, nY + nRet, lpText, dwLen, bCenter ? nWidth : -1);
		dwWriteLen = strlen(cStr);
		pDevOEM->pOEMHelp->DrvWriteSpoolBuf(pdevobj, cStr, (DWORD)dwWriteLen, &dwResult);
	}

	return nRet;
}

/**
	@brief This function writes a bitmap image directly into the PostScript file
	@param pdevobj Pointer to the device object representing the PostScript printer
	@param pDevOEM Pointer to the CC PDF Converter render plugin object
	@param hBmp The bitmap to write
	@param rectTargetArea In: Horizontal bound box, Y location of the bitmap; Out: Actual location of the bitmap
	@return TRUE if written successfully, FALSE if failed
*/
BOOL PrintImage(PDEVOBJ pdevobj, POEMPDEV pDevOEM, HBITMAP hBmp, RECTL& rectTargetArea)
{
	double dMultiplier = 1.0;
	if (pdevobj->pPublicDM->dmPrintQuality > 0)
		dMultiplier = min(2.0, pdevobj->pPublicDM->dmPrintQuality / 72.0);

	DIBSECTION dib;
	if (::GetObject(hBmp, sizeof(dib), &dib) == 0)
		return FALSE;

	long nDrawWidth = (long) (dib.dsBmih.biWidth * dMultiplier);
	long nDrawHeight = (long) (dib.dsBmih.biHeight * dMultiplier);
	rectTargetArea.left = ((rectTargetArea.left + rectTargetArea.right) / 2) - (nDrawWidth / 2);
	rectTargetArea.right = rectTargetArea.left + nDrawWidth;
	rectTargetArea.bottom = rectTargetArea.top + nDrawHeight;

	char cStr[1024];
	sprintf_s(cStr, _S(cStr), PS_IMAGE_START, rectTargetArea.left, rectTargetArea.top, 
		nDrawWidth, nDrawHeight, 
		dib.dsBm.bmWidth, dib.dsBm.bmHeight, dib.dsBm.bmBitsPixel, dib.dsBm.bmWidth, dib.dsBm.bmHeight, dib.dsBm.bmHeight);
	std::string sWrite(cStr);

	int nByteWidth = dib.dsBmih.biSizeImage / dib.dsBmih.biHeight;
	int nRealByteWidth = dib.dsBmih.biBitCount == 0 ? 8 : (dib.dsBm.bmWidth * 8) / dib.dsBmih.biBitCount;
	for (int i=0;i<dib.dsBm.bmHeight;i++)
	{
		for (int j=0;j<nRealByteWidth;j++)
		{
			sprintf_s(cStr, _S(cStr), "%02x", *(((unsigned char*)dib.dsBm.bmBits) + (i * nByteWidth) + j));
			sWrite += cStr;
		}
		sWrite += "\n";
	}

	sWrite += PS_IMAGE_END;
	DWORD dwWriteLen = (DWORD) sWrite.size(), dwResult;
	DWORD dwRes = pDevOEM->pOEMHelp->DrvWriteSpoolBuf(pdevobj, (void*)sWrite.c_str(), dwWriteLen, &dwResult);
	return (dwRes == S_OK) && (dwWriteLen == dwResult);
}

/**
	@brief This function adds the license page to the PostScript file
	@param pso Pointer to the surface object representing the writing PostScript file
	@return TRUE if written successfully, FALSE if failed to write
*/
BOOL DoLicensePage(SURFOBJ* pso)
{
    PDEVOBJ     pdevobj;
    POEMPDEV    poempdev;

    pdevobj = (PDEVOBJ)pso->dhpdev;
    poempdev = (POEMPDEV)pdevobj->pdevOEM;
	PCOEMDEV pDevMode = (PCOEMDEV)pdevobj->pOEMDM;

	// Print the license destination, as it should reach here:
	PrintJumpDestination(pdevobj, poempdev, "TheLicense");

	// ### later - mind the language
	int nLang = 6;

	// Retrieve the text we wonna write on the page:
	SQLite::DB db;
	std::tstring sPath = CCPrintRegistry::GetRegistryString(pdevobj->hPrinter, _T("DB Path"), _T(""));
	if (sPath.empty())
		return FALSE;
	if (!db.Open(sPath.c_str()))
		return FALSE;

	// Write stuff in proper location:
	int nHeight = pso->sizlBitmap.cy;
	int nLineHeight = nHeight / 60;
	int nFontSize = nLineHeight - 2;
	int nY = nLineHeight * 6;
	int nX = nLineHeight * 2;

	// Find license type:
	TCHAR cQuery[256];
	int nMode;
	switch (pDevMode->info.m_eLicense)
	{
		case LicenseInfo::LTCC:
			nMode = 0;
			_stprintf_s(cQuery, _S(cQuery), _T("SELECT * FROM TblLicenseType WHERE commercial = %d AND derivs = %d AND mode = %d"), pDevMode->info.m_bCommercialUse ? 1 : 0, pDevMode->info.m_eModification, nMode);
			break;
		case LicenseInfo::LTSampling:
			nMode = 1;
			_stprintf_s(cQuery, _S(cQuery), _T("SELECT * FROM TblLicenseType WHERE derivs = %d AND mode = %d"), pDevMode->info.m_eSampling, nMode);
			break;
		case LicenseInfo::LTDevelopingNations:
			nMode = 2;
			_stprintf_s(cQuery, _S(cQuery), _T("SELECT * FROM TblLicenseType WHERE mode = %d"), nMode);
			break;
		default:
			return FALSE;
	}

	// Retrieve the license data record
	SQLite::Recordset records = db.Query(cQuery);
	if (!records.IsValid() || (records.GetRecordCount() == 0))
	{
		std::tstring s = db.GetLastError();
		return FALSE;
	}

	int nLicenseID = records.GetRecord(0).GetNumField(_T("LicenseTypeID"));
	std::string sLicenseShortName = MakeAnsiString(records.GetRecord(0).GetField(_T("LicenseShortName")));
	_stprintf_s(cQuery, _S(cQuery), _T("SELECT LicenseName FROM tblLicenseName WHERE LicenseTypeID = %d AND LangID = %d"), nLicenseID, nLang);
	records = db.Query(cQuery);
	if (!records.IsValid() || (records.GetRecordCount() == 0))
	{
		std::tstring s = db.GetLastError();
		return FALSE;
	}

	// Get the license name
	std::tstring sLicenseName = records.GetRecord(0).GetField(_T("LicenseName"));

	int nJuri = 1;
	if (pDevMode->info.HasJurisdiction())
	{
		// Use the jurisdiction name to find the ID
		_stprintf_s(cQuery, _S(cQuery), _T("SELECT jurisdictionsID FROM tblJurisdictionName WHERE LanguageID = 6 AND JurisdictionName = '%s'"), pDevMode->info.m_cJurisdiction);
		records = db.Query(cQuery);
		if (records.IsValid() && (records.GetRecordCount() > 0))
			nJuri = records.GetRecord(0).GetNumField(_T("jurisdictionsID"));
	}

	// Get the actual jurisdiction record
	_stprintf_s(cQuery, _S(cQuery), _T("SELECT * FROM tblJurisdiction WHERE JurisdictionID = %d AND mode = %d"), nJuri, nMode);
	records = db.Query(cQuery);
	SQLite::Record recJuri;
	if (records.IsValid() && (records.GetRecordCount() > 0))
		recJuri = records.GetRecord(0);

	std::string sVersion, sJuriShortName;
	if (recJuri.IsValid())
	{
		sVersion = MakeAnsiString(recJuri.GetField(_T("Version")));
		sJuriShortName = MakeAnsiString(recJuri.GetField(_T("ShortName")));
		sLicenseName += _T(" ") + MakeTString(sVersion);
	}
	
	// Retrieve the jurisdiction license name
	_stprintf_s(cQuery, _S(cQuery), _T("SELECT JurisdictionName FROM TblJurisdictionName WHERE JurisdictionsID = %d AND LanguageID = %d"), nJuri, nLang);
	records = db.Query(cQuery);
	if (records.IsValid() && (records.GetRecordCount() > 0))
	{
		sLicenseName += _T(" ") + records.GetRecord(0).GetField(_T("JurisdictionName"));
	}
	else
	{
		std::tstring s = db.GetLastError();
		return FALSE;
	}

	// Print the logo on the top
	sPath = CCPrintRegistry::GetRegistryString(pdevobj->hPrinter, _T("Image Path"), sPath.c_str());
	if (sPath.empty() || (*sPath.rbegin() != '\\'))
		sPath += '\\';
	RECTL rectTarget;
	rectTarget.left = 0;
	rectTarget.right = pso->sizlBitmap.cx;
	rectTarget.top = nY;
	rectTarget.bottom = nY + 1;
	HBITMAP hBmp = (HBITMAP)LoadImage(ghInstance, (sPath + _T("CCLogo.bmp")).c_str(), IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION|LR_LOADFROMFILE);
	if (hBmp != NULL)
	{
		if (PrintImage(pdevobj, poempdev, hBmp, rectTarget))
			nY = rectTarget.bottom + nLineHeight;
		::DeleteObject(hBmp);
	}

	// Write hyperlink function...
	PrintPS(pdevobj, poempdev, HYPERLINK_FUNC);
	int nTextHeight;

	CenterText(pdevobj, poempdev, (nFontSize * 3) / 2, 0, nY, pso->sizlBitmap.cx, MakeAnsiString(sLicenseName).c_str());
	std::string sLink = "http://creativecommons.org/licenses/" + sLicenseShortName + "/";
	if (!sVersion.empty())
		sLink += sVersion + "/";
	if (!sJuriShortName.empty())
		sLink += sJuriShortName + "/";
	PrintHyperlink(pdevobj, poempdev, (nFontSize * 3) / 2, MakeAnsiString(sLicenseName).c_str(), sLicenseName.size(), sLink.c_str());

	nY += (nFontSize * 3) / 2;
	nY += nLineHeight;
	
	// Now start writing the text:
	_stprintf_s(cQuery, _S(cQuery), _T("SELECT * FROM tblTextOrder INNER JOIN tblText ON tblText.Text_ID = tblTextOrder.textid LEFT JOIN tblImages ON tblText.Imageid = tblImages.ImageID WHERE tblTextOrder.LicenseTypeID = %d AND tblText.Lang = %d ORDER BY textorder"), nLicenseID, nLang);
	SQLite::Recordset setText = db.Query(cQuery);
	if (!setText.IsValid())
	{
		std::tstring s = db.GetLastError();
		return FALSE;
	}
	
	for (int iText = 0; iText < setText.GetRecordCount(); iText++)
	{
		SQLite::Record recText = setText.GetRecord(iText);
		if (recText.GetNumField(_T("header")) == 1)
		{
			nY += PrintText(pdevobj, poempdev, (nFontSize * 5) / 4, nX, nY, pso->sizlBitmap.cx - (2 * nX), (nLineHeight * 5) / 4, MakeAnsiString(recText.GetField(_T("Data"))).c_str());
		}
		else
		{
			int nImageHeight = 0;
			std::tstring sImage = recText.GetField(_T("ImageFile"));
			if (!sImage.empty())
			{
				hBmp = (HBITMAP)LoadImage(ghInstance, (sPath + sImage).c_str(), IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION|LR_LOADFROMFILE);
				if (hBmp != NULL)
				{
					nY += nLineHeight / 2;
					rectTarget.left = nX;
					rectTarget.top = nY;
					rectTarget.right = nX + nLineHeight * 5;
					rectTarget.bottom = rectTarget.top + 1;
					if (PrintImage(pdevobj, poempdev, hBmp, rectTarget))
						nImageHeight = rectTarget.bottom - nY;
					::DeleteObject(hBmp);
				}
			}
			nTextHeight = PrintText(pdevobj, poempdev, nLineHeight, nX + nLineHeight * 5, nY, pso->sizlBitmap.cx - ((2 * nX) + (nLineHeight * 5)), nLineHeight, MakeAnsiString(recText.GetField(_T("Data"))).c_str());
			if (nImageHeight == 0)
			{
				// Put bullet:
				PrintCircle(pdevobj, poempdev, nX + (nLineHeight * 5 / 2), nY + (nLineHeight / 2), nLineHeight / 4);
			}
			nY += max(nTextHeight, nImageHeight);
		}
		nY += nLineHeight / 2;
	}

	// Bottom of page: put a link to cc pdf converter
	CenterText(pdevobj, poempdev, nFontSize, 0, nHeight - 2 * nLineHeight, pso->sizlBitmap.cx, CREATEDBY_TEXT CREATEDBY_LINK);
	PrintText(pdevobj, poempdev, CREATEDBY_TEXT);
	PrintHyperlink(pdevobj, poempdev, nFontSize, CREATEDBY_LINK_TEXT, (DWORD)strlen(CREATEDBY_LINK_TEXT), CREATEDBY_LINK);

	// Don't do the regular page operations...
    return (((PFN_DrvSendPage)(poempdev->pfnPS[UD_DrvSendPage]))(pso));
}



/**
	@brief This function is called when a new page is printed into the PostScript file
	@param pso Pointer to the surface object representing the writing PostScript file
	@return TRUE if written successfully, FALSE if failed to write
*/
BOOL APIENTRY OEMStartPage(SURFOBJ* pso)
{
    PDEVOBJ     pdevobj;
    POEMPDEV    poempdev;

    VERBOSE(DLLTEXT("OEMStartPage() entry.\r\n"));

    pdevobj = (PDEVOBJ)pso->dhpdev;
    poempdev = (POEMPDEV)pdevobj->pdevOEM;
	poempdev->nPage++;

    //
    // turn around to call PS
    //

	// Only get text if needed
	if (poempdev->bLoadedData)
		poempdev->bNeedText = poempdev->dataLinks.IsTestPage() || poempdev->dataLinks.GetPageData(poempdev->nPage).HasTextLink();
#ifdef _DEBUG
	poempdev->bNeedText = true;
#endif

    return (((PFN_DrvStartPage)(poempdev->pfnPS[UD_DrvStartPage]))(pso));

}

/**
	@brief This function is called when a page has been printed into the PostScript file
	@param pso Pointer to the surface object representing the writing PostScript file
	@return TRUE if written successfully, FALSE if failed to write
*/
BOOL APIENTRY OEMSendPage(SURFOBJ* pso)
{
    PDEVOBJ     pdevobj;
    POEMPDEV    poempdev;

    VERBOSE(DLLTEXT("OEMSendPage() entry.\r\n"));

    pdevobj = (PDEVOBJ)pso->dhpdev;
    poempdev = (POEMPDEV)pdevobj->pdevOEM;

	// Work with external data (i.e., links file)
	POEMDEV pDevMode = (POEMDEV)pdevobj->pOEMDM;
	if (poempdev->dataLinks.HasData())
	{
		poempdev->bUsedPrintData = true;
		RECTL rcArea;

		// Get data for this page
		VERBOSE(DLLTEXT("Processing page %d for links\r\n"), poempdev->nPage);
		if (poempdev->dataLinks.IsTestPage())
		{
			// This is a text print job: we'll write back the link locations in the INI file
			if (poempdev->oText.empty())
			{
				// Not yet
				poempdev->bUsedPrintData = false;
			}
			else
			{
				// OK, this is the real thing. Only work on the first page
				if (poempdev->nPage == 1)
				{
					// Prepare the result object
					CCPrintData dataCompute;
					dataCompute.SetTestPage();

					// Get the link data to test for
					const CCPrintData::PageData& data = poempdev->dataLinks.GetPageData(poempdev->nPage);
					for (CCPrintData::PageData::const_iterator i = data.begin(); i != data.end(); i++)
					{
						// For each link, try to find it
						const CCPrintData::LinkData& link = (*i);
						ASSERT(!link.IsLocation());
						ASSERT(link.nRepeat == 1);

						// We don't know the order, so start at the beginning
						poempdev->oText.InitSearch();
						STRLIST words;
						words.push_back(link.sText);
					
						if (poempdev->oText.SearchFor(words, rcArea, link.nRepeat))
							// Found, add to the results
							dataCompute.AddLink(link.sURL, rcArea, 1);
					}
					// OK, set it up as the results and update the file!
					dataCompute.SetPageSize(1, pso->sizlBitmap);
					dataCompute.UpdateProcessData(pdevobj->hPrinter);
					// Make sure we don't clean up the updated file at the end of the print!
					poempdev->bUsedPrintData = false;
				}
			}
		}
		else
		{
			// Real print job; get the data
			const CCPrintData::PageData& data = poempdev->dataLinks.GetPageData(poempdev->nPage);
			if (!data.empty())
			{
				// Make sure we can search for text links
				poempdev->oText.InitSearch();
				for (CCPrintData::PageData::const_iterator i = data.begin(); i != data.end(); i++)
				{
					// Get next link
					const CCPrintData::LinkData& link = (*i);

					// Is this link a location link?
					if (link.IsLocation())
					{
						// Location link
						if (link.IsInner())
						{
							// Internal link, add it to the list
							PrintInternalLink(pdevobj, poempdev, link.rectLocation, link.nPage, link.ptOffset.x, link.ptOffset.y);
						}
						else
						{
							// External link, add it to the list
							poempdev->pLinks = new InnerEscapeLinkData(link.rectLocation, MakeAnsiString(link.sURL).c_str(), poempdev->pLinks, link.sTitle.empty() ? NULL : MakeAnsiString(link.sTitle).c_str());
							VERBOSE(DLLTEXT("Adding location-based link to %s:\r\n(%d,%d)-(%d,%d)\r\n"), link.sURL.c_str(), link.rectLocation.left, link.rectLocation.top, link.rectLocation.right, link.rectLocation.bottom);
						}
					}
					else
					{
						// Text link: break it into words
						std::tstring::size_type pos = link.sText.find(' '), oldpos = 0;
						STRLIST words;
						while ((oldpos < link.sText.size()) && (pos != std::tstring::npos))
						{
							if (pos > oldpos)
								words.push_back(link.sText.substr(oldpos, pos - oldpos));
							oldpos = pos + 1;
							pos = link.sText.find(' ', oldpos);
						}
						if (oldpos < link.sText.size())
							words.push_back(link.sText.substr(oldpos));

						// Try to find the words
						if (!poempdev->oText.SearchFor(words, rcArea, link.nRepeat))
							break;
						// Found, so mark the location
						poempdev->pLinks = new InnerEscapeLinkData(rcArea, MakeAnsiString(link.sURL).c_str(), poempdev->pLinks, link.sTitle.empty() ? NULL : MakeAnsiString(link.sTitle).c_str());
					}
				}
			}
		}
	}
	else if (pDevMode->bAutoURLs)
	{
		// Find and highlight URLs if so set by the user
		std::wstring sURL;
		RECTL rcArea;
		poempdev->oText.InitSearch();
		while (poempdev->oText.SearchForURL(rcArea, sURL))
			// Found a URL, add it to the list of links
			poempdev->pLinks = new InnerEscapeLinkData(rcArea, MakeAnsiString(sURL).c_str(), poempdev->pLinks);
	}
	poempdev->oText.clear();

	// Do we have links to add to this page?
	if (poempdev->pLinks != NULL)
	{
		while (poempdev->pLinks != NULL)
		{
			// Get the link
			InnerEscapeLinkData* pLink = (InnerEscapeLinkData*)poempdev->pLinks;
			poempdev->pLinks = pLink->pNext;

			// Prepare the necessary data
			RECTL rectTarget;
			rectTarget.left = pLink->pData->left;
			rectTarget.top = pLink->pData->top;
			rectTarget.right = pLink->pData->right;
			rectTarget.bottom = pLink->pData->bottom;

			if (pLink->pData->lTitleOffset > 0)
				// Print with title
				PrintURLLink(pdevobj, poempdev, pLink->pData->url, rectTarget, pLink->pData->url + pLink->pData->lTitleOffset);
			else
				// Print without title
				PrintURLLink(pdevobj, poempdev, pLink->pData->url, rectTarget);

			delete pLink;
		}
	}

	// Check were we write the license info
	bool bFirstPage = poempdev->nPage == 1;
	LicenseLocation eLocation = LLNone;
	if (bFirstPage)
		eLocation = pDevMode->location.eFirstPage;
	else
	{
		eLocation = pDevMode->location.eOtherPages;
		if (eLocation == LLOther)
			eLocation = pDevMode->location.eFirstPage;
	}

	if (bFirstPage && pDevMode->bSetProperties)
	{
		// Write the license info!
		switch (pDevMode->info.m_eLicense)
		{
			case LicenseInfo::LTPublicDomain:
				// Put a public domain license notice:
				{
					char c[2048];
					sprintf_s(c, _S(c), PS_NO_LICENSE_INFO);
					std::string sURL = MakeAnsiString(pDevMode->info.m_cURI);
					AddPSText(c, sURL);
					strcat_s(c, _S(c), PS_LICENSE_INFO_END);
					PrintPS(pdevobj, poempdev, c);
				}
				break;
			case LicenseInfo::LTNone:
			case LicenseInfo::LTUnknown:
				// Do nothing!
				break;
			default:
				// Put the license information notice:
				{
					char c[2048];
					sprintf_s(c, _S(c), PS_LICENSE_INFO_START);
					std::string sURL = MakeAnsiString(pDevMode->info.m_cURI), sName = MakeAnsiString(pDevMode->info.m_cName);
					AddPSText(c, sURL);
					strcat_s(c, _S(c), PS_LICENSE_INFO_CONTINUE);
					AddPSText(c, sName);
					strcat_s(c, _S(c), " license ");
					AddPSText(c, sURL);
					strcat_s(c, _S(c), PS_LICENSE_INFO_END);
					PrintPS(pdevobj, poempdev, c);
				}
				break;
		}

	}

	if (eLocation != LLNone)
	{
		// Load and use PNG
		UINT uImage = GetLicenseImage(pDevMode->info);
		if (uImage > 0)
		{
			// 1. Load the bitmap
			PngImage png;
			if (png.LoadFromResource(uImage, true, ghInstance))
			{
				// Create the target location
				RECTL rectTarget;
				SIZEL szTarget;
				double dMultiplier = 1.0;
				if (pdevobj->pPublicDM->dmPrintQuality > 0)
					dMultiplier = pdevobj->pPublicDM->dmPrintQuality / 72.0;
				szTarget.cx = (long) (png.GetWidth() * dMultiplier);
				szTarget.cy = (long) (png.GetHeight() * dMultiplier);

				POINT ptTarget = pDevMode->location.LocationForPage(bFirstPage, pso->sizlBitmap, szTarget);
				rectTarget.left = ptTarget.x;
				rectTarget.top = ptTarget.y;
				rectTarget.right = rectTarget.left + szTarget.cx;
				rectTarget.bottom = rectTarget.top + szTarget.cy;

				DrawImage(pso, png, rectTarget);

				// Make this a link:
				switch (pDevMode->info.m_eLicense)
				{
					case LicenseInfo::LTCC:
					case LicenseInfo::LTSampling:
					case LicenseInfo::LTDevelopingNations:
						PrintJumpLink(pdevobj, poempdev, "TheLicense", rectTarget);
						break;
					default:
						if (pDevMode->info.m_cURI[0] != '\0')
							PrintURLLink(pdevobj, poempdev, MakeAnsiString(&pDevMode->info.m_cURI[0]).c_str(), rectTarget);
						break;
				}
			}
		}
	}

    //
    // turn around to call PS
    //

    return (((PFN_DrvSendPage)(poempdev->pfnPS[UD_DrvSendPage]))(pso));

}

/**
	@brief This function is called when a print job has started
	@param pso Pointer to the surface object representing the writing PostScript file
	@param pwszDocName Name of the document printed
	@param dwJobId Print job ID
	@return TRUE if started successfully, FALSE if failed
*/
BOOL APIENTRY OEMStartDoc(SURFOBJ* pso, PWSTR pwszDocName,DWORD dwJobId)
{
    PDEVOBJ     pdevobj;
    POEMPDEV    poempdev;

    VERBOSE(DLLTEXT("OEMStartDoc() entry.\r\n"));

    pdevobj = (PDEVOBJ)pso->dhpdev;
    poempdev = (POEMPDEV)pdevobj->pdevOEM;

    //
    // turn around to call PS
    //

    if (!(((PFN_DrvStartDoc)(poempdev->pfnPS[UD_DrvStartDoc])) (
            pso,
            pwszDocName,
            dwJobId)))
		return FALSE;

	POEMDEV pDevMode = (POEMDEV)pdevobj->pOEMDM;
	std::string sFilename = MakeAnsiString(pDevMode->cFilename);

	poempdev->pLinks = NULL;
	if (poempdev->pTranslator == NULL)
		poempdev->pTranslator = new GlyphTranslator;
	poempdev->bNeedText = pDevMode->bAutoURLs ? true : false;

	// Check registry for data file for this print job
	poempdev->bUsedPrintData = false;
	if (poempdev->bLoadedData = poempdev->dataLinks.LoadProcessData(pdevobj->hPrinter))
	{
		VERBOSE(DLLTEXT("Found process data\r\n"));
		pDevMode->bAutoURLs = false;
		if (poempdev->dataLinks.IsTestPage())
		{
			if (poempdev->dataLinks.GetPageCount() != 1)
				poempdev->dataLinks.SetTestPage(false);
			else
			{
				poempdev->bNeedText = true;
				sFilename = ":dropfile:";
			}
		}
	}

	// Do we know the filename we will use?
	if (!sFilename.empty())
	{
		// Yeh, so put it in the file (the converter app will get it from there)
		sFilename = "%%File: " + sFilename + "\r\n";
		DWORD dwResult;
		std::tstring::size_type dwLen = sFilename.size();

		poempdev->pOEMHelp->DrvWriteSpoolBuf(pdevobj, (LPVOID)sFilename.c_str(), (DWORD)dwLen, &dwResult);
		if (dwResult != dwLen)
			return FALSE;

		// Also write the auto open command, if we should:
		if (pDevMode->bAutoOpen)
		{
			if (pDevMode->bCreateAsTemp)
				sFilename = "%%CreateAsTemp\r\n";
			else
				sFilename = "%%FileAutoOpen\r\n";
			dwLen = (DWORD) sFilename.size();
			poempdev->pOEMHelp->DrvWriteSpoolBuf(pdevobj, (LPVOID)sFilename.c_str(), (DWORD)dwLen, &dwResult);
			if (dwResult != dwLen)
				return FALSE;

		}
	}

	return TRUE;
}

/**
	@brief This function is called the printing has ended
	@param pso Pointer to the surface object representing the writing PostScript file
	@param fl Status of the print job
	@return TRUE if written successfully, FALSE if failed to write
*/
BOOL APIENTRY OEMEndDoc(SURFOBJ* pso, FLONG fl)
{
    PDEVOBJ     pdevobj;
    POEMPDEV    poempdev;

    VERBOSE(DLLTEXT("OEMEndDoc() entry.\r\n"));

	// Prepare stuff
    pdevobj = (PDEVOBJ)pso->dhpdev;
    poempdev = (POEMPDEV)pdevobj->pdevOEM;
	POEMDEV pDevMode = (POEMDEV)pdevobj->pOEMDM;

	// Clean up the translator
	if (poempdev->pTranslator != NULL)
	{
		delete poempdev->pTranslator;
		poempdev->pTranslator = NULL;
	}
	// Clean up the link data file (only if actually printed: the printer driver is called for setting up stuff before the actual printing)
	if (poempdev->bUsedPrintData)
	{
		poempdev->dataLinks.CleanSaved(pdevobj->hPrinter);
	}

	if (fl != ED_ABORTDOC)
	{
		switch (pDevMode->info.m_eLicense)
		{
			case LicenseInfo::LTCC:
			case LicenseInfo::LTSampling:
			case LicenseInfo::LTDevelopingNations:
				// Add another page:
				if (!OEMStartPage(pso))
					return FALSE;
				if (!DoLicensePage(pso))
					return FALSE;
				break;
		}
	}

    //
    // turn around to call PS
    //

    return (((PFN_DrvEndDoc)(poempdev->pfnPS[UD_DrvEndDoc])) (
            pso,
            fl));

}

/**
	@brief This function is called when an escape code is sent to the printer
	@param pso Pointer to the surface object representing the writing PostScript file
	@param iEsc The escape code
	@param cjIn Amount of data in the escape sequence
	@param pvIn Pointer to the escape sequence data
	@param cjOut Size of output buffer
	@param pvOut Output buffer
	@return Depends on the escape sequence (0 if not supported)

    This printer supports sending a Link escape sequense as detailed in oemps.h !
*/
ULONG APIENTRY OEMEscape(SURFOBJ* pso, ULONG iEsc, ULONG cjIn, PVOID pvIn, ULONG cjOut, PVOID pvOut)
{
    PDEVOBJ     pdevobj;
    POEMPDEV    poempdev;

	// Initialize stuff
	pdevobj = (PDEVOBJ)pso->dhpdev;
	poempdev = (POEMPDEV)pdevobj->pdevOEM;

	// Is this a code-support query?
	if ((iEsc == QUERYESCSUPPORT) && (cjIn == 4))
	{
		// Yeah, check the code
		DWORD* pQuery = (DWORD*)pvIn;
		switch (*pQuery)
		{
			case ESCAPE_LINK_DATA:
			case ESCAPE_DISABLE_AUTO_URL:
				// We support those (see oemps.h)
				return TRUE;
			default:
				break;
		}
	}
	else
	{
#ifdef _DEBUG
		VERBOSE(DLLTEXT("OEMEscape(%d, %d, %p) entry.\r\n"), iEsc, cjIn, pvIn);
		if (cjIn == 4)
		{
			VERBOSE(DLLTEXT("... Data: %p\r\n"), *((DWORD*)pvIn));
		}
#endif

		// Check out what code is this escape
		switch (iEsc)
		{
			case ESCAPE_LINK_DATA:
				// A link escape
				VERBOSE(DLLTEXT("OEMEscape: size(%d,%d), logpixel(%d)\r\n"), pso->sizlBitmap.cx, pso->sizlBitmap.cy, pdevobj->pPublicDM->dmLogPixels);
				if (cjIn < sizeof(EscapeLinkData))
					return FALSE;
				// Do something:
				poempdev->pLinks = new InnerEscapeLinkData((const char*)pvIn, cjIn, poempdev->pLinks);
				return TRUE;
			case ESCAPE_DISABLE_AUTO_URL:
				// A disable-auto-URL-linking escape
				poempdev->bNeedText = false;
				return TRUE;
		}
	}

	// Call the base driver to handle this
	return (((PFN_DrvEscape)(poempdev->pfnPS[UD_DrvEscape])) (pso, iEsc, cjIn, pvIn, cjOut, pvOut));
}


/**
	@brief This function is called when a text is writen to the printer
	@param pso Pointer to the surface object representing the writing PostScript file
	@param pstro Pointer to a printing glyph list and location for each
	@param pfo Pointer to the font structure
	@param pco Pointer to the clipping information structure
	@param prclExtra Not used (always NULL)
	@param prclOpaque Opaque rectangle location
	@param pboFore Foreground brush object
	@param pboOpaque Opaque brush object
	@param pptlOrg Brush origin
	@param mix The brush raster operation
	@return TRUE if successful, FALSE if failed in any way
*/
BOOL APIENTRY OEMTextOut(SURFOBJ *pso, STROBJ *pstro, FONTOBJ *pfo, CLIPOBJ *pco, RECTL *prclExtra, RECTL *prclOpaque, BRUSHOBJ *pboFore, BRUSHOBJ *pboOpaque, POINTL *pptlOrg, MIX mix)
{
    PDEVOBJ     pdevobj;
    POEMPDEV    poempdev;
    VERBOSE(DLLTEXT("OEMTextOut\r\n"));

	// Initialize stuff
	pdevobj = (PDEVOBJ)pso->dhpdev;
	poempdev = (POEMPDEV)pdevobj->pdevOEM;

	// Do we need to save the text location for later?
	if (poempdev->bNeedText && ((pstro->cGlyphs > 0) && (pstro->pwszOrg != NULL)))
	{
		// Yes...
		PGLYPHPOS pGlyphPos;
		POINTQF* pWidths = NULL;
		bool bValid = false;

		// Is this a fixed-font?
		if (pstro->ulCharInc == 0)
		{
			// No, so get glyph locations
			ULONG uCount;
			BOOL bRet = STROBJ_bEnumPositionsOnly(pstro, &uCount, &pGlyphPos);

			if ((bRet != DDI_ERROR) && (uCount == pstro->cGlyphs))
			{
				pWidths = new POINTQF[uCount];
				if (STROBJ_bGetAdvanceWidths(pstro, 0, uCount, pWidths))
					bValid = true;
			}
		}
		else
			bValid = true;

		if (bValid)
		{
			// OK, we can go on and read the actual text now
			std::wstring sText;
			if ((pstro->flAccel & SO_GLYPHINDEX_TEXTOUT) != 0)
			{
				// Those are glyphs, not an actual string, so we need to translate them:
				const GlyphToText* pGlyphMap = NULL;
				PIFIMETRICS pifi;

				TRACE(DLLTEXT("Found glyph text:\r\n"));

				// Get font data:
				pifi = FONTOBJ_pifi(pfo);
				if (pifi != NULL)
				{
					LOGFONT lfFont;
					// Calculate font point size
					XFORMOBJ* pFormObj = FONTOBJ_pxoGetXform(pfo);
					XFORML xForm;
					ULONG uXForm = XFORMOBJ_iGetXform(pFormObj, &xForm);

					double dXScale = sqrt(xForm.eM11 * xForm.eM11 + xForm.eM12 * xForm.eM12);
					double dYScale = sqrt(xForm.eM22 * xForm.eM22 + xForm.eM21 * xForm.eM21);

					// Populate font object for translation
					lfFont.lfHeight = (int)(0.5 + dYScale * pifi->fwdUnitsPerEm * 72) / pfo->sizLogResPpi.cy;
					lfFont.lfWidth = 0;
					lfFont.lfEscapement = 0;
					lfFont.lfOrientation = 0;
					lfFont.lfWeight = pifi->usWinWeight;
					lfFont.lfItalic = (pifi->fsSelection & FM_SEL_ITALIC) ? TRUE : FALSE;
					lfFont.lfUnderline = (pifi->fsSelection & FM_SEL_UNDERSCORE) ? TRUE : FALSE;
					lfFont.lfStrikeOut = (pifi->fsSelection & FM_SEL_STRIKEOUT) ? TRUE : FALSE;
					lfFont.lfCharSet = pifi->jWinCharSet;
					lfFont.lfOutPrecision = OUT_DEFAULT_PRECIS;
					lfFont.lfClipPrecision = CLIP_DEFAULT_PRECIS;
					lfFont.lfQuality = DEFAULT_QUALITY;
					lfFont.lfPitchAndFamily = pifi->jWinPitchAndFamily;
					wcsncpy_s(lfFont.lfFaceName, _S(lfFont.lfFaceName), (TCHAR*)(((char*)pifi) + (DWORD)pifi->dpwszFamilyName), LF_FACESIZE);

					// Get the glyph map
					pGlyphMap = (poempdev->pTranslator == NULL) ? NULL : poempdev->pTranslator->GetFontTranslation(lfFont);
					if (pGlyphMap != NULL)
					{
						// Found it
						GlyphToText::const_iterator iChar;
						for (unsigned int i=0;i<pstro->cGlyphs;i++)
						{
							// Map each glyph into a character
							iChar = pGlyphMap->find(pstro->pwszOrg[i]);
							if (iChar != pGlyphMap->end())
								sText += (*iChar).second;
							else
								sText += (WCHAR)0x7F;
						}
						TRACE(DLLTEXT("%s\r\n"), sText.c_str());
					}
				}
				if (sText.empty())
				{
					TRACE(DLLTEXT("Could not unglyph it...\r\n"));
				}
			}
			else
			{
				// This is just a string, use it
				sText.assign(pstro->pwszOrg, pstro->cGlyphs);
				TRACE(DLLTEXT("Got the following text:\r\n"));
			}

			if (!sText.empty())
			{
				// OK we have data
				TRACE(DLLTEXT("%s [at %d,%d-%d,%d]\r\n"), sText.c_str(), pstro->rclBkGround.left, pstro->rclBkGround.top, pstro->rclBkGround.right, pstro->rclBkGround.bottom);
		
				// Add to the parts
				if (pstro->ulCharInc == 0)
				{
					// Use variable locations
					ASSERT(pWidths != NULL);
					poempdev->oText.AddLine(TextLine(sText, pstro->rclBkGround, pGlyphPos, pWidths));
					delete [] pWidths;
				}
				else
				{
					// Fixed font
					ASSERT(pWidths == NULL);
					poempdev->oText.AddLine(TextLine(sText, pstro->rclBkGround, pstro->ulCharInc));
				}
			}
		}
	}

	// Call base driver to do the actual printing work...
	return (((PFN_DrvTextOut)(poempdev->pfnPS[UD_DrvTextOut])) (pso, pstro, pfo, pco, prclExtra, prclOpaque, pboFore, pboOpaque, pptlOrg, mix));
}
