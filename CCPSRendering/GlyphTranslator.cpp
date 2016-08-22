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

/**
	@param lf Font description
	@param hDC Handle to the DC to use for translation
	@return true if the map was populated successfully, false if failed
*/
bool GlyphToText::Initialize(const LOGFONT& lf, HDC hDC)
{
	// Get a matching font from Windows
	HFONT hFont = ::CreateFontIndirect(&lf);
	if (hFont == NULL)
		// Dah! Not found!
		return false;

	// Get range of characters in the font
	HGDIOBJ hOldFont = ::SelectObject(hDC, hFont);
	DWORD dwSize = ::GetFontUnicodeRanges(hDC, NULL);
	if (dwSize == 0)
	{
		// None - nothing to add
		::SelectObject(hDC, hOldFont);
		return true;
	}

	// Create a translation set and get it
	LPGLYPHSET pSet = (LPGLYPHSET)new char[dwSize];
	::GetFontUnicodeRanges(hDC, pSet);

	// Go over the set:
	WORD dwGlyphs[1];
	TCHAR c[1];
	int nCount = 0;
	for (UINT i=0;i<pSet->cRanges;i++)
	{
		// For each range
		for (UINT u=0;u<pSet->ranges[i].cGlyphs;u++)
		{
			// Put the glyph in a buffer
			c[0] = pSet->ranges[i].wcLow + u;
			// And retrieve the character for it
			if (::GetGlyphIndices(hDC, c, 1, (LPWORD)&dwGlyphs, GGI_MARK_NONEXISTING_GLYPHS) != 1)
			{
				// Cannot get it, fail
				delete [] pSet;
				::SelectObject(hDC, hOldFont);
				return false;
			}
			// OK, replace it if not already there (there can be two glyphs for the same character)
			if (find(dwGlyphs[0]) == end())
				operator[](dwGlyphs[0]) = c[0];
		}
	}

	// OK, this is it
	delete [] pSet;
	::SelectObject(hDC, hOldFont);
	return true;
}

/**
	@param lf Font description
	@param hDC Handle to the DC to use for translation (if NULL, uses the default screen DC)
	@return A pointer to the translation map (NULL if cannot traslate this font)
*/
const GlyphToText* GlyphTranslator::GetFontTranslation(const LOGFONT& lf, HDC hDC /* = NULL */)
{
	// Check out if we have this font cached
	TCHAR cFontID[LF_FACESIZE + 128];
	_stprintf_s(cFontID, _S(cFontID), _T("%.*s|%d|%d|%d|%d|%d"), LF_FACESIZE, lf.lfFaceName, lf.lfHeight, lf.lfWeight, (lf.lfItalic ? 1 : 0) | (lf.lfUnderline ? 2 : 0) | (lf.lfStrikeOut ? 4 : 0), lf.lfPitchAndFamily);
	iterator iGlyphData;

	if ((iGlyphData = find(cFontID)) == end())
	{
		// Nope, so add an empty translation map
		insert(std::pair<std::tstring, GlyphToText>(cFontID, GlyphToText()));
		iGlyphData = find(cFontID);
		_ASSERT(iGlyphData != end());
		// And initialize it for the font
		HDC hUseDC = (hDC == NULL) ? GetDC(NULL) : hDC;
		if (hUseDC == NULL)
			return NULL;
		bool bRet = !(*iGlyphData).second.Initialize(lf, hUseDC);
		if (hUseDC != hDC)
			DeleteDC(hUseDC);
		if (!bRet)
		{
			// Cannot, so fail
			erase(iGlyphData);
			return NULL;
		}
	}

	// OK, found it
	return &((*iGlyphData).second);
}
