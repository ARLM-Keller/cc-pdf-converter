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

#ifndef _GLYPHTRANSLATOR_H_
#define _GLYPHTRANSLATOR_H_

#include <map>
#include "CCTChar.h"

/**
    @brief This class holds glyph-to-Unicode-character data for a specific font
*/
struct GlyphToText : public std::map<WORD, WCHAR>
{
public:
	/**
		@brief Default constructor
	*/
	GlyphToText() {};
	/// Adds the font data to the object
	bool	Initialize(const LOGFONT& lf, HDC hDC);
};

/**
    @brief This class retrieves a glyph-to-Unicode map for a font (and caches the map)
*/
class GlyphTranslator : protected std::map<std::tstring, GlyphToText>
{
public:
	/// Get a translation map for a font
	const GlyphToText* GetFontTranslation(const LOGFONT& lf, HDC hDC = NULL);
};

#endif   //#define _GLYPHTRANSLATOR_H_
