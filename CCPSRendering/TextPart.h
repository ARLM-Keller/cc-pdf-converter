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

#ifndef _TEXTPART_H_
#define _TEXTPART_H_

#include <string>
#include <list>
#include "CCTChar.h"

/// Definition: string list
typedef std::list<std::tstring> STRLIST;

/**
    @brief Helper object for location of a printed letter
*/
struct TextLetter
{
	// Ctors
	/**
		@brief Default constructor
	*/
	TextLetter() : wLetter(0), x(0), nWidth(0) {};
	/**
		@brief Constructor
		@param w The letter
		@param pos The start location
		@param width The width
	*/
	TextLetter(WCHAR w, long _x, int width) : wLetter(w), x(_x), nWidth(width) {};
	/**
		@brief Copy constructor
		@param other The letter data object to copy
	*/
	TextLetter(const TextLetter& other) : wLetter(other.wLetter), x(other.x), nWidth(other.nWidth) {};

	// Members
	/// The letter
	WCHAR						wLetter;
	/// The x-location
	long						x;
	/// The width
	int							nWidth;
};

/**
    @brief Helper object for printed word location
*/
struct TextWord : public std::list<TextLetter>
{
	// Ctors
	/**
		@brief Default constructor
	*/
	TextWord() {};
	/**
		@brief Copy constructor
		@param other Printer work location object to copy
	*/
	TextWord(const TextWord& other) {assign(other.begin(), other.end());};
	/// Constructor: create from a variable-width font string
	TextWord(const std::wstring& s, const PGLYPHPOS& arGlyphPos, const POINTQF* pWidths, std::tstring::size_type nStart, std::tstring::size_type nEnd = -1);
	/// Constructor: create from a fixed-width font string
	TextWord(const std::wstring& s, int nCharWidth, std::tstring::size_type nStart, std::tstring::size_type nEnd = -1);

	// Data Access methods
	/**
		@brief Returns the left border of the requested letter
		@param nLetter The letter to get the location of
		@return Location of the requested letter
	*/
	long GetStart(size_type nLetter) const 
	{
		if ((nLetter < 0) || (nLetter >= size())) 
			return 0; 
		const_iterator i; 
		for (i = begin(); nLetter > 0; nLetter--, i++) ; 
		return (*i).x;
	};
	/**
		@brief Returns the right border of the requested letter
		@param nLetter The letter to get the location of
		@return Location of the requested letter
	*/
	long GetEnd(size_type nLetter) const 
	{
		if ((nLetter < 0) || (nLetter >= (int)size())) 
			return 0; 
		const_iterator i; 
		for (i = begin(); nLetter > 0; nLetter--, i++) ; 
		return (*i).x + (*i).nWidth;
	};
	/**
		@brief Returns the word's contents
		@return The text
	*/
	std::wstring GetText() const {std::wstring s; for (const_iterator i = begin(); i != end(); i++) s += (*i).wLetter; return s;};

	// Operators
	/**
		@brief Adds another word's letters after the current letters
		@param other The word to add
		@return This object
	*/
	const TextWord& operator+=(const TextWord& other) {insert(end(), other.begin(), other.end()); return *this;};
};

/**
    @brief Helper object for printed text line
*/
struct TextLine : public std::list<TextWord>
{
	// Ctors
	/**
		@brief Default constructor
	*/
	TextLine() {rcArea.left = 0; rcArea.top = 0; rcArea.right = 0; rcArea.bottom = 0;};
	/**
		@brief Copy constructor
		@param other The text line object to copy
	*/
	TextLine(const TextLine& other) {rcArea = other.rcArea; assign(other.begin(), other.end());};
	/// Constructor: from a string of variable-width font
	TextLine(const std::wstring& s, const RECTL& rc, const PGLYPHPOS& arGlyphPos, const POINTQF* pWidths);
	/// Constructor: from a string of fixed-width font
	TextLine(const std::wstring& s, const RECTL& rc, int nCharWidth);

	// Members
	/// The line's area
	RECTL rcArea;

	/// This function checks if the received line of text is on the same line as this one, and if so, adds its data to this line
	bool AddTextSameLine(const TextLine& other);

protected:
	/// Updates the line's rectangle using the word locations
	void SetSides();
};

/**
    @brief Page text helper object, used to search for specific strings to find their location
*/
struct TextArea : public std::list<TextLine>
{
protected:
	// Members (for forward-only search)
	/// Current search line
	const_iterator m_iLine;
	/// Current search word
	TextLine::const_iterator m_iWord;

public:
	// Data Access
	/// Add a new text line to the object
	void	AddLine(const TextLine& line);

	// Methods
	/// Start a new search
	void	InitSearch();
	/**
		@brief Search for an expression (a list of words in the exact order) in the page text; must be on the same line
		@param words The words to search for
		@param[out] rectArea The page location in which the words were found
		@param nRepeat The amount of times to jump over the expression before reporting success
		@return true if the expression was found, false if not

		Use the nRepeat to jump over the first nRepeat times the expression is found.
		Note that this is a forward only search, so after the first false result, you have to re-initialize the search to start from the beginning
	*/
	bool	SearchFor(const STRLIST& words, RECTL& rectArea, int nRepeat) {do {if (!SearchFor(words, rectArea)) return false; nRepeat--;} while (nRepeat > 0); return true;};
	/// Search for an expression (a list of words in the exact order) in the page text; must be on the same line
	bool	SearchFor(const STRLIST& words, RECTL& rectArea);
	/// Search for the next string that starts with http:// or https:// and return its location
	bool	SearchForURL(RECTL& rectArea, std::wstring& sWord);
};

#endif   //#define _TEXTPART_H_
