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
#include "TextPart.h"

/**
	@param s The sentence's text
	@param arGlyphPos The location of each letter
	@param pWidths The widths of each letter
	@param nStart The first letter to use
	@param nEnd The last letter to use
*/
TextWord::TextWord(const std::wstring& s, const PGLYPHPOS& arGlyphPos, const POINTQF* pWidths, std::tstring::size_type nStart, std::tstring::size_type nEnd /* = -1 */)
{
	// Calculate the end of the word if not specified
	if (nEnd == -1)
		nEnd = s.size();
	// Go over the data
	for (; nStart < nEnd; nStart++)
	{
		// Add a letter object from each letter and position
		ASSERT(s[nStart] != ' ');
		push_back(TextLetter(s[nStart], arGlyphPos[nStart].ptl.x, pWidths[nStart].x.HighPart >> 4));
	}
}

/**
	@param s The sentence's text
	@param nCharWidth The width of each glyph
	@param nStart The first letter to use
	@param nEnd The last letter to use
*/
TextWord::TextWord(const std::wstring& s, int nCharWidth, std::tstring::size_type nStart, std::tstring::size_type nEnd /* = -1 */)
{
	// Calculate the end of the word if not specified
	if (nEnd == -1)
		nEnd = s.size();
	// Go over the data
	for (; nStart < nEnd; nStart++)
	{
		// Add a letter object from each letter and position
		ASSERT(s[nStart] != ' ');
		push_back(TextLetter(s[nStart], (long) (nStart * nCharWidth), nCharWidth));
	}
}



/**
	@param s The line's text
	@param rc The line's printed location
	@param arGlyphPos Array of glyph locations
	@param pWidths Array of glyph widths
*/
TextLine::TextLine(const std::wstring& s, const RECTL& rc, const PGLYPHPOS& arGlyphPos, const POINTQF* pWidths) : rcArea(rc)
{
	// Go over the text, break on spaces and such
	size_t nPos = 0, nEndPos = s.find_first_of(_T(" \r\n\t"));
	while (nEndPos != std::wstring::npos)
	{
		// Do we have something here?
		if (nEndPos > nPos)
			// Yes, add the word
			push_back(TextWord(s, arGlyphPos, pWidths, nPos, nEndPos));

		// Move on
		nPos = nEndPos + 1;
		if (nPos < s.size())
			// Find end of space
			nEndPos = s.find_first_of(_T(" \r\n\t"), nPos);
		else
			// Ah, finished
			nEndPos = std::wstring::npos;
	}
	if (nPos < s.size())
		// Add the last word
		push_back(TextWord(s, arGlyphPos, pWidths, nPos));

	// Update the line rectangle
	SetSides();
}

/**
	@param s The line's text
	@param rc The line's printed location
	@param nCharWidth The width of each glyph (fixed font)
*/
TextLine::TextLine(const std::wstring& s, const RECTL& rc, int nCharWidth) : rcArea(rc)
{
	// Go over the text, break on spaces and such
	std::tstring::size_type nPos = 0, nEndPos = s.find_first_of(_T(" \r\n\t"));
	while (nEndPos != std::wstring::npos)
	{
		// Do we have something here?
		if (nEndPos > nPos)
			// Yes, add the word
			push_back(TextWord(s, nCharWidth, nPos, nEndPos));

		// Move on
		nPos = nEndPos + 1;
		if (nPos < s.size())
			// Find end of space
			nEndPos = s.find_first_of(_T(" \r\n\t"), nPos);
		else
			// Ah, finished
			nEndPos = std::wstring::npos;
	}
	if (nPos < s.size())
		// Add the last word
		push_back(TextWord(s, nCharWidth, nPos));

	// Update the line rectangle
	SetSides();
}

/**
	@brief Checks if two rectangles are on the same line
	@param rect1 First line to check
	@param rect2 Second line to check
	@return true if the rectangles are more-or-less on the same line, false if not
*/
bool OnSameLine(const RECTL& rect1, const RECTL& rect2)
{
	int nMiddle1 = (rect1.top + rect1.bottom) / 2, nMiddle2 = (rect2.top + rect2.bottom) / 2;
	return ((nMiddle1 >= rect2.top) && (nMiddle1 <= rect2.bottom)) || ((nMiddle2 >= rect1.top) && (nMiddle2 <= rect1.bottom));
}

/**
	@param other The text line object to check
	@return true if the other line object's data was added to this one, false if not
*/
bool TextLine::AddTextSameLine(const TextLine& other)
{
	// Are we on the same line?
	if (!OnSameLine(rcArea, other.rcArea))
		// No, just leave it
		return false;

	// Same line: before or after?
	int nOffset;
	if (rcArea.left > other.rcArea.left)
	{
		nOffset = rcArea.left - other.rcArea.left;
		if (rcArea.left > other.rcArea.right + 2)
		{
			/* <other> <me> */
			insert(begin(), other.begin(), other.end());
		}
		else
		{
			/* <other><me> */
			ASSERT(!other.empty());
			ASSERT(!empty());
			TextWord part(other.back());
			part += front();
			erase(begin());
			push_front(part);
			if (other.size() > 1)
			{
				const_iterator ci = other.end();
				ci--;
				insert(begin(), other.begin(), ci);
			}
		}
		rcArea.left = other.rcArea.left;
	}
	else
	{
		nOffset = other.rcArea.left - rcArea.left;
		std::tstring::size_type nSize = size();
		if (rcArea.right < other.rcArea.left - 2)
		{
			/* <me> <other> */
			insert(end(), other.begin(), other.end());
		}
		else
		{
			/* <me><other> */
			ASSERT(!other.empty());
			ASSERT(!empty());
			const_iterator i = other.begin();
			back() += (*i);
			i++;
			if (i != other.end())
				insert(end(), i, other.end());
		}
		rcArea.right = other.rcArea.right;
	}

	// Combine the areas
	rcArea.top = min(rcArea.top, other.rcArea.top);
	rcArea.bottom = max(rcArea.bottom, other.rcArea.bottom);
	return true;
}

/**
	
*/
void TextLine::SetSides() 
{
	// Do we have any data?
	if (empty()) 
		// No, make it VERY small :>
		rcArea.right = rcArea.left; 
	else 
	{
		// Get the range from the actual words (it could be larger)
		const TextWord& word = back(); 
		rcArea.right = word.GetEnd(word.size() - 1);
		rcArea.left = front().GetStart(0);
	}
}





/**
	@param line The line to add
*/
void TextArea::AddLine(const TextLine& line)
{
	// Don't put in empty lines
	if (line.empty())
		return;

	if (empty())
		// Just put it in
		push_back(line);
	else
	{
		// Check against the previous line
		TextLine& last = back();
		if (!last.AddTextSameLine(line))
			// OK, it's not on the same line, so add it
			push_back(line);
	}
}

/**
	
*/
void TextArea::InitSearch()
{
	// Initialize the search location to the first line
	m_iLine = begin();
	if (m_iLine != end())
		// And the first word in the line
		m_iWord = (*m_iLine).begin();
}

/**
	@param words The exression to search for
	@param[put] rectArea The page location of the expression
	@return true if found the expression, false if failed
*/
bool TextArea::SearchFor(const STRLIST& words, RECTL& rectArea)
{
	// Is there something to find?
	if (words.empty())
		// Nope
		return false;

	// Start searching
	const_iterator iPrev = m_iLine;
	std::wstring sWord;
	std::wstring::size_type pos;

	for (; m_iLine != end(); m_iLine++)
	{
		// Did we move to a new line?
		if (iPrev != m_iLine)
		{
			// Yeah, does it have enough words to cover the expression?
			if ((*m_iLine).size() < words.size())
				// No, go to the next line
				continue;
			// Initialize the search to the first word in the line
			m_iWord = (*m_iLine).begin();
		}

		// Go over the words in the current line
		for (; m_iWord != (*m_iLine).end(); m_iWord++)
		{
			// Get the text
			sWord = (*m_iWord).GetText();
			if ((pos = sWord.find(words.front().c_str())) != (sWord.size() - words.front().size()))
				// Can't be this word: didn't find the first expression's word as the end of this word
				// This is for matching:
				// 'and' as the end of 'wand'!
				continue;

			// Initialize the search for the rest of the expression
			TextLine::const_iterator iTestThis = m_iWord;
			STRLIST::const_iterator iTestWords = words.begin(), iTestWordsNext;
			iTestWords++;
			while (iTestWords != words.end())
			{
				// Get next word
				iTestThis++;
				if (iTestThis == (*m_iLine).end())
					// Nothing in this line, so not found
					break;
				sWord = (*iTestThis).GetText();

				iTestWordsNext = iTestWords;
				iTestWordsNext++;
				if (iTestWordsNext == words.end())
				{
					// Check partial word at the end
					if (sWord.find((*iTestWords).c_str()) != 0)
						break;
					// Found!
					iTestWords = words.end();
				}
				else
				{
					if (sWord != (*iTestWords))
						// Not it!
						break;
				}
				// Go on
				iTestWords = iTestWordsNext;
			}

			if (iTestWords == words.end())
			{
				// Found it
				rectArea = (*m_iLine).rcArea;
				rectArea.left = (*m_iWord).GetStart(pos);
				rectArea.right = (*iTestThis).GetEnd((*iTestThis).size() - 1);
				m_iWord = iTestThis;
				m_iWord++;
				return true;
			}
		}			
		iPrev = m_iLine;
	}
	return false;
}

/**
	@param[out] rectArea Location of next URL
	@param[out] sURL The URL found
	@return true if a URL was found, false if none were found
*/
bool TextArea::SearchForURL(RECTL& rectArea, std::wstring& sURL)
{
	// Continue from last search
	const_iterator iPrev = m_iLine;
	std::wstring sWord;
	std::tstring::size_type nStart, nEnd;
	for (; m_iLine != end(); m_iLine++)
	{
		// We are gonna assume that a URL cannot have a space in it. Which is basically true.

		if (iPrev != m_iLine)
			// Start from the beginning of the next line...
			m_iWord = (*m_iLine).begin();

		// Go over the words
		for (; m_iWord != (*m_iLine).end(); m_iWord++)
		{
			// Get the text
			sWord = (*m_iWord).GetText();
			if (sWord.size() < 8)
				// Too short
				continue;
			if (_wcsnicmp(sWord.c_str(), _T("http"), 4) != 0)
				// No http
				continue;

			nStart = 4;
			if ((sWord[nStart] == 's') || (sWord[nStart] == 'S'))
				// https
				nStart++;
			// Does it have :// now?
			if (wcsncmp(sWord.c_str() + nStart, _T("://"), 3) != 0)
				// Nope
				continue;
			nStart += 3;

			// Find the end of the URL
			nEnd = sWord.find_first_not_of(_T("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.+$-_@&+-!*\"'(),%/?"), nStart);
			if (nEnd == std::wstring::npos)
				nEnd = sWord.size();
			// Remove all kinds of non-URL stuff sometimes found at the end of the URL
			while ((nEnd > nStart) && wcschr(_T(").'\"?"), sWord[nEnd-1]) != NULL)
				nEnd--;

			// Calculate position
			rectArea = (*m_iLine).rcArea;
			rectArea.right = (*m_iWord).GetEnd(nEnd);
			rectArea.left = (*m_iWord).GetStart(0);
			sURL = sWord.substr(0, nEnd);

			// Leave it pointing to the next part
			m_iWord++;
			return true;
		}
		iPrev = m_iLine;
	}
	return false;
}

