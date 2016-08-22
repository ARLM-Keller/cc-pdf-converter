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

#ifndef _FILEINI_H_
#define _FILEINI_H_

// Includes
#include <string>
#include <list>
#include <map>
#include "CCTChar.h"

/// A list of strings
typedef std::list<std::tstring> TCHARSTRLIST;
/// A map of string to string.
typedef std::map<std::tstring, std::tstring> TCHARSTR2STR;
/// A map of string, to string-list
typedef std::map<std::tstring, TCHARSTRLIST> TCHARSTR2STRLIST;

/** \addtogroup FileINIFlags FileINI Flags
 * Flags for GetLines() or GetKeys()
 * @{
 */
/// Skips over double-slash (//) comment lines in the INI file
#define		FILEINI_IGNORE_COMMENTS		0x00000001
/// Trims the keys and values found in the INI file
#define		FILEINI_TRIM				0x00000002
/**
 * @}
 */

/**
    @brief This class represents a line with a key-value pair of strings, inside an INI file
*/
struct INILine
{
	/**
		@brief Default constructor
	*/
	INILine() 
	{
	};

	/**
		@brief Standard constructor
		@param Key The key string of the new object
		@param Value The value string of the new object
	*/
	INILine(const std::tstring& Key, const std::tstring& Value) 
	{
		// Initialize members
		sKey = Key; 
		sValue = Value;
	};
	/**
		@brief Copy constructor
		@param line An existing INILine object to copy from
	*/
	INILine(const INILine& line) 
	{
		sKey = line.sKey; 
		sValue = line.sValue;
	};
	/// The line's key string
	std::tstring sKey;
	/// The line's value string
	std::tstring sValue;
};

/// A list of INILine objects, representing a group of INI file lines, in their original order
typedef std::list<INILine> INILINELIST;

/**
    @brief This class, which is composed completely of static methods, can be used to read and parse INI files
*/
class FileINI
{
public:
	/**
		@brief Default constructor
	*/
	FileINI() : m_pData(NULL), m_bOwn(false), m_bUnicode(false), m_nOffset(0) {};
	/// Load INI data from file
	FileINI(LPCTSTR lpFilename);
	/// Load INI data from memory buffer
	FileINI(char* pData, DWORD dwSize, bool bOwn = true);
	/// Load memory data from buffer (unicode flag)
	FileINI(bool bUnicode, char* pData, DWORD dwSize, bool bOwn = true);
	/**
		@brief Destructor
	*/
	~FileINI() {if (m_bOwn && (m_pData != NULL)) delete [] m_pData;};

protected:
	// Members
	/// Buffer data
	char*		m_pData;
	/// Buffer ownership flag (true means the buffer should be deleted with the object)
	bool		m_bOwn;
	/// true if this is a unicode buffer
	bool		m_bUnicode;
	/// Current offset in the buffer
	int			m_nOffset;

public:
	// Data Access
	/**
		@brief Checks if there's any data in the buffer
		@return true if there's a loaded buffer, false if there's none
	*/
	bool		HasData() const {return m_pData != NULL;};

	// Non-static methods
	/// Loads a file
	bool	LoadINIFile(LPCTSTR lpFilename);
	/// Retrieves all lines in the specified section of the INI file
	bool	GetLines(LPCTSTR pSection, TCHARSTRLIST& listLines, UINT uFlags = 0);
	/// Retrieves all key names and values in all of the lines in the specified section of the INI file
	bool	GetKeys(LPCTSTR pSection, TCHARSTR2STR& mapKeys, UINT uFlags = 0);
	/// Retrieves all key names and values in all of the lines in the specified section of the INI file
	bool	GetKeys(LPCTSTR pSection, INILINELIST& listLines, UINT uFlags = 0);
	/// Retrieves the list of sections inside the INI file
	bool	GetAllSections(TCHARSTRLIST& lSections);

	// Static Methods
	/// Retrieves all lines in the specified section of the INI file
	static bool	GetLines(LPCTSTR pFilename, LPCTSTR pSection, TCHARSTRLIST& list, UINT uFlags = 0);
	/// Retrieves all key names and values in all of the lines in the specified section of the INI file
	static bool	GetKeys(LPCTSTR pFilename, LPCTSTR pSection, TCHARSTR2STR& mapKeys, UINT uFlags = 0);
	/// Retrieves all key names and values in all of the lines in the specified section of the INI file
	static bool	GetKeys(LPCTSTR pFilename, LPCTSTR pSection, INILINELIST& list, UINT uFlags = 0);

	/// Parses a comma separated list of values, and returns their parsed list. Supports quotes.
	static bool GetVarValues(LPCTSTR pVarValue, TCHARSTRLIST& lValues);
	/// Performs previously defined string replacements on a string
	static bool ReplaceVariables(LPCTSTR pValue, const TCHARSTR2STRLIST& lVariables, TCHARSTRLIST& lResults);
	/// Retrieves the list of sections inside the INI file
	static bool GetAllSections(LPCTSTR pFilename, TCHARSTRLIST& lSections);

protected:
	// Helpers
	/// Check if the line is the specified section header
	static bool	CheckLineForSection(const WCHAR* pLine, std::tstring::size_type nLineLen, LPCTSTR pSection, std::tstring::size_type nSectionLen, bool bUnicode, bool& bFoundSection);
	/// Parses a (non-section) line and adds the contents into the values pair list
	static void DoInitLine(const WCHAR* pLine, std::tstring::size_type nLineLen, INILINELIST& list, UINT uFlags, bool bUnicode);
	/// Check to see if the file is a unicode INI file; tested at the opened location (doesn't seek to the start!)
	static bool TestUnicode(FILE* pFile);

	/// Check to see if file INI file is in unicode
	bool	TestUnicode();
};

#endif   //#define _FILEINI_H_
