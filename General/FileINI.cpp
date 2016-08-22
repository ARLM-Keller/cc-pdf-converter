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
#include "FileIni.h"

#include <vector>
#include <tchar.h>
#include <io.h>
#include <fcntl.h>

/// Array of integers
typedef std::vector<int>		INTARRAY;
typedef std::vector<std::tstring::size_type>		SIZETARRAY;

/**
	@brief Trims a character string of any _trailing_ whitespace, tab, newline, or carriage-return
	@param pStr A character string to trim
	@param[in, out] nLen The length of the string. Returns the modified new length of the string
	@param bSpace true to drop spaces at the end too
	@param bUnicode true if this is a unicode string, false to assume the string is char
*/
void TrimLen(const WCHAR* pStr, std::tstring::size_type& nLen, bool bSpace, bool bUnicode)
{
	// Keep counting the trailing characters backwards
	if (bUnicode)
	{
		while ((nLen > 0) && (wcschr(bSpace ? L" \t\n\r" : L"\n\r", pStr[nLen-1]) != NULL))
			nLen--;
	}
	else
	{
		while ((nLen > 0) && (strchr(bSpace ? " \t\n\r" : "\n\r", ((const char*)pStr)[nLen-1]) != NULL))
			nLen--;
	}
}

/**
	@brief Returns the trimmed string, without any leading to trailing whitespace, tab, newline, or carriage-return characters
	@param pStr The string to trim
	@return The trimmed string
*/
std::tstring Trim(const WCHAR* pStr, std::tstring::size_type nLen, bool bUnicode)
{
	std::tstring sRet;
	if (bUnicode)
		sRet = MakeTString(std::wstring(pStr, nLen));
	else
		sRet = MakeTString(std::string((const char*)pStr, nLen));
	
	// Check for empty string
	if (sRet.empty())
		return sRet;

	// Trim leading characters by advancing the beginning of the string
	std::tstring::size_type pos = sRet.find_last_not_of(_T(" \t\n\r"));
	if (pos != std::tstring::npos)
	{
		pos++;
		if (pos < sRet.size())
			sRet.erase(pos);
	}

	pos = sRet.find_first_not_of(_T(" \t\n\r"));
	if (pos != std::tstring::npos)
		sRet.erase(0, pos);

	return sRet;
}

/**
	@brief Returns the trimmed string, without any leading to trailing whitespace, tab, newline, or carriage-return characters
	@param pStr The string to trim
	@return The trimmed string
*/
std::wstring Trim(const WCHAR* pStr, std::tstring::size_type nLen)
{
	// Check for empty string
	if (nLen == 0)
		return L"";
	// Trim leading characters by advancing the beginning of the string
	while ((nLen > 0) && (wcschr(L" \t\n\r", *pStr) != NULL))
	{
		pStr++;
		nLen--;
	}

	// Trim trailing characters by re-adjusting the length
	while ((nLen > 0) && (wcschr(L" \t\n\r", pStr[nLen-1]) != NULL))
		nLen--;

	// Return the trimmed string
	if (nLen == 0)
		return L"";
	else
		return std::wstring(pStr, nLen);
}

/**
	@param s String to trim
	@return Trimmed string
*/
std::tstring Trim(const std::string& s)
{
	return Trim((const WCHAR*)s.c_str(), s.size(), false);
}

/**
	@param s String to trim
	@return Trimmed string
*/
std::tstring Trim(const std::wstring& s)
{
	return Trim(s.c_str(), s.size(), true);
}

/**
	@brief Checks if a line is a comment line
	@param pStr The string line to check
	@param nLen The length of the string specified by pStr
	@param bUnicode true if this is a unicode string, false to cast it to char
	@return true if it's a comment line, false otherwise
*/
bool IsComment(const WCHAR* pStr, std::tstring::size_type nLen, bool bUnicode)
{
	// Too short lines are not comment lines
	if (nLen < 2)
		return false;
	// Test for double-slash
	return bUnicode ? ((pStr[0] == '/') && (pStr[1] == '/')) : ((((char*)pStr)[0] == (char)'/') && (((char*)pStr)[1] == (char)'/'));
}

/**
	@param pLine [in]Position to start searching from, [out]Start of next line
	@param nLen Number of characters to jump over
	@param bUnicode true if this is a Unicode text, false if it's ASCII
*/
void FindNextLine(const WCHAR*& pLine, std::tstring::size_type nLen, bool bUnicode)
{
	// Is it unicode?
	if (bUnicode)
	{
		// Yes, find the end of the line (or of the text
		pLine += nLen;
		while ((*pLine != '\0') && (wcschr(L"\r\n", *pLine) != NULL))
			pLine++;
		return;
	}

	// ASCII, so use char:
	const char* pNext = (const char*)pLine;
	pNext += nLen;
	while ((*pNext != '\0') && (strchr("\r\n", *pNext) != NULL))
		pNext++;
	// OK, convert back to the common pointer
	pLine = (const WCHAR*)pNext;
}



/**
	@param pData Pointer to a data buffer
	@param dwSize Size of the data buffer
	@param bOwn true if the buffer should be released by this object, false if not
*/
FileINI::FileINI(char* pData, DWORD dwSize, bool bOwn /* = true */) : m_bUnicode(false), m_nOffset(0)
{
	// Do we have any data?
	if (pData == NULL)
	{
		// Nope, just leave it
		m_pData = NULL;
		m_bOwn = false;
		return;
	}

	// OK, can we use it as it is?
	if (pData[dwSize-1] == '\0')
	{
		// Well, NULL-terminated we can use
		m_pData = pData;
		m_bOwn = bOwn;
	}
	else
	{
		// The data needs to be copied, as we have to get NULL-terminated string here:
		m_pData = new char[dwSize + 1];
		memcpy(m_pData, pData, dwSize);
		m_pData[dwSize] = (char)'\0';
		m_bOwn = true;
		if (bOwn)
			// Kill old data
			delete [] pData;
	}
	// Remember if it's unicode
	m_bUnicode = TestUnicode();
	if (m_bUnicode)
		m_nOffset = 2;
}

/**
	@param bUnicode true for Unicode data, false of ASCII
	@param pData The data
	@param dwSize The length of the data
	@param bOwn true if the buffer should be released by this object, false if not
*/
FileINI::FileINI(bool bUnicode, char* pData, DWORD dwSize, bool bOwn /* = true */) : m_bUnicode(bUnicode), m_nOffset(0)
{
	if (pData == NULL)
	{
		m_pData = NULL;
		m_bOwn = false;
		return;
	}

	// Do we have a NULL terminator on the string?
	if ((pData[dwSize-1] == (char)'\0') && (!bUnicode || (pData[dwSize-2] == (char)'\0')))
	{
		// Yes, keep it the same
		m_pData = pData;
		m_bOwn = bOwn;
	}
	else
	{
		// No, we want to copy it:
		m_pData = new char[dwSize + 2];
		memcpy(m_pData, pData, dwSize);
		// And add a terminator
		m_pData[dwSize] = (char)'\0';
		m_pData[dwSize + 1] = (char)'\0';
		m_bOwn = true;
		if (bOwn)
			// Delete unnecessary buffer
			delete [] pData;
	}
}

/**
	@param lpFilename Name of file to load
*/
FileINI::FileINI(LPCTSTR lpFilename) : m_pData(NULL), m_bOwn(false), m_bUnicode(false), m_nOffset(0)
{
	// Just load the file
	LoadINIFile(lpFilename);
}

/**
	@param lpFilename The file to load
	@return true if loaded successfully, false if failed
*/
bool FileINI::LoadINIFile(LPCTSTR lpFilename)
{
	// Open the file for reading, in text mode
	FILE* pFile;		
	if (NULL != _tfopen_s(&pFile, lpFilename, _T("rt")))
		return false;

	// Clean old data
	if ((m_pData != NULL) && m_bOwn)
	{
		delete [] m_pData;
		m_pData = NULL;
	}

	// Find the size of the file
	fseek(pFile, 0, SEEK_END);
	long lSize = ftell(pFile);
	std::tstring::size_type lReadSize;
	fseek(pFile, 0, SEEK_SET);

	// Create buffer
	m_pData = new char[(lSize + 1) * 2];
	m_bOwn = true;

	// Read the file
	lReadSize = fread(m_pData, 1, lSize, pFile);
	fclose(pFile);

	// Terminate
	m_pData[lReadSize] = (char)'\0';
	m_pData[lReadSize+1] = (char)'\0';

	// Set up unicode (if it is unicode)
	m_bUnicode = TestUnicode();
	if (m_bUnicode)
		m_nOffset = 2;

	return true;
}

/**
	@param pFile Pointer to file to test
	@return true if the file is unicode
*/
bool FileINI::TestUnicode(FILE* pFile)
{
	// Get the two characters here
	char c[2];
	long lSize = (long)fread(c, sizeof(char), 2, pFile);
	// fread might return something bigger than long... but we use it in fseek, which won't accept something bigger than long...
	if (lSize < 2)
	{
		if (lSize > 0)
			fseek(pFile, -lSize, SEEK_CUR);
		// Weren't found, isn't Unicode by default
		return false;
	}

	// Are those the Unicode marker? (0xFFFE)
	if ((c[0] == (char)0xFF) && (c[1] == (char)0xFE))
	{
		// Yes, read as binary from here on
		_setmode(_fileno(pFile), _O_BINARY);
		return true;
	}
	// OK, move back
	fseek(pFile, -2, SEEK_CUR);
	return false;
}

/**
	@return true if the data is Unicode, false otherwise
*/
bool FileINI::TestUnicode()
{
	if (m_pData == NULL)
		return false;
	return ((m_pData[0] == (char)0xFF) && (m_pData[1] == (char)0xFE));
}

/**
	@param pFilename The full path to the INI file
	@param pSection The name of the section, from which to return lines, without the brackets []
	@param[out] listLines Returns the list of lines. Each line will be added as a string in this list.
	@param uFlags Any combination of \ref FileINIFlags "FileINI Flags", which modify the behaviour of this method.
	@return true if successful, false if error

	NOTE: If the requested section is not found in the file, an empty list is returned, and no error is indicated
*/
bool FileINI::GetLines(LPCTSTR pFilename, LPCTSTR pSection, TCHARSTRLIST& listLines, UINT uFlags)
{
	// Open the file for reading, in text mode
	FILE* pFile;
	if (NULL != _tfopen_s(&pFile, pFilename, _T("rt")))
		return false;

	// Check Unicodeness
	bool bUnicode = TestUnicode(pFile);

	// Initialize variables
	bool bFoundSection = false, bFoundSectionNow;
	WCHAR c[MAX_PATH + 1];
	std::tstring::size_type nLen, nSectionLen = _tcslen(pSection);

	// Keep reading lines until the end of file is reached
	while (!feof(pFile))
	{
		// Read the next line from the file
		if (bUnicode)
		{
			// Get the next UNICODE string
			if (fgetws(c, MAX_PATH, pFile) == NULL)
				break;

			// Get length
			nLen = wcslen(c);
		}
		else
		{
			// Get the next ASCII string
			if (fgets((char*)c, MAX_PATH, pFile) == NULL)
				break;
		
			// Get length
			nLen = strlen((char*)c);
		}

		// Trim it
		TrimLen(c, nLen, (uFlags & FILEINI_TRIM) != 0, bUnicode);
		// Skip empty lines
		if (nLen < 1)
			continue;

		// Check if we are in the section; bFoundSection will be updated correctly
		bFoundSectionNow = CheckLineForSection(c, nLen, pSection, nSectionLen, bUnicode, bFoundSection);
		if (bFoundSection && !bFoundSectionNow)
		{
			// Ignore comments if so specified by the flags
			if ((uFlags & FILEINI_IGNORE_COMMENTS) && IsComment(c, nLen, bUnicode))
				continue;
			if (uFlags & FILEINI_TRIM)
			{
				// Trim the line (and make it a tstring)
				std::tstring s = Trim(c, nLen, bUnicode);
				if (!s.empty())
					// OK, add it
					listLines.push_back(s);
			}
			else
			{
				// Translate the line to something we can use
				if (bUnicode)
					listLines.push_back(MakeTString(std::wstring(c, nLen)));
				else
					listLines.push_back(MakeTString(std::string((const char*)c, nLen)));
			}
		}
	}

	// Cleanup
	fclose(pFile);
	return true;
}

/**
	@param pFilename The full filename of the INI file to read
	@param pSection The section from which to get the list of keys and values
	@param[out] listLines Returns a list of INILine objects with all the keys and values found in the specified section of the INI file
	@param uFlags Any combination of \ref FileINIFlags "FileINI Flags", which modify the behaviour of this method.
	@return true if successful, false if error

	NOTE: If the requested section is not found in the file, an empty list is returned, and no error is indicated
*/
bool FileINI::GetKeys(LPCTSTR pFilename, LPCTSTR pSection, INILINELIST& listLines, UINT uFlags)
{
	// Open the file for reading, in text mode
	FILE* pFile;
	if (NULL != _tfopen_s(&pFile, pFilename, _T("rt")))
		return false;

	// Check Unicodeness
	bool bUnicode = TestUnicode(pFile);

	// Initialize variables
	bool bFoundSection = false, bFoundSectionNow;
	WCHAR c[MAX_PATH+1];
	std::tstring::size_type nLen, nSectionLen = _tcslen(pSection);

	// Keep reading lines until the end of file is reached
	while (!feof(pFile))
	{
		// Read the next line from the file
		if (bUnicode)
		{
			// Get the next UNICODE string
			if (fgetws(c, MAX_PATH, pFile) == NULL)
				break;

			// Get length
			nLen = wcslen(c);
		}
		else
		{
			// Get the next ASCII string
			if (fgets((char*)c, MAX_PATH, pFile) == NULL)
				break;

			// Get length
			nLen = strlen((char*)c);
		}

		// Trim it
		TrimLen(c, nLen, (uFlags & FILEINI_TRIM) != 0, bUnicode);
		// And skip the empty lines
		if (nLen < 1)
			continue;

		// Check if we are in the section; bFoundSection will be updated correctly
		bFoundSectionNow = CheckLineForSection(c, nLen, pSection, nSectionLen, bUnicode, bFoundSection);
		if (bFoundSection && !bFoundSectionNow)
			// Add the values pair to the list
			DoInitLine(c, nLen, listLines, uFlags, bUnicode);
	}

	// Finished
	fclose(pFile);
	return true;
}

/**
	@param pLine The line of data
	@param nLineLen Length of line (either chars or WCHARS according to the unicode flag)
	@param pSection The section header to look for
	@param nSectionLen Length of section header name
	@param bUnicode true if the data is in Unicode, false if in ASCII
	@param[out] bFoundSection Flag updated to indicate if the line is in the section
	@return true if this line starts the section, false otherwise
*/
bool FileINI::CheckLineForSection(const WCHAR* pLine, std::tstring::size_type nLineLen, LPCTSTR pSection, std::tstring::size_type nSectionLen, bool bUnicode, bool& bFoundSection)
{
	// First make the line into something we can use
	std::tstring sLine;
	if (bUnicode)
		sLine = MakeTString(std::wstring(pLine, nLineLen));
	else
		sLine = MakeTString(std::string((const char*)pLine, nLineLen));

	// Is this a section header?
	if ((sLine[0] != '[') || (sLine[nLineLen-1] != ']'))
		// No, leave it as it is
		return false;
	// OK, this is a section header, check it
	if (nLineLen != nSectionLen + 2)
		bFoundSection = false;
	else
		bFoundSection = _tcsnicmp(pSection, sLine.c_str() + 1, nSectionLen) == 0;

	// We return true only if this is the requested section
	return bFoundSection;
}

/**
	@param pLine The line to parse
	@param nLineLen Length of line (in chars or wide chars according to the bUnicode flag)
	@param[out] listLines List of pairs to update
	@param uFlags Parsing flags (combination of \ref FileINIFlags)
	@param bUnicode true if the data is in Unicode, false if in ASCII
*/
void FileINI::DoInitLine(const WCHAR* pLine, std::tstring::size_type nLineLen, INILINELIST& listLines, UINT uFlags, bool bUnicode)
{
	// Check comment if requested
	if ((uFlags & FILEINI_IGNORE_COMMENTS) && IsComment(pLine, nLineLen, bUnicode))
		return;

	// Translate this string into something workable
	std::tstring sLine;
	if (bUnicode)
		sLine = MakeTString(std::wstring(pLine, nLineLen));
	else
		sLine = MakeTString(std::string((const char*)pLine, nLineLen));

	// Find the equal sign
	std::tstring::size_type nPos = sLine.find('=');
	std::tstring sBefore, sAfter;
	if (nPos == std::tstring::npos)
	{
		// Not found, just a name
		sBefore = sLine;
		sAfter = _T("");
	}
	else
	{
		// OK, found a name and a value
		sBefore = sLine.substr(0, nPos);
		sAfter = sLine.substr(nPos + 1);
	}

	// Trim if flagged
	if (uFlags & FILEINI_TRIM)
	{
		sBefore = Trim(sBefore);
		sAfter = Trim(sAfter);
	}
	// Add the value pair
	listLines.push_back(INILine(sBefore, sAfter));
}

/**
	@param pFilename The full filename of the INI file to read
	@param pSection The section header to look for
	@param mapKeys Map of name/value pairs
	@param uFlags Parsing flags (combination of \ref FileINIFlags)
	@return true if all went well, false if something failed
*/
bool FileINI::GetKeys(LPCTSTR pFilename, LPCTSTR pSection, TCHARSTR2STR& mapKeys, UINT uFlags)
{
	// First load the section pairs in the regular way
	INILINELIST listLines;
	if (!GetKeys(pFilename, pSection, listLines, uFlags))
		return false;

	// Now translate this into the map
	for (INILINELIST::iterator i = listLines.begin(); i != listLines.end(); i++)
		mapKeys[(*i).sKey] = (*i).sValue;

	return true;
}

/**
	@param pVarValue String of comma separated values
	@param[out] lValues List of values
	@return true if all went well, false if failed

	Supports values separated by commas. If a value should include a comma, it MUST be wrapped in a quote before and after, like so:
	a,b,"c,d",e ==> 4 values ['a', 'b', 'c,d', 'e']
	a,b,"c,d"e,f --> Error (quote not closed correctly!)
	a,b,c"d,e"f ==> 4 values ['a', 'b', 'c"d', 'e"f']
*/
bool FileINI::GetVarValues(LPCTSTR pVarValue, TCHARSTRLIST& lValues)
{
	// Clean up
	lValues.clear();
	std::tstring sValue, sVarValue = Trim(std::tstring(pVarValue, _tcslen(pVarValue)));

	bool bQuote;
	std::tstring::size_type nLen;

	// Go over the string
	while (!sVarValue.empty())
	{
		// Quote support: check if we have one
		if (sVarValue[0] == (TCHAR)'"')
		{
			// Yes, remember we are in a quote
			bQuote = true;

			// Remove it
			sVarValue.erase(sVarValue.begin());

			// Make sure everything works
			if (sVarValue.empty())
				return false;
		}
		else
			// Nope
			bQuote = false;

		// Count the size of the value
		nLen = 0;
		while (nLen < sVarValue.size())
		{
			if (bQuote && (sVarValue[nLen] == (TCHAR)'"'))
				// Found the other quote, end here
				break;
			if (!bQuote && (sVarValue[nLen] == (TCHAR)','))
				// Found the comma, end here
				break;
			nLen++;
		}

		// OK, get the value
		if (nLen == sVarValue.size())
		{
			if (bQuote)
				// Quote started but didn't finish, fail
				return false;

			// OK, the whole thing is our last value
			sValue = sVarValue;
			sVarValue = _T("");
		}
		else
		{
			// Cut up the value
			sValue = sVarValue.substr(0, nLen-1);
			// Drop the comma or quote we stopped at
			sVarValue.erase(0, nLen);
			if (!sVarValue.empty() && (sVarValue[0] != ','))
				// This function doesn't support quotes that doesn't end the value, so fail:
				return false;
		}

		// Add the value
		if (!sValue.empty())
			lValues.push_back(sValue);
	}
	return true;
}

/**
	@param pValue The string to start with
	@param lVariables Map of variables to replace in the string; each variable can have multiple values
	@param iPos Current variable to check on the string
	@param lResults List to add the result to
	@return true if all went well, false if something went wrong
*/
bool ReplaceVariables(LPCTSTR pValue, const TCHARSTR2STRLIST& lVariables, TCHARSTR2STRLIST::const_iterator& iPos, TCHARSTRLIST& lResults)
{
	// Find if we need to replace anything
	TCHAR* pLoc;
	while ((iPos != lVariables.end()) && ((pLoc = (TCHAR*)_tcsstr(pValue, (*iPos).first.c_str())) == NULL))
		iPos++;

	if (iPos == lVariables.end())
	{
		// No matching variable to replace, so add to the results and return
		lResults.push_back(pValue);
		return true;
	}

	SIZETARRAY ar;
	std::tstring sRet = pValue;

	// OK, find all the locations of the current variable in the string:
	std::tstring::size_type nPos = (pLoc - pValue), nLen = (*iPos).first.size();
	ar.push_back(nPos);
	while ((nPos = sRet.find((*iPos).first.c_str(), nPos + nLen)) != std::string::npos)
		ar.push_back(nPos);

	// Keep the current variable
	TCHARSTR2STRLIST::const_iterator iWork = iPos;
	iPos++;

	// Replace with each possible value, and recurse
	for (TCHARSTRLIST::const_iterator i = (*iWork).second.begin(); i != (*iWork).second.end(); i++)
	{
		// Replace all the occurances
		sRet = pValue;
		for (SIZETARRAY::iterator j = ar.begin(); j != ar.end(); j++)
			sRet.replace((*j), nLen, (*i).c_str());

		// Recurse to handle the rest of the variables on this variant
		if (!ReplaceVariables(sRet.c_str(), lVariables, iPos, lResults))
			return false;
	}
	return true;
}

/**
	@param pValue The string to start with
	@param lVariables Map of variables to replace in the string; each variable can have multiple values
	@param lResults List to add the result to
	@return true if all went well, false if something went wrong
*/
bool FileINI::ReplaceVariables(LPCTSTR pValue, const TCHARSTR2STRLIST& lVariables, TCHARSTRLIST& lResults)
{
	return ::ReplaceVariables(pValue, lVariables, lVariables.begin(), lResults);
}

/**
	@param pFilename The full filename of the INI file to read
	@param[out] lSections List of available sections
	@return true if all went well, false if something failed
*/
bool FileINI::GetAllSections(LPCTSTR pFilename, TCHARSTRLIST& lSections)
{
	// Open the file for reading, in text mode
	FILE* pFile;
	if (NULL != _tfopen_s (&pFile, pFilename, _T("rt")))
		return false;

	// Check Unicodeness
	bool bUnicode = TestUnicode(pFile);

	// Initialize variables
	WCHAR c[MAX_PATH+1];
	std::tstring::size_type nLen;

	// Keep reading lines until the end of file is reached
	while (!feof(pFile))
	{
		// Read the next line from the file
		if (bUnicode)
		{
			// Get the next UNICODE string
			if (fgetws(c, MAX_PATH, pFile) == NULL)
				break;
			
			// Get length
			nLen = wcslen(c);
		}
		else
		{
			// Get the next ASCII string
			if (fgets((char*)c, MAX_PATH, pFile) == NULL)
				break;
			
			// Get length
			nLen = strlen((char*)c);
		}

		// Trim it
		TrimLen(c, nLen, true, bUnicode);

		// Skip empties
		if (nLen < 1)
			continue;

		if (bUnicode)
		{
			// Is this line is a section header?
			if ((c[0] == '[') && (c[nLen-1] == ']'))
				// Yep, add it
				lSections.push_back(MakeTString(std::wstring(c+1, nLen-2)));
		}
		else
		{
			// Is this line is a section header?
			if ((((char*)c)[0] == (char)'[') && (((char*)c)[nLen-1] == (char)']'))
				// Yep, add it
				lSections.push_back(MakeTString(std::string(((char*)c)+1, nLen-2)));
		}
	}

	// Clean up
	fclose(pFile);
	return true;
}

/**
	@param[out] lSections List of available sections
	@return true if all went well, false if something failed
*/
bool FileINI::GetAllSections(TCHARSTRLIST& lSections)
{
	// Initialize variables
	bool bFoundSection = false;
	std::tstring::size_type nLen, nLineLen;
	const WCHAR* pLine = (const WCHAR*)(m_pData + m_nOffset);
	std::tstring s;

	// Keep reading lines until the end of file is reached
	do
	{
		// Find the length of the line
		if (m_bUnicode)
			nLen = wcscspn(pLine, L"\r\n");
		else
			nLen = strcspn((const char*)pLine, "\r\n");

		// Do we have something here?
		if (nLen > 0)
		{
			// Get the trimmed length
			nLineLen = nLen;
			TrimLen(pLine, nLineLen, true, m_bUnicode);

			if (nLineLen > 1)
			{
				// Check if this is a section header
				if (m_bUnicode)
				{
					if ((pLine[0] == '[') && (pLine[nLen-1] == ']'))
						lSections.push_back(MakeTString(std::wstring(pLine+1, nLen-2)));
				}
				else
				{
					if ((((char*)pLine)[0] == (char)'[') && (((char*)pLine)[nLen-1] == (char)']'))
						lSections.push_back(MakeTString(std::string(((char*)pLine)+1, nLen-2)));
				}
			}
		}
		// Jump over empty lines
		FindNextLine(pLine, nLen, m_bUnicode);
	} while (m_bUnicode ? ((*pLine) != '\0') : ((*((const char*)pLine)) != (char)'\0'));

	return true;
}

/**
	@param pSection The section header to look for
	@param mapKeys Map of name/value pairs
	@param uFlags Parsing flags (combination of \ref FileINIFlags)
	@return true if all went well, false if something failed
*/
bool FileINI::GetKeys(LPCTSTR pSection, TCHARSTR2STR& mapKeys, UINT uFlags)
{
	// First load the section pairs in the regular way
	INILINELIST listLines;
	if (!GetKeys(pSection, listLines, uFlags))
		return false;

	// Now translate this into the map
	for (INILINELIST::iterator i = listLines.begin(); i != listLines.end(); i++)
		mapKeys[(*i).sKey] = (*i).sValue;

	return true;
}

/**
	@param pSection The section header to look for
	@param[out] listLines Returns a list of INILine objects with all the keys and values found in the specified section of the INI file
	@param uFlags Any combination of \ref FileINIFlags "FileINI Flags", which modify the behaviour of this method.
	@return true if successful, false if error
*/
bool FileINI::GetKeys(LPCTSTR pSection, INILINELIST& listLines, UINT uFlags /* = 0 */)
{
	// Do we have any data?
	if (m_pData == NULL)
		return false;

	// Initialize
	bool bFoundSection = false, bFoundSectionNow;
	std::tstring::size_type nLen, nLineLen, nSectionLen = _tcslen(pSection);
	const WCHAR* pLine = (const WCHAR*)(m_pData + m_nOffset);

	do
	{
		// Get the next line length
		if (m_bUnicode)
			nLen = wcscspn(pLine, L"\r\n");
		else
			nLen = strcspn((const char*)pLine, "\r\n");

		// Do we have data?
		if (nLen > 0)
		{
			// Trim line (keep length)
			nLineLen = nLen;
			TrimLen(pLine, nLineLen, (uFlags & FILEINI_TRIM) != 0, m_bUnicode);

			// Do we have something to work with?
			if (nLineLen > 1)
			{
				// Is this the section we want?
				bFoundSectionNow = CheckLineForSection(pLine, nLineLen, pSection, nSectionLen, m_bUnicode, bFoundSection);
				if (bFoundSection && !bFoundSectionNow)
					// Yes, do something with the line
					DoInitLine(pLine, nLineLen, listLines, uFlags, m_bUnicode);
			}
		}

		// Go to the next line (jump over empty lines)
		FindNextLine(pLine, nLen, m_bUnicode);
	} while (m_bUnicode ? ((*pLine) != '\0') : ((*((const char*)pLine)) != (char)'\0'));

	return true;
}

/**
	@param pSection The name of the section, from which to return lines, without the brackets []
	@param[out] listLines Returns the list of lines. Each line will be added as a string in this list.
	@param uFlags Any combination of \ref FileINIFlags "FileINI Flags", which modify the behaviour of this method.
	@return true if successful, false if error

	NOTE: If the requested section is not found in the file, an empty list is returned, and no error is indicated
*/
bool FileINI::GetLines(LPCTSTR pSection, TCHARSTRLIST& listLines, UINT uFlags /* = 0 */)
{
	// Do we have anything to work with?
	if (m_pData == NULL)
		return false;

	// Initialize variables
	bool bFoundSection = false, bFoundSectionNow;
	std::tstring::size_type nLen, nLineLen, nSectionLen = _tcslen(pSection);
	const WCHAR* pLine = (const WCHAR*)(m_pData + m_nOffset);
	std::tstring s;

	// Keep reading lines until the end of file is reached
	do
	{
		// Find length of current line
		if (m_bUnicode)
			nLen = wcscspn(pLine, L"\r\n");
		else
			nLen = strcspn((const char*)pLine, "\r\n");

		// Do we have anything here?
		if (nLen > 0)
		{
			// Find the trimmed length
			nLineLen = nLen;
			TrimLen(pLine, nLineLen, (uFlags & FILEINI_TRIM) != 0, m_bUnicode);

			// Is there anything left?
			if (nLineLen > 1)
			{
				// Are we in the section we are looking for?
				bFoundSectionNow = CheckLineForSection(pLine, nLineLen, pSection, nSectionLen, m_bUnicode, bFoundSection);
				if (bFoundSection && !bFoundSectionNow)
				{
					// Yes; ignore comments if so specified by the flags
					if ((uFlags & FILEINI_IGNORE_COMMENTS) && IsComment(pLine, nLineLen, m_bUnicode))
						continue;

					// Convert into something we can work with
					if (m_bUnicode)
						s = MakeTString(std::wstring(pLine, nLineLen));
					else
						s = MakeTString(std::string((const char*)pLine, nLineLen));

					// Trim if wanted
					if (uFlags & FILEINI_TRIM)
						s = Trim(s);

					// Add if not empty
					if (!s.empty())
						listLines.push_back(s);
				}
			}
		}

		// Jump over empty lines
		FindNextLine(pLine, nLen, m_bUnicode);
	} while (m_bUnicode ? ((*pLine) != '\0') : ((*((const char*)pLine)) != (char)'\0'));

	return true;
}
