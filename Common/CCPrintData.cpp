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
#include "CCPrintData.h"

#include "CCPrintRegistry.h"

#include <time.h>

// File format
/*
[Job]
PageCount=<page_count> // not every page HAS to appear as a section!
TestPage=<flag> // Optional, defaults to 0; set to 1 to get back the actual link location info: in goes TEXT based links, out comes LOCATION based links

// for each page
[Page <page_num>]
URLCount=<url_count>
// Optional sync data
Width=<width>
Height=<height>

 // External links (URLs):
URL<num>=<url>
Title<num>=<title> // Optional

// Inner links
Page<num>=<page>
OffsetX<num>=<offset-x>
OffsetY<num>=<offset-y>

// for each TEXT-based URL:
Text<num>=<text>
Repeat<num>=<repeat> // Optional (defaults to 1)

// For each LOCATION-based link
Left<num>=<left-location>
Right<num>=<right-location>
Top<num>=<top-location>
Bottom<num>=<bottom-location>

*/

// Registry values
/// Registry value name for job time
#define JOBDATA_TIME_PREFIX				_T("CCJobTime")
/// Registry value name for job time with process
#define JOBDATA_TIME_KEY				JOBDATA_TIME_PREFIX _T("%d")
/// Registry value name for data file (with process)
#define JOBDATA_FILE_KEY				_T("CCJobFile%d")

// INI-like file values
/// General section name
#define DATAFILE_SECTION_MAIN			_T("Job")
/// General section name (write)
#define DATAFILE_SECTION_MAIN_WRITE		_T("[Job]\r\n")
/// Page section name
#define DATAFILE_SECTION_PAGE			_T("Page %d")
/// Page section name (write)
#define DATAFILE_SECTION_PAGE_WRITE		_T("[Page %d]\r\n")

// General section values
/// Page count value
#define DATAFILE_PAGE_COUNT				_T("PageCount")
/// Page count value (write)
#define DATAFILE_PAGE_COUNT_WRITE		_T("PageCount=%d\r\n")
/// Test page value
#define DATAFILE_PAGE_TEST				_T("TestPage")
/// Test page value (write)
#define DATAFILE_PAGE_TEST_WRITE		_T("TestPage=%d\r\n")

// Page section values
/// Number of links in page
#define DATAFILE_LINK_COUNT				_T("LinkCount")
/// Number of links in page (write)
#define DATAFILE_LINK_COUNT_WRITE		_T("LinkCount=%d\r\n")

// Page section values (per link)
/// Link: URL value
#define DATAFILE_LINK_URL				_T("URL%d")
/// Link: Title value
#define DATAFILE_LINK_TITLE				_T("Title%d")

// Internal link information
/// Internal link location: Page value
#define DATAFILE_LINK_PAGE				_T("Page%d")
/// Internal link location: Page value (write)
#define DATAFILE_LINK_PAGE_WRITE		_T("Page%d=%d\r\n")
/// Internal link location: X Offset value
#define DATAFILE_LINK_OFFSET_X			_T("OffsetX%d")
/// Internal link location: X Offset value (write)
#define DATAFILE_LINK_OFFSET_X_WRITE	_T("OffsetX%d=%d\r\n")
/// Internal link location: Y Offset value
#define DATAFILE_LINK_OFFSET_Y			_T("OffsetY%d")
/// Internal link location: Y Offset value (write)
#define DATAFILE_LINK_OFFSET_Y_WRITE	_T("OffsetY%d=%d\r\n")

// Link location by text
/// Link location: text to find
#define DATAFILE_LINK_TEXT				_T("Text%d")
/// Link location: text repeat count
#define DATAFILE_LINK_REPEAT			_T("Repeat%d")
/// Link location: text repeat count (write)
#define DATAFILE_LINK_REPEAT_WRITE		_T("Repeat%d=%d\r\n")

// Link location rectangle
/// Link location: left
#define DATAFILE_LINK_LOC_LEFT			_T("Left%d")
/// Link location: left (write)
#define DATAFILE_LINK_LOC_LEFT_WRITE	_T("Left%d=%d\r\n")
/// Link location: right
#define DATAFILE_LINK_LOC_RIGHT			_T("Right%d")
/// Link location: right (write)
#define DATAFILE_LINK_LOC_RIGHT_WRITE	_T("Right%d=%d\r\n")
/// Link location: top
#define DATAFILE_LINK_LOC_TOP			_T("Top%d")
/// Link location: top (write)
#define DATAFILE_LINK_LOC_TOP_WRITE		_T("Top%d=%d\r\n")
/// Link location: bottom
#define DATAFILE_LINK_LOC_BOTTOM		_T("Bottom%d")
/// Link location: bottom (write)
#define DATAFILE_LINK_LOC_BOTTOM_WRITE	_T("Bottom%d=%d\r\n")

// Test page return values
/// Page width
#define DATAFILE_PAGE_WIDTH				_T("Width")
/// Page width (write)
#define DATAFILE_PAGE_WIDTH_WRITE		_T("Width=%d\r\n")
/// Page height
#define DATAFILE_PAGE_HEIGHT			_T("Height")
/// Page height (write)
#define DATAFILE_PAGE_HEIGHT_WRITE		_T("Height=%d\r\n")

/**
	
*/
void CCPrintData::LinkData::CleanText()
{
	std::tstring::size_type pos;
	while ((pos = sText.find_first_of(_T("\r\n"))) != std::tstring::npos)
		sText.erase(pos, 1);
}

/**
	@param data The key/value set to read the data from
	@param nNum The number of the link to read the data of
	@return true if loaded successfully, false if failed
*/
bool CCPrintData::LinkData::FromFile(const TCHARSTR2STR& data, int nNum)
{
	// Initialize variables
	TCHAR cName[32];
	TCHARSTR2STR::const_iterator iKey;
	sText = _T("");
	sURL = _T("");
	sTitle = _T("");
	nPage = 0;

	// Could be an internal link
	_stprintf_s(cName, _S(cName), DATAFILE_LINK_PAGE, nNum);
	if ((iKey = data.find(cName)) != data.end())
	{
		// Internal link, read X, Y and page
		nPage = _ttoi((*iKey).second.c_str());
		if (nPage < 1)
			return false;
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_OFFSET_X, nNum);
		if ((iKey = data.find(cName)) != data.end())
			ptOffset.x = _ttoi((*iKey).second.c_str());
		else
			ptOffset.x = 0;
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_OFFSET_Y, nNum);
		if ((iKey = data.find(cName)) != data.end())
			ptOffset.y = _ttoi((*iKey).second.c_str());
		else
			ptOffset.y = 0;
	}
	else
	{
		// External link, read URL
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_URL, nNum);
		if ((iKey = data.find(cName)) == data.end())
			return false;
		// This is an external link (URL)
		sURL = (*iKey).second;
	}

	// Title (if we have one)
	_stprintf_s(cName, _S(cName), DATAFILE_LINK_TITLE, nNum);
	if ((iKey = data.find(cName)) != data.end())
		sTitle = (*iKey).second;

	_stprintf_s(cName, _S(cName), DATAFILE_LINK_TEXT, nNum);
	if ((iKey = data.find(cName)) != data.end())
	{
		// Text-based link, read the text and repeat count
		sText = (*iKey).second;
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_REPEAT, nNum);
		if ((iKey = data.find(cName)) != data.end())
		{
			int nTemp = _ttoi((*iKey).second.c_str());
			if (nTemp > 0)
				nRepeat = nTemp;
		}
	}
	else
	{
		// Location-based link, read left, right, top, bottom
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_LOC_LEFT, nNum);
		if ((iKey = data.find(cName)) == data.end())
			return false;
		rectLocation.left = _ttoi((*iKey).second.c_str());
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_LOC_RIGHT, nNum);
		if ((iKey = data.find(cName)) == data.end())
			return false;
		rectLocation.right = _ttoi((*iKey).second.c_str());
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_LOC_TOP, nNum);
		if ((iKey = data.find(cName)) == data.end())
			return false;
		rectLocation.top = _ttoi((*iKey).second.c_str());
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_LOC_BOTTOM, nNum);
		if ((iKey = data.find(cName)) == data.end())
			return false;
		rectLocation.bottom = _ttoi((*iKey).second.c_str());
	}
	return true;
}

/**
	@param sData String to write data to
	@param nNum The number of this link
	@return true if writen successfully, false if failed
*/
bool CCPrintData::LinkData::ToFile(std::tstring& sData, int nNum) const
{
	TCHAR cName[64];

	// Is this an internal link?
	if (IsInner())
	{
		// Yes, write page, X and Y
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_PAGE_WRITE, nNum, nPage);
		sData += cName;
		if (ptOffset.x != 0)
		{
			_stprintf_s(cName, _S(cName), DATAFILE_LINK_OFFSET_X_WRITE, nNum, ptOffset.x);
			sData += cName;
		}
		if (ptOffset.y != 0)
		{
			_stprintf_s(cName, _S(cName), DATAFILE_LINK_OFFSET_Y_WRITE, nNum, ptOffset.y);
			sData += cName;
		}
	}
	else
	{
		// No, write URL
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_URL, nNum);
		sData += cName;
		sData += _T("=") + sURL + _T("\n\r");
	}

	if (!sTitle.empty())
	{
		// Write tooltip
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_TITLE, nNum);
		sData += cName;
		sData += _T("=") + sTitle + _T("\n\r");
	}

	// Is this a location-based link?
	if (IsLocation())
	{
		// Yes, write left, right, top and bottom
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_LOC_LEFT_WRITE, nNum, rectLocation.left);
		sData += cName;
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_LOC_RIGHT_WRITE, nNum, rectLocation.right);
		sData += cName;
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_LOC_TOP_WRITE, nNum, rectLocation.top);
		sData += cName;
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_LOC_BOTTOM_WRITE, nNum, rectLocation.bottom);
		sData += cName;
	}
	else
	{
		// No, write the text to match
		_stprintf_s(cName, _S(cName), DATAFILE_LINK_TEXT, nNum);
		sData += cName;
		sData += _T("=") + sText + _T("\n\r");
		if (nRepeat > 1)
		{
			// And the repeat count if more then one
			_stprintf_s(cName, _S(cName), DATAFILE_LINK_REPEAT_WRITE, nNum, nRepeat);
			sData += cName;
		}
	}
	return true;
}



/**
	@param file INI file object to read from
	@param nPage Number of page to read from
	@return true if read successfully, false if failed
*/
bool CCPrintData::PageData::FromFile(class FileINI& file, int nPage)
{
	// Variables
	TCHAR cName[32];
	TCHARSTR2STR data;
	TCHARSTR2STR::const_iterator iKey;

	// Clean this object
	Clear();

	// Get INI key/value pairs from the page's section
	_stprintf_s(cName, _S(cName), DATAFILE_SECTION_PAGE, nPage);
	if (!file.GetKeys(cName, data) || ((iKey = data.find(DATAFILE_LINK_COUNT)) == data.end()))
		// None found, just go on
		return true;

	// Get the page link count
	int nLinkCount = _ttoi((*iKey).second.c_str());
	if (nLinkCount < 1)
		// This shouldn't happen
		return false;

	// Read all links
	for (int iLink = 1; iLink <= nLinkCount; iLink++)
	{
		LinkData link;
		if (link.FromFile(data, iLink))
			push_back(link);
	}

	// Get page width and height if we found them
	if ((iKey = data.find(DATAFILE_PAGE_WIDTH)) != data.end())
		szPage.cx = _ttoi((*iKey).second.c_str());
	if ((iKey = data.find(DATAFILE_PAGE_HEIGHT)) != data.end())
		szPage.cy = _ttoi((*iKey).second.c_str());

	return true;
}

/**
	@param sData String to write the data to
	@param nPage The number of this page
	@return true if writen successfully, false if failed
*/
bool CCPrintData::PageData::ToFile(std::tstring& sData, int nPage) const
{
	// If this page has no data, don't write anything...
	if (empty())
		return true;

	// Write the section name
	TCHAR cName[64];
	_stprintf_s(cName, _S(cName), DATAFILE_SECTION_PAGE_WRITE, nPage);
	sData += cName;
	// Write the link count
	_stprintf_s(cName, _S(cName), DATAFILE_LINK_COUNT_WRITE, size());
	sData += cName;

	// Write all the links
	int iLink = 1;
	for (const_iterator i = begin(); i != end(); i++, iLink++)
		if (!(*i).ToFile(sData, iLink))
			return false;

	if (szPage.cx != 0)
	{
		// Write page width
		_stprintf_s(cName, _S(cName), DATAFILE_PAGE_WIDTH_WRITE, szPage.cx);
		sData += cName;
	}
	if (szPage.cy != 0)
	{
		// Write page height
		_stprintf_s(cName, _S(cName), DATAFILE_PAGE_HEIGHT_WRITE, szPage.cy);
		sData += cName;
	}

	return true;
}


/**
	@param hPrinter Handle to printer
	@return true if updated the process data successfully, false if failed
*/
bool CCPrintData::UpdateProcessData(HANDLE hPrinter)
{
	// Get process ID
	TCHAR	cKeyFile[64];
	DWORD	dwProcessID = GetCurrentProcessId();
	
	// Check registry for file
	_stprintf_s(cKeyFile, _S(cKeyFile), JOBDATA_FILE_KEY, dwProcessID);
	std::tstring sFilename = CCPrintRegistry::GetRegistryString(hPrinter, cKeyFile, _T(""));
	if (sFilename.empty())
		// Not found, cannot update if not found
		return false;

	// Write the data to it
	return WriteToFile(sFilename.c_str());
}

/**
	@param hPrinter Handle to printer
	@return true if reloaded the data auccessfully, false if failed
*/
bool CCPrintData::ReloadProcessData(HANDLE hPrinter)
{
	// Get process ID
	TCHAR	cKeyFile[64];
	DWORD	dwProcessID = GetCurrentProcessId();

	// Check registry for file
	_stprintf_s(cKeyFile, _S(cKeyFile), JOBDATA_FILE_KEY, dwProcessID);
	std::tstring sFilename = CCPrintRegistry::GetRegistryString(hPrinter, cKeyFile, _T(""));
	if (sFilename.empty())
		// Not found, cannot reload
		return false;

	// Read the data
	return ReadFromFile(sFilename.c_str());
}

/**
	@param hPrinter Handle to the printer
	@return true if loaded the data auccessfully, false if failed to load
*/
bool CCPrintData::LoadProcessData(HANDLE hPrinter)
{
	// Get process ID
	TCHAR	cKeyTime[64], cKeyFile[64];
	DWORD	dwProcessID = GetCurrentProcessId(), dwTimeKey;
	time_t	tNow, tKey;

	// Clean this data and all the old data
	CleanThis();
	CleanOldData(hPrinter);

	// Check registry for time
	_stprintf_s(cKeyTime, _S(cKeyTime), JOBDATA_TIME_KEY, dwProcessID);
	if ((dwTimeKey = CCPrintRegistry::GetRegistryDWORD(hPrinter, cKeyTime, 0)) == 0)
		// No time, no data
		return false;
	
	// Compare times:
	tKey = (time_t)dwTimeKey;
	tNow = time(NULL);

	// Check if time stamp current (up to 5 minutes)
	bool bRet = false;
	if (difftime(tNow, tKey) < 300)
	{
		// OK, not too old, read the file name
		_stprintf_s(cKeyFile, _S(cKeyFile), JOBDATA_FILE_KEY, dwProcessID);
		std::tstring sFilename = CCPrintRegistry::GetRegistryString(hPrinter, cKeyFile, _T(""));
		if (!sFilename.empty())
			// Found a filename, read the file
			bRet = ReadFromFile(sFilename.c_str());
	}
	
	return bRet;
}

/**
	@param lpFilename Path to file to read
	@return true if read successfully, false if failed
*/
bool CCPrintData::ReadFromFile(LPCTSTR lpFilename)
{
	//	Read the file
	FileINI file;
	if (!file.LoadINIFile(lpFilename))
		return false;

	TCHARSTR2STR data;
	TCHARSTR2STR::const_iterator iKey;

	// Get main section
	file.GetKeys(DATAFILE_SECTION_MAIN, data);
	if ((iKey = data.find(DATAFILE_PAGE_TEST)) != data.end())
		// Get test page value
		m_bTestPage = _ttoi((*iKey).second.c_str()) != 0;

	// Get page count
	iKey = data.find(DATAFILE_PAGE_COUNT);
	if (iKey != data.end())
	{
		int nCount = _ttoi((*iKey).second.c_str());
		if (nCount > 0)
		{
			// Set pages array and read datas data
			m_pages.resize(nCount);
			for (int nPage = 1; nPage <= nCount; nPage++)
				m_pages[nPage-1].FromFile(file, nPage);
		}
	}

	return true;
}

/**
	@param hPrinter Handle to printer
*/
void CCPrintData::CleanOldData(HANDLE hPrinter)
{
	// Do we have any timestamp registry keys for the printer?
	STRLIST lValues;
	if (!CCPrintRegistry::EnumRegistryValues(hPrinter, lValues, JOBDATA_TIME_PREFIX))
		// Nothing for us
		return;

	DWORD dwTimeKey, dwProcessID;
	time_t	tNow, tKey;

	// Go over values
	for (STRLIST::const_iterator i = lValues.begin(); i != lValues.end(); i++)
	{
		// Get the value
		if ((dwTimeKey = CCPrintRegistry::GetRegistryDWORD(hPrinter, (*i).c_str(), 0)) == 0)
			// Nothing here
			continue;

		tKey = (time_t)dwTimeKey;
		tNow = time(NULL);

		// Check if this is old (5 minutes or more)
		if (difftime(tNow, tKey) > 300)
		{
			// Yes, so get the process ID
			dwProcessID = _ttoi((*i).c_str() + _tcslen(JOBDATA_TIME_PREFIX));
			if (dwProcessID == 0)
				continue;
			// And clean it up
			CleanData(hPrinter, dwProcessID);
		}
	}
}

/**
	@param hPrinter Handle to printer
	@param dwProcessID Process ID
*/
void CCPrintData::CleanData(HANDLE hPrinter, DWORD dwProcessID)
{
	TCHAR	cKey[64];
	// Get the filename from the key
	_stprintf_s(cKey, _S(cKey), JOBDATA_FILE_KEY, dwProcessID);
	std::tstring sFilename = CCPrintRegistry::GetRegistryString(hPrinter, cKey, _T(""));
	// Delete file
	if (!sFilename.empty())
		_tunlink(sFilename.c_str());
	// Erase the registry keys
	CCPrintRegistry::EraseRegistryValue(hPrinter, cKey);
	// Delete time key
	_stprintf_s(cKey, _S(cKey), JOBDATA_TIME_KEY, dwProcessID);
	CCPrintRegistry::EraseRegistryValue(hPrinter, cKey);
}

/**
	@param hPrinter Handle to the printer
*/
void CCPrintData::CleanSaved(HANDLE hPrinter)
{
	// Clean the data for this process
	CleanData(hPrinter, GetCurrentProcessId());
}

/**
	@param hPrinter Handle to the printer
	@return true if successfully saved, false if failed
*/
bool CCPrintData::SaveProcessData(HANDLE hPrinter)
{
	// Get process ID
	DWORD	dwProcessID = GetCurrentProcessId();

	// Remove old data for this process, if any
	CleanData(hPrinter, dwProcessID);

	// Save temporary file:
	// Get temp filename
	TCHAR cFilename[MAX_PATH + 1], cPath[MAX_PATH + 1], cFileKey[64], cTimeKey[64];
	if (GetTempPath(MAX_PATH, cPath) == 0)
		return false;
	if (GetTempFileName(cPath, _T("CCPDF"), 0, cFilename) == 0)
		return false;

	// Save to the file
	if (!WriteToFile(cFilename))
	{
		// Didn't work! Bale!
		_tunlink(cFilename);
		return false;
	}

	// Write timestamp to registry
	_stprintf_s(cTimeKey, _S(cTimeKey), JOBDATA_TIME_KEY, dwProcessID);
	if (!CCPrintRegistry::SetRegistryDWORD(hPrinter, cTimeKey, (DWORD) time(NULL)))
	{
		_tunlink(cFilename);
		return false;
	}
	// Write filename to registry
	_stprintf_s(cFileKey, _S(cFileKey), JOBDATA_FILE_KEY, dwProcessID);
	if (!CCPrintRegistry::SetRegistryString(hPrinter, cFileKey, cFilename))
	{
		CCPrintRegistry::EraseRegistryValue(hPrinter, cTimeKey);
		_tunlink(cFilename);
		return false;
	}
	return true;
}

/**
	@param lpFilename Name of write to write to
	@return true if successfully write, false if failed
*/
bool CCPrintData::WriteToFile(LPCTSTR lpFilename)
{
	// Variables
	std::tstring sData;
	TCHAR cName[64];

	// Write main section name
	sData += DATAFILE_SECTION_MAIN_WRITE;

	// Write amount of pages
	_stprintf_s(cName, _S(cName), DATAFILE_PAGE_COUNT_WRITE, m_pages.size());
	sData += cName;
	if (m_bTestPage)
	{
		// Write test page flag
		_stprintf_s(cName, _S(cName), DATAFILE_PAGE_TEST_WRITE, 1);
		sData += cName;
	}

	// Write all the pages
	for (unsigned int iPage = 0; iPage < m_pages.size(); iPage++)
		if (!m_pages[iPage].ToFile(sData, iPage + 1))
			return false;
	
	// Now open the file
	FILE* pFile;
	if (NULL != _tfopen_s(&pFile, lpFilename, _T("wt")))
		return false;
#ifdef _UNICODE
	// Write unicode identifier at the beginning if it's a unicode string
	unsigned char cUni[2];
	cUni[0] = 0xFF;
	cUni[1] = 0xFE;
	if (fwrite(cUni, sizeof(unsigned char), 2, pFile) != 2)
	{
		fclose(pFile);
		return false;
	}
#endif
	// Write the data
	if (fwrite(sData.c_str(), sizeof(TCHAR), sData.size(), pFile) != sData.size())
	{
		fclose(pFile);
		return false;
	}

	// That's it!
	fclose(pFile);
	return true;
}

/**
	@param nPage The page to ensure exists
*/
void CCPrintData::EnsurePage(int nPage)
{
	while ((int)m_pages.size() < nPage)
		m_pages.push_back(PageData());
}

/**
	@param sURL The link's URL
	@param sText The text to look for
	@param nPage The page in which this link is printed
	@param nRepeat The amount of times to look for the text before making it the link
*/
void CCPrintData::AddLink(const std::tstring& sURL, const std::tstring& sText, int nPage, int nRepeat /* = 1 */)
{
	// Ensure we have enough pages
	EnsurePage(nPage);
	// Add the new link data
	m_pages[nPage - 1].push_back(LinkData(sURL, sText, nRepeat));
}

/**
	@param sURL The link's URL
	@param rect The link location
	@param nPage The page in which this link is printed
	@param lpTitle The link's tooltip (future)
*/
void CCPrintData::AddLink(const std::tstring& sURL, const RECTL& rect, int nPage, LPCTSTR lpTitle /* = NULL */)
{
	// Ensure we have enough pages
	EnsurePage(nPage);
	// Add the new link data
	m_pages[nPage - 1].push_back(LinkData(sURL, rect, lpTitle));
}

/**
	@param rect Location of the link
	@param nPage The page in which this link is printed
	@param nDestPage The page to which the internal link links
	@param lX X Offset in the linked page
	@param lY Y Offset in the linked page
	@param lpTitle The link's tooltip (future)
*/
void CCPrintData::AddLink(const RECTL& rect, int nPage, int nDestPage, long lX, long lY, LPCTSTR lpTitle /* = NULL */)
{
	// Ensure we have enough pages
	EnsurePage(nPage);
	// Add the new link data
	m_pages[nPage - 1].push_back(LinkData(rect, nDestPage, lX, lY, lpTitle));
}

#ifdef _DEBUG
/**
	
*/
void CCPrintData::Dump()
{
	std::tstring sData;
	TCHAR cName[64];
	sData += DATAFILE_SECTION_MAIN_WRITE;
	_stprintf_s(cName, _S(cName), DATAFILE_PAGE_COUNT_WRITE, m_pages.size());
	sData += cName;
	if (m_bTestPage)
	{
		_stprintf_s(cName, _S(cName), DATAFILE_PAGE_TEST_WRITE, 1);
		sData += cName;
	}
	for (unsigned int iPage = 0; iPage < m_pages.size(); iPage++)
		if (!m_pages[iPage].ToFile(sData, iPage + 1))
			return;
	
	::OutputDebugString(sData.c_str());
}
#endif
