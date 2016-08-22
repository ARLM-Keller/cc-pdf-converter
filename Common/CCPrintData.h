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

#ifndef _CCPRINTDATA_H_
#define _CCPRINTDATA_H_

#include "FileINI.h"
#include <vector>

/**
    @brief PDF printind data class; written to a file with a link to it in the registry
*/
class CCPrintData
{
public:
	/**
		@brief Default constructor
	*/
	CCPrintData() : m_bTestPage(false) {};

	// Helper structures
	/**
	    @brief Internal helper class for link data
	*/
	struct LinkData
	{
		/**
			@brief Default constructor
		*/
		LinkData() : nRepeat(1), nPage(0) {rectLocation.left = rectLocation.right = rectLocation.top = rectLocation.bottom = 0; ptOffset.x = ptOffset.y = 0;};
		/**
			@brief Copy constructor
			@param other The link data to copy
		*/
		LinkData(const LinkData& other) : sText(other.sText), sURL(other.sURL), sTitle(other.sTitle), nRepeat(other.nRepeat), rectLocation(other.rectLocation), ptOffset(other.ptOffset), nPage(other.nPage) {};
		/**
			@brief Create a text-based link data
			@param sU The URL
			@param sT The text to find
			@param n The repeat count
			@param lpTitle The link tooltip (future)
		*/
		LinkData(const std::tstring& sU, const std::tstring& sT, int n, LPCTSTR lpTitle = NULL) : sURL(sU), sText(sT), nRepeat(n), nPage(0), sTitle(lpTitle == NULL ? _T("") : lpTitle) {CleanText();};
		/**
			@brief Creates a location-based link data
			@param sU The URL
			@param rect The location
			@param lpTitle The link tooltip (future)
		*/
		LinkData(const std::tstring& sU, const RECTL& rect, LPCTSTR lpTitle = NULL) : sURL(sU), rectLocation(rect), nPage(0), sTitle(lpTitle == NULL ? _T("") : lpTitle) {};
		/**
			@brief Creates a location-based internal link data
			@param rect The link location
			@param nP The page to link to
			@param lX The X offset to link to
			@param lY The Y offset to link to
			@param lpTitle The link tooltip (future)
		*/
		LinkData(const RECTL& rect, int nP, long lX, long lY, LPCTSTR lpTitle = NULL) : nRepeat(1), nPage(nP), rectLocation(rect), sTitle(lpTitle == NULL ? _T("") : lpTitle) {CleanText(); ptOffset.x = lX; ptOffset.y = lY;};

		// Data
		/// Text of link; empty for location links
		std::tstring	sText;
		/// URL for external links, text for internal links
		std::tstring	sURL;
		/// Tooltip text (can be empty)
		std::tstring	sTitle;
		/// Repeast count for text links
		int				nRepeat;
		/// Location of link (location links only)
		RECTL			rectLocation;
		/// Page for internal location links
		int				nPage;
		/// Offset in page (internal links only)
		POINTL			ptOffset;

		/// Write the data in an INI file format
		bool			ToFile(std::tstring& sData, int nNum) const;
		/// Read the data from a key/value set
		bool			FromFile(const TCHARSTR2STR& data, int nNum);
		/**
			@brief Check if it is a location (vs. text-based) link
			@return true if this is a location link, false for text links
		*/
		bool			IsLocation() const {return sText.empty();};
		/**
			@brief Checks if this is an internal link
			@return true if this is an internal link, false for external (URL) links
		*/
		bool			IsInner() const {return IsLocation() && (nPage != 0);};
		/// Cleans the text-link representation
		void			CleanText();
	};

	/**
	    @brief Structure for keeping data for all the links in a page (and some extra page-specific data)
	*/
	struct PageData : std::list<LinkData>
	{
		/**
			@brief Default constructor
		*/
		PageData() {szPage.cx = szPage.cy = 0;};
		/**
			@brief Copy constructor
			@param other Data to copy
		*/
		PageData(const PageData& other) : szPage(other.szPage) {insert(begin(), other.begin(), other.end());};

		// Data
		/// Size of page
		SIZEL			szPage;

		/// Write the data in an INI file format
		bool			ToFile(std::tstring& sData, int nPage) const;
		/// Read the data from an INI file
		bool			FromFile(class FileINI& file, int nPage);
		/**
			@brief Checks if there's at least one text link in the data
			@return true if there's a text link, false if all the links are location-based
		*/
		bool			HasTextLink() const {for (const_iterator i = begin(); i != end(); i++) if (!(*i).IsLocation()) return true; return false;};
		/**
			@brief Clears the page data
		*/
		void			Clear() {clear(); szPage.cx = szPage.cy = 0;};
	};

protected:
	// Data
	/// Array of pages data
	std::vector<PageData> m_pages;
	/// Default (empty) page data
	PageData m_dummy;
	/// true if this is a test run (for finding Excel factors, for example)
	bool	m_bTestPage;

public:
	// Data Access
	/**
		@brief Checks if this is a test run
		@return true if this is a test run, false if not
	*/
	bool	IsTestPage() const {return m_bTestPage;};
	/**
		@brief Check if there's any link data here
		@return true if there's any data, false for no links
	*/
	bool	HasData() const {return !m_pages.empty();};
	/**
		@brief Returns the link data for the requested page
		@param nPage The page to get data for (1-based)
		@return The data for the page, or an empty (dummy) data if there's no such page defined
	*/
	const PageData& GetPageData(int nPage) const 
	{
		if ((nPage < 1) || (nPage > (int)m_pages.size())) 
			return m_dummy; 
		return m_pages[nPage - 1];
	};

	/**
		@brief Returns the number of pages we have data for
		@return The number of pages in the structure
	*/
	size_t		GetPageCount() const {return m_pages.size();};
	/**
		@brief Set the page test flag
		@param bSet Value to set
	*/
	void	SetTestPage(bool bSet = true) {m_bTestPage = bSet;};
	/// Add a text-based link data
	void	AddLink(const std::tstring& sURL, const std::tstring& sText, int nPage, int nRepeat = 1);
	/// Add a location-based link data
	void	AddLink(const std::tstring& sURL, const RECTL& rect, int nPage, LPCTSTR lpTitle = NULL);
	/// Add a location-based INTERNAL link data
	void	AddLink(const RECTL& rect, int nPage, int nDestPage, long lX, long lY, LPCTSTR lpTitle = NULL);
	/**
		@brief Set the size of a page (return data only)
		@param nPage The page to update
		@param sz The page printing size
	*/
	void	SetPageSize(int nPage, const SIZEL& sz) 
	{
		if ((nPage < 1) || (nPage > (int)m_pages.size())) 
			return; 
		m_pages[nPage - 1].szPage = sz;
	};

		// Loading/Saving data for process
	/// Load the print data for this process from a file (use correct registry key)
	bool	LoadProcessData(HANDLE hPrinter);
	/// Save the print data for this process to a file (and update the registry key)
	bool	SaveProcessData(HANDLE hPrinter);
	/// Save the print data for this process to a file (use same file)
	bool	UpdateProcessData(HANDLE hPrinter);
	/// Load the print data for this process from a file (use same file)
	bool	ReloadProcessData(HANDLE hPrinter);

	// Methods
	/// Clean the file data for this process and the registry keys; also remove
	void	CleanSaved(HANDLE hPrinter);
	/**
		@brief Clean this object
	*/
	void	CleanThis() {m_pages.clear(); m_bTestPage = false;};

#ifdef _DEBUG
	/// Dump the object's data
	void	Dump();
#endif

protected:
	// Helper methods
	/// Clean any old data for any process (older then 5 minutes)
	void	CleanOldData(HANDLE hPrinter);
	/// Clean the current data for THIS process
	void	CleanData(HANDLE hPrinter, DWORD dwProcessID);

	/// Write the link data to a file
	bool	WriteToFile(LPCTSTR lpFilename);
	/// Read the link data from a file
	bool	ReadFromFile(LPCTSTR lpFilename);

	/// Ensure we have enough pages to put data in the requested page
	void	EnsurePage(int nPage);
};

#endif   //#define _CCPRINTDATA_H_
