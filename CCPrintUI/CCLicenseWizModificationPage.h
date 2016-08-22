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

#ifndef _CCLICENSEWIZMODIFICATIONPAGE_H_
#define _CCLICENSEWIZMODIFICATIONPAGE_H_

#include "CCLicenseWizBasePage.h"
#include "resource.h"

/**
    @brief License Wizard: Modification page
*/
class CCLicenseWizModificationPage : public CCLicenseWizBasePage
{
public:
	// Ctor
	/**
		@brief Constructor
		@param pHelper Pointer to the Print UI Core object
		@param pInfo Pointer to the current license information structure
	*/
	CCLicenseWizModificationPage(IPrintOemDriverUI* pHelper, LicenseInfo* pInfo) : CCLicenseWizBasePage(IDD_LICENSEPAGE_MODIFICATION, pHelper, pInfo) {};

public:
	// CCPrintPropPage public overrides
	/// Called to fill the wizard page structure
	virtual bool PreparePage(PROPSHEETPAGE& page);

protected:
	// CCPrintPropPage protected overrides
	/// Called when a windows message is sent to the page
	virtual BOOL	PageProc(UINT uMsg, WPARAM wParam, LPARAM lParam);

	// CCPrintWizBasePage overrides
	/// Called when the page is being activated (displayed)
	virtual BOOL	OnSetActive();
	/// Called when the page is being deactivated (hidden)
	virtual BOOL	OnKillActive();

	// Helpers
	/// Updates the page controls according to the selection
	void			UpdateControls();
};

#endif   //#define _CCLICENSEWIZMODIFICATIONPAGE_H_
