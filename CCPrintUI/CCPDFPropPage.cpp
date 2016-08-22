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
#include "CCPDFPropPage.h"
#include "globals.h"

#include "helpers.h"

/**
	@param uMsg ID of the message
	@param wParam First parameter of the message
	@param lParam Second parameter of the message
	@return TRUE if handled, FALSE to continue handling the message
*/
BOOL CCPDFPropPage::PageProc(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	// Which message is this?
    switch (uMsg)
    {
		case WM_INITDIALOG:
			// Page creation:
			// Initialize page controls
			InitControls();
			return FALSE;

        case WM_COMMAND:
            switch(HIWORD(wParam))
            {
                case BN_CLICKED:
					// Some button/box clicked:
                    switch(LOWORD(wParam))
                    {
						case IDC_AUTOOPEN:
							// User changed the open-on-print option
							SetChanged();
							if (!GetDlgItemCheck(IDC_AUTOOPEN))
								SetDlgItemCheck(IDC_TEMP, false);
							break;
                    }
                    break;

				default:
					// Nothing we can handle
                    return FALSE;
            }
			// We got here, meaning we handled it
            return TRUE;

		case WM_NOTIFY:
            {
				// Notification, what is it?
                switch (((LPNMHDR)lParam)->code)  // type of notification message
                {
                    case PSN_SETACTIVE:
						// Page made active, reload the data
						InitControls();
                        break;
    
                    case PSN_KILLACTIVE:
						// OK, get the data
						m_pDevMode->bAutoOpen = GetDlgItemCheck(IDC_AUTOOPEN);
						m_pDevMode->bAutoURLs = GetDlgItemCheck(IDC_AUTOURL);
						m_pDevMode->bCreateAsTemp = GetDlgItemCheck(IDC_TEMP);
						return TRUE;

                    case PSN_APPLY:
						// Apply, we want to save the data
						m_pDevMode->bAutoOpen = GetDlgItemCheck(IDC_AUTOOPEN);
						m_pDevMode->bAutoURLs = GetDlgItemCheck(IDC_AUTOURL);
						m_pDevMode->bCreateAsTemp = GetDlgItemCheck(IDC_TEMP);
						// Notify the page handle that it applies
						(*m_pfnComPropSheet)(m_hComPropSheet, CPSFUNC_SET_RESULT, (LPARAM)m_hPage, CPSUI_OK);
						return TRUE;

                    case PSN_RESET:
						// Move along, nothing to see here
                        break;
                }
            }
            break;
	}

	// Got here: could not handle it
	return FALSE;
}

/**
	
*/
void CCPDFPropPage::InitControls()
{
	// Set the page controls
	SetDlgItemCheck(IDC_AUTOOPEN, m_pDevMode->bAutoOpen);
	SetDlgItemCheck(IDC_AUTOURL, m_pDevMode->bAutoURLs);
	SetDlgItemCheck(IDC_TEMP, m_pDevMode->bCreateAsTemp);

	// Do we have a handle for PDF files?
	if (!CanOpenPDFFiles())
	{
		// No, so disable the option to auto-open them
		SetDlgItemCheck(IDC_AUTOOPEN, FALSE);
		EnableDlgItem(IDC_AUTOOPEN, FALSE);
		// And if you can't auto-open, then temporary files are out of the question
		SetDlgItemCheck(IDC_TEMP, FALSE);
		EnableDlgItem(IDC_TEMP, FALSE);
	}
}
