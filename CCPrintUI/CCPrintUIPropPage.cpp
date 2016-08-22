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
#include "CCPrintUIPropPage.h"

/**
	@param uResourceID ID of the page resource
	@param hPrinter Handle to the printer
	@param pHelper Pointer to the Print UI Core object
	@param pDevMode Pointer to the CCPrinter data structure (can be NULL)
*/
CCPrintUIPropPage::CCPrintUIPropPage(UINT uResourceID, HANDLE hPrinter, IPrintOemDriverUI* pHelper, POEMDEV pDevMode /* = NULL */) : CCPrintPropPage(uResourceID, pHelper), m_pDevMode(pDevMode), m_hPrinter(hPrinter)
{
}
