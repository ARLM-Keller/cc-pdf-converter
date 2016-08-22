/**
	@file
	@brief Define function prototypes for Print Plugin UI Module
			based on oemui.h
			Printer Driver UI Plugin Sample
			by Microsoft Corporation
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

#ifndef _OEMUI_H
#define _OEMUI_H

#include "CCPrintUI.H"
#include "DEVMODE.H"
#include "globals.h"


////////////////////////////////////////////////////////
//      Prototypes
////////////////////////////////////////////////////////

/// Initializes OptItems to display the device or document property UI
HRESULT hrOEMPropertyPage(DWORD dwMode, POEMCUIPPARAM pOEMUIParam);
/// Document property page callback function
HRESULT hrOEMDocumentPropertySheets(PPROPSHEETUI_INFO pPSUIInfo, LPARAM lParam, struct IPrintOemDriverUI* pOEMHelp);
/// Device property page callback function
HRESULT hrOEMDevicePropertySheets(PPROPSHEETUI_INFO pPSUIInfo, LPARAM lParam, struct IPrintOemDriverUI* pOEMHelp);

#endif
