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

#ifndef _CCCOMMON_H_
#define _CCCOMMON_H_

/// Generic printer plugin definitions
#define CCPRINT_SIGNATURE			'CCPS'
#define CCPRINT_VERSION				0x00000006L
#define CCPRINT_VERSION_NO_URLS		0x00000005L
#define CCPRINT_VERSION_RUNTIME		0x00000004L
#define CCPRINT_VERSION_NO_TRANS	0x00000003L
#define CCPRINT_VERSION_NO_EXTRA	0x00000002L

/// Registry settings
#define SETTINGS_AUTOOPEN			_T("AutoOpen")
#define SETTINGS_WRITEPROPERTIES	_T("WriteProperties")
#define SETTINGS_LICENSELOCATION	_T("LicenseLocation")
#define SETTINGS_AUTOURLS			_T("AutoURLs")
#define SETTINGS_CREATEASTEMP		_T("CreateAsTemp")


#endif   //#define _CCCOMMON_H_
