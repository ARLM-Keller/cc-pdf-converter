/**
	@file
	@brief include file for standard system include files,
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

#if !defined(AFX_STDAFX_H__FD0C5ED7_D98D_4F65_BABF_87BD46E20095__INCLUDED_)
#define AFX_STDAFX_H__FD0C5ED7_D98D_4F65_BABF_87BD46E20095__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
#define _ATL_APARTMENT_THREADED

#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>

// This is taken from the appropriate Office folder
#pragma warning (push)
#pragma warning (disable: 4146)
#import "mso9.dll" rename_namespace("Office2000")
#pragma warning (pop)
using namespace Office2000;
// This is taken from the COMMON FILES folder
#import "Microsoft Shared\\VBA\\VBA6\\VBE6EXT.olb" rename_namespace("VBE6")
using namespace VBE6;

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__FD0C5ED7_D98D_4F65_BABF_87BD46E20095__INCLUDED)
