/**
	@file
	@brief Precompiled header creator file
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

void SAMTrace(LPCWSTR lpszFormat, ...)
{
#if defined(_DEBUG) || defined(ENABLE_TRACE)
	// Prepare the rest of the parameters (the ellipsis)
	va_list args;
	va_start(args, lpszFormat);

	// Prepare buffer to format the string in
	int nBuf;
	WCHAR szBuffer[1024];

	// Format all parameters inside the string
	nBuf = _vsnwprintf_s (szBuffer, sizeof(szBuffer)/sizeof(szBuffer[0]), lpszFormat, args);

	// was there an error? was the expanded string too long?
	_ASSERT(nBuf >= 0);

	// Send to debug output
	::OutputDebugString(szBuffer);

	// Cleanup variables
	va_end(args);
#endif
}

