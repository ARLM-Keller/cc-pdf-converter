/**
	@file
	@brief Implementation of CCCPDFExcelAddinObj
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

#include "stdafx.h"
#ifdef CC_PDF_CONVERTER
#include "CCPDFExcelAddin.h"
#elif EXCEL_TO_PDF
#include "XL2PDFExcelAddin.h"
#else
#error "Please define one of the printer types"
#endif
#include "CCPDFExcelAddinObj.h"

#include "CCPrintData.h"
#include "CCRegistry.h"

/////////////////////////////////////////////////////////////////////////////
// CCCPDFExcelAddinObj

#ifdef CC_PDF_CONVERTER
/// Product name (for error messages)
#define PRODUCT_NAME _T("CC PDF Converter")
/// Printer name (for error messages)
#define PRINTER_NAME _T("CC PDF Printer")
/// Registry path for Addin data
#define REGISTRY_DATA_PATH	_T("Software\\Cogniview\\CC PDF Converter\\Excel Addin")
/// Printer driver name
#define PRINTER_DRIVER_NAME _T("CC PDF Virtual Printer")
/// Excel menu/toolbar tag
#define BUTTON_TAG		_T("CCPDF Save PDF")
#elif EXCEL_TO_PDF
/// Product name (for error messages)
#define PRODUCT_NAME _T("Excel to PDF")
/// Printer name (for error messages)
#define PRINTER_NAME _T("Excel to PDF Printer")
/// Registry path for Addin data
#define REGISTRY_DATA_PATH	_T("Software\\Cogniview\\Excel to PDF Converter\\Excel Addin")
/// Printer driver name
#define PRINTER_DRIVER_NAME _T("Excel to PDF Virtual Printer")
/// Excel menu/toolbar tag
#define BUTTON_TAG		_T("XL2PDF Save PDF")
#else
#error "Please define one of the printer types"
#endif

/// Find printer error message
#define FIND_PRINTER_ERROR_MSG _T("Cannot access the ") PRINTER_NAME
/// Title for error message dialogs
#define PRINTER_ERROR_TITLE PRODUCT_NAME
/// Access printer error message
#define ACCESS_PRINTER_ERROR_MSG _T("Error when accessing the ") PRINTER_NAME
/// Working with printer error message
#define WORKING_PRINTER_ERROR_MSG _T("Error when working with the ") PRINTER_NAME

/// Excel menu item/toolbar button callback definition
typedef IDispEventSimpleImpl</*nID =*/ 2, CCCPDFExcelAddinObj, &__uuidof(Office2000::_CommandBarButtonEvents)> CommandButtonEvents;
/// Excel's document open notification definition
_ATL_FUNC_INFO CCCPDFExcelAddinObj::DocumentOpenInfo = {CC_STDCALL,VT_EMPTY, 1, {VT_DISPATCH|VT_BYREF}};
/// Excel's new document notification definition
_ATL_FUNC_INFO CCCPDFExcelAddinObj::DocumentNew = {CC_STDCALL,VT_EMPTY, 1, {VT_DISPATCH|VT_BYREF}};
/// Excel's closing document notification definition
_ATL_FUNC_INFO CCCPDFExcelAddinObj::DocumentBeforeCloseInfo = {CC_STDCALL,VT_EMPTY, 2, {VT_DISPATCH|VT_BYREF, VT_BYREF|VT_BOOL}};
/// Excel's button click notification
_ATL_FUNC_INFO CCCPDFExcelAddinObj::OnClickButtonInfo ={CC_STDCALL,VT_EMPTY,2,{VT_DISPATCH,VT_BYREF | VT_BOOL}};

/**
	@param riid The interface to check
	@return S_OK if interface is supported, S_FALSE if not
*/
STDMETHODIMP CCCPDFExcelAddinObj::InterfaceSupportsErrorInfo(REFIID riid)
{
	// Define the interfaces we support
	static const IID* arr[] = 
	{
#ifdef CC_PDF_CONVERTER
		&IID_ICCPDFExcelAddinObj
#elif EXCEL_TO_PDF
		&IID_IXL2PDFExcelAddinObj
#else
#error "Please define one of the printer types"
#endif
	};

	// Check if we support the requested interface
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (IsEqualGUID(*arr[i],riid))
			// Yep
			return S_OK;
	}
	// Nope
	return S_FALSE;
}

/**
	@param Application Excel application interface object
	@param ConnectMode Type of connection to Excel (not used)
	@param AddInInst Addin instance interface object (not used)
	@param custom Additional data (not used)
	@return S_OK if all went well, error code otherwise
*/
STDMETHODIMP CCCPDFExcelAddinObj::OnConnection(IDispatch * Application, ext_ConnectMode ConnectMode, IDispatch * AddInInst, SAFEARRAY * * custom)
{
	HRESULT hr;

	// Remember the application object
	CComQIPtr<MSExcel::_Application> spApp(Application);
	ATLASSERT(spApp);
	m_spApp = spApp;

	// Add the toolbar button, if not found:
	CComPtr <_CommandBars> pCmdBars = m_spApp->GetCommandBars();
	CComPtr <CommandBar> pStandardBar = NULL;
	int nBefore;

	// Get the standard toolbar
	hr = pCmdBars->get_Item(_variant_t(_T("Standard")), &pStandardBar);
	if (!FAILED(hr) && (pStandardBar != NULL))
	{
		// Try to find the button, or the location if not found
		CComPtr <CommandBarControls> pBarControls = pStandardBar->GetControls();

		bool bFound = false;

		// Check existing buttons:
		int nCount = pBarControls->GetCount();
		nBefore = nCount;
		for (int i=nCount; i>=1; i--)
		{
			CComPtr<CommandBarControl> pControl = pBarControls->GetItem(_variant_t((long)i));

			CComQIPtr<Office2000::_CommandBarButton> pButton(pControl);
			if (pButton == NULL)
				// Not a button, go on
				continue;

			// Is this our button?
			_bstr_t tag = pButton->GetTag();
			if (pButton->GetId() == 3)
			{
				// This is the Save, I hope
				nBefore = i + 1;
			}

			if (_tcscmp(BUTTON_TAG, (LPCTSTR)tag) == 0)
			{
				// This is our button - keep reference to button disp
				m_buttonToolbar = pButton;
				bFound = true;
				break;
			}
		}

		if (!bFound)
		{
			// We need to add a button:
			CComVariant vToolBarType(Office2000::msoControlButton);
			CComVariant vEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR); 
			CComVariant vBefore(nBefore);

			// Add button
			CComPtr<CommandBarControl> pControl = pBarControls->Add(vToolBarType, vEmpty, vEmpty, vBefore, vEmpty); 
			ATLASSERT(pControl);

			// create
			CComQIPtr<_CommandBarButton> pButton(pControl);
			ATLASSERT(pButton);

			// set the button's image
			HBITMAP hBitmap = (HBITMAP)::LoadImage(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDB_BUTTON), IMAGE_BITMAP, 0, 0, LR_LOADMAP3DCOLORS);
			if (hBitmap != NULL)
			{
				::OpenClipboard(NULL);
				::EmptyClipboard();
				::SetClipboardData(CF_BITMAP, (HANDLE)hBitmap);
				::CloseClipboard();
				::DeleteObject(hBitmap);
				pButton->PutStyle(Office2000::msoButtonIcon);
				hr = pButton->PasteFace();
				if (FAILED(hr))
					return hr;
			}
			else
			{
				// Didn't work, just use a caption
				pButton->PutStyle(Office2000::msoButtonCaption);
			}

			// Set the button data
			pButton->PutCaption(_T("Save As PDF File...")); 
			pButton->PutTooltipText(_T("Save the current sheet as a PDF file"));
			pButton->PutEnabled(VARIANT_TRUE);
			pButton->PutVisible(VARIANT_TRUE); 

			// enable tag
			pButton->PutTag(BUTTON_TAG);
			// keep reference to button disp
			m_buttonToolbar = pButton;
		}

		// Call us when clicked!
		hr = CommandButtonEvents::DispEventAdvise((IDispatch*)m_buttonToolbar);
		if (FAILED(hr))
			return hr;

		// Enable the button
		m_buttonToolbar->PutEnabled(VARIANT_TRUE);
	}

	// Add to the menu bar, if not there:
	CComPtr <CommandBar> pMenuBar = NULL;
	hr = pCmdBars->get_Item(_variant_t(_T("Worksheet Menu Bar")), &pMenuBar);
	if (!FAILED(hr) && (pMenuBar != NULL))
	{
		// Test if the button's already there
		CComVariant vToolBarType(Office2000::msoControlButton);
		CComVariant vEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR); 

		CComPtr <CommandBarControls> pBarControls = pMenuBar->GetControls();
		CComPtr<CommandBarControl> pControl = pBarControls->GetItem(_T("File"));
		if (pControl != NULL)
		{
			CComQIPtr<Office2000::CommandBarPopup> pPopup(pControl);
			if (pPopup != NULL)
			{
				// OK, go over the menus:
				CComPtr <CommandBarControls> pMenuControls = pPopup->GetControls();
				long lBefore = 1;
				long i;
				for (i = 1; i <= pMenuControls->GetCount(); i++)
				{
					// Get item
					CComPtr<CommandBarControl> pControl = pMenuControls->GetItem(i);
					CComQIPtr<Office2000::_CommandBarButton> pButton(pControl);
					if (pButton == NULL)
						// Not a button, go on
						continue;

					if (pButton->GetId() == 3823)
					{
						// This is the menu item we'll be added after
						lBefore = i;
						continue;
					}
					
					// Is this OUR item?
					_bstr_t tag = pButton->GetTag();
					if (_tcscmp(BUTTON_TAG, (LPCTSTR)tag) == 0)
					{
						// Yes, keep for later reference
						m_buttonMenu = pButton;
						break;
					}
				}

				// Did we get it?
				if (i > pMenuControls->GetCount())
				{
					// No, so add a new menu item
					CComVariant vToolBarType(Office2000::msoControlButton);
					CComVariant vEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR); 
					CComVariant vBefore(lBefore + 1);

					// Add menu item
					CComPtr<CommandBarControl> pControl = pMenuControls->Add(vToolBarType, vEmpty, vEmpty, vBefore, vEmpty); 
					ATLASSERT(pControl);

					// create
					CComQIPtr<_CommandBarButton> pButton(pControl);
					ATLASSERT(pButton);
	
					// Try to set up the image
					HBITMAP hBitmap = (HBITMAP)::LoadImage(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDB_BUTTON), IMAGE_BITMAP, 0, 0, LR_LOADMAP3DCOLORS);
					if (hBitmap != NULL)
					{
						::OpenClipboard(NULL);
						::EmptyClipboard();
						::SetClipboardData(CF_BITMAP, (HANDLE)hBitmap);
						::CloseClipboard();
						::DeleteObject(hBitmap);
						pButton->PutStyle(Office2000::msoButtonIconAndCaption);
						hr = pButton->PasteFace();
					}
					else
						// Failed, use text mode
						pButton->PutStyle(Office2000::msoButtonCaption);

					// Set up the button
					pButton->PutCaption(_T("Save As PDF &File...")); 
					pButton->PutTooltipText(_T("Save the current sheet as a PDF file"));
					pButton->PutEnabled(VARIANT_TRUE);
					pButton->PutVisible(VARIANT_TRUE); 

					// enable tag
					pButton->PutTag(BUTTON_TAG);
					// keep reference to button disp
					m_buttonMenu = pButton;
				}
			}
		}

		m_buttonMenu->PutEnabled(VARIANT_TRUE);
	}

	// Set up to notify us when events occur
	hr = IDispEventSimpleImpl<1, CCCPDFExcelAddinObj, &__uuidof(MSExcel::AppEvents)>::DispEventAdvise(m_spApp);
	if (FAILED(hr))
		return hr;

	return S_OK;
}

/**
	@param RemoveMode Type of disconnection from Excel (not used)
	@param custom Additional data (not used)
	@return S_OK if all went well, error code otherwise
*/
STDMETHODIMP CCCPDFExcelAddinObj::OnDisconnection(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
{
	// We don't want any more notifications on the button
	HRESULT hr = CommandButtonEvents::DispEventUnadvise((IDispatch*)m_buttonToolbar);
	if (FAILED(hr))
		return hr;
	
	// We don't want any more notifications from Excel
	hr = IDispEventSimpleImpl<1, CCCPDFExcelAddinObj, &__uuidof(MSExcel::AppEvents)>::DispEventUnadvise(m_spApp);
	if (FAILED(hr))
		return hr;

	// Did the user remove the Addin?
	if (RemoveMode == ext_dm_UserClosed)
	{
		// Yes, so remove the buttons
		m_buttonToolbar->Delete(VARIANT_FALSE);
		m_buttonMenu->Delete(VARIANT_FALSE);
	}

	m_spApp = NULL;
	return hr;
}

/**
	@param workbook The workbook being activated
*/
void __stdcall CCCPDFExcelAddinObj::OnWorkbookActivate(IDispatch * workbook)
{
	// At least one wookbook: enable button
	EnableButtons(true);
}

/**
	@param workbook The workbook being added
*/
void __stdcall CCCPDFExcelAddinObj::OnNewWorkbook(IDispatch * workbook)
{
	// At least one wookbook: enable button
	EnableButtons(true);
}

/**
	@param workbook The workbook being closed
	@param CancelDefault Pointer to cancel variable
*/
void __stdcall CCCPDFExcelAddinObj::OnWorkbookBeforeClose(IDispatch * workbook, VARIANT_BOOL * CancelDefault)
{
	// We want to enable the button if we have at least one workbook:LEFT after the close
	CComPtr <MSExcel::Workbooks> books = m_spApp->GetWorkbooks();
	long lCount = books->GetCount();
	EnableButtons(lCount > 1);
}

/**
	@param sPrinter Name of the printer to use
*/
void CCCPDFExcelAddinObj::SetPrinterAndBreaks(const std::tstring& sPrinter)
{
	// Set the active printer
	m_spApp->PutActivePrinter(0, sPrinter.c_str());

	// Recalculate page breaks by changing views
	CComPtr <MSExcel::Window> window = m_spApp->GetActiveWindow();
	window->PutView(MSExcel::xlPageBreakPreview);
	window->PutView(MSExcel::xlNormalView);
}

/**
	@param button Interface of the button
	@param CancelDefault Pointer to cancel variable
*/
void __stdcall CCCPDFExcelAddinObj::OnButtonClick(IDispatch * button, VARIANT_BOOL * CancelDefault)
{
	// Find a printer we can use:
	std::tstring sPrinter, sPrinterName, sOldPrinter;
	DWORD dwSize, dwCount;
	// Enumerate printers: calculate size
	EnumPrinters(PRINTER_ENUM_LOCAL, NULL, 2, NULL, 0, &dwSize, &dwCount);
	if (dwSize > 0)
	{
		// OK, create an appropriately-sized buffer
		PRINTER_INFO_2* pInfo = (PRINTER_INFO_2*)new char[dwSize];
		// And read the info
		if (EnumPrinters(PRINTER_ENUM_LOCAL, NULL, 2, (LPBYTE)pInfo, dwSize, &dwSize, &dwCount) > 0)
		{
			DWORD i;
			for (i=0;i<dwCount;i++)
			{
				// Is this it?
				if (_tcscmp(pInfo[i].pDriverName, PRINTER_DRIVER_NAME) == 0)
					break;
			}
			if (i < dwCount)
				// Ah, found it
				sPrinterName = pInfo[i].pPrinterName;
		}
		delete [] pInfo;
	}

	// Did we find a printer?
	if (sPrinterName.empty())
	{
		// Nope
		MessageBox(NULL, FIND_PRINTER_ERROR_MSG, PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
		return;
	}
	// Get the port name - we must use the correct one for some wierd Office reason
	bool bSuccess;
	std::tstring sPort = CCRegistry::GetString(HKEY_CURRENT_USER, _T("Software\\Microsoft\\Windows NT\\CurrentVersion\\Devices"), sPrinterName.c_str(), &bSuccess);
	if (!bSuccess || sPort.empty())
	{
		MessageBox(NULL, FIND_PRINTER_ERROR_MSG, PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
		return;
	}
	std::tstring::size_type pos = sPort.find_first_of(',');
	if (pos == std::tstring::npos)
	{
		MessageBox(NULL, FIND_PRINTER_ERROR_MSG, PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
		return;
	}
	sPrinter = sPrinterName + _T(" on ") + sPort.substr(pos + 1);

	// Set it as the current printer
	sOldPrinter = m_spApp->GetActivePrinter();
	if (sPrinter != sOldPrinter)
		SetPrinterAndBreaks(sPrinter);

	// Get printer handle - we'll need it later
	HANDLE hPrinter = NULL;
	PRINTER_DEFAULTS defs;
	defs.pDatatype = NULL;
	defs.pDevMode = NULL;
	defs.DesiredAccess = PRINTER_ACCESS_ADMINISTER;
	if (!OpenPrinter((LPTSTR)sPrinterName.c_str(), &hPrinter, &defs))
	{
		// Can't do it for some reason
		MessageBox(NULL, FIND_PRINTER_ERROR_MSG, PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
		if (sPrinter != sOldPrinter)
			SetPrinterAndBreaks(sOldPrinter);
		return;
	}

	// OK, run the job
	m_hDialog = NULL;
	try
	{
		DoPrint(hPrinter, sPrinter);
	}
	catch( _com_error e )
    {
		// Failed - com error
		CleanupDialog();
        MessageBox(NULL, e.ErrorMessage(), PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
    }
	catch(...)
	{
		// Failed - something else
		CleanupDialog();
		MessageBox(NULL, _T("Unspecified error"), PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
	}

	// OK, clean up
	CCPrintData data;
	data.CleanSaved(hPrinter);
	::ClosePrinter(hPrinter);

	// Return to the original printer, if necessary
	if (sPrinter != sOldPrinter)
		SetPrinterAndBreaks(sOldPrinter);
}

/**
	@brief Recursively enables or disables the addin buttons in a menu or toolbar
	@param pControls Interface of command bar to go over
	@param bEnable true to enable the addin buttons, false to disable them
	@return true if all when well, false if something happened
*/
bool EnableButtons(CComPtr<CommandBarControls>& pControls, bool bEnable)
{
	// Go over all the toolbar/menu controls
	for (long i=1;i<pControls->GetCount();i++)
	{
		// Is this a regular control?
		CComPtr<CommandBarControl> pControl = pControls->GetItem(i);
		switch (pControl->GetType())
		{
			case Office2000::msoControlButton:
				{
					// A button, get a reference to it
					CComQIPtr<Office2000::_CommandBarButton> pButton(pControl);
					// Check if this is OUR button
					_bstr_t tag = pButton->GetTag();
					if (_tcscmp(BUTTON_TAG, (LPCTSTR)tag) == 0)
						pButton->PutEnabled(bEnable ? VARIANT_TRUE : VARIANT_FALSE);
				}
				break;
			default:
				{
					// Is this a popup menu?
					CComQIPtr<Office2000::CommandBarPopup> pPopup(pControl);
					if (pPopup != NULL)
					{
						// Yes, get interface
						CComPtr <CommandBarControls> pInnerControls = pPopup->GetControls();
						// And recurse
						if (!EnableButtons(pInnerControls, bEnable))
							return false;
					}
				}
				break;
		}
	}
	return true;
}

/**
	@param bEnable true to enable the addin buttons, false to disable them
	@return true if all when well, false if something happened
*/
bool CCCPDFExcelAddinObj::EnableButtons(bool bEnable)
{
	// Start with all the command bars
	CComPtr <_CommandBars> pCmdBars = m_spApp->GetCommandBars();

	// Go over the bars
	for (long i=1;i<pCmdBars->GetCount();i++)
	{
		CComPtr <CommandBar> pCommandBar = NULL;
		HRESULT hr = pCmdBars->get_Item((variant_t)i, &pCommandBar);
		if (!FAILED(hr) && (pCommandBar != NULL))
		{
			// OK, get the controls array
			CComPtr <CommandBarControls> pControls = pCommandBar->GetControls();
			// Enable if found:
			if (!::EnableButtons(pControls, bEnable))
				return false;
		}
	}
	return true;
}

/**
    @brief Data about a sheet's row or column (page, size, location)
*/
struct CellInfo
{
public:
	/**
		Default ctor
	*/
	CellInfo() : nPage(0), dPos(0.0), dEnd(0.0) {};
	/**
		@brief Constructor
		@param n Page number
		@param d Location within the page
		@param dS Size (height of row, width of column)
	*/
	CellInfo(int n, double d, double dS) : nPage(n), dPos(d), dEnd(d + dS) {};
	/**
		@brief Copy constructor
		@param other structure to copy
	*/
	CellInfo(const CellInfo& other) : nPage(other.nPage), dPos(other.dPos), dEnd(other.dEnd) {};

	/// The page into which this row or column is printed
	int		nPage;
	/// The location in the page
	double	dPos;
	/// The end location in the page
	double	dEnd;

	/**
		@brief Data setting function
		@param n Page number
		@param d Location within the page
		@param dS Size (height of row, width of column)
	*/
	void	Set(int n, double d, double dS) {nPage = n; dPos = d; dEnd = d + dS;};
};

/**
	@brief Calculate the Excel 'name' for a cell
	@param nRow The row number (1-based)
	@param nColumn The column number (1-based)
	@return Excel 'name' for the cell (for example, 'A1' or 'BA128')
*/
std::tstring CellName(long nRow, long nColumn)
{
	std::tstring sColumn;

	// Create column name (with letters)
	do 
	{
		nColumn--;
		int nLetter = nColumn % 26;
		// Get the letter
		TCHAR cLetter = _T('A') + nLetter;
		sColumn = cLetter + sColumn;
		// Divide by 26 to get the rest of the letter
		nColumn /= 26;
	} while (nColumn > 0);

	// Now print the name
	TCHAR cRet[1024];
	_stprintf(cRet, _T("%.768s%d"), sColumn.c_str(), nRow);
	return cRet;
}

/**
    @brief Page setup information (either horizontal or vertical)
*/
struct PrintPageData
{
	/**
		@brief Constructor
		@param lFont The size of the 'normal' style font
		@param sFont The name of the 'normal' style font
	*/
	PrintPageData(long lFont, bstr_t& sFont) : dDPI(0.0), dStartMargin(0.0), dEndMargin(0.0), bCenter(false), lFontSize(lFont), sFontName(sFont) {};
	/// DPI of the page
	double	dDPI;
	/// Print starting position
	double	dStartMargin;
	/// Print ending position
	double	dEndMargin;
	/// true if centered, false if left/top aligned
	bool	bCenter;
	/// Size of 'normal' style font
	long	lFontSize;
	/// Name of 'normal' style font
	std::tstring sFontName;

	/// Returns a descriptive data name (for saving/loading purposes)
	std::tstring CreateName() const;
};

/**
	@return A descriptive name for the object
*/
std::tstring PrintPageData::CreateName() const
{
	TCHAR c[128];
	_stprintf(c, _T("%d%0.2f%0.2f%d%d%s"), (long)dDPI, dStartMargin, dEndMargin, bCenter ? 1 : 0, lFontSize, sFontName.c_str());
	return c;
}

/**
    @brief Calculated factors for the current PDF printing (either horizontal or vertical)
*/
struct PrintCalcData
{
	/**
		Default constructor
	*/
	PrintCalcData() : dFactor(0.0), lOffset(0), lFullSize(0), bCenter(false), bForceDPI(false) {};
	/// The factor to multiple Excel's (wrong) location values with
	double	dFactor;
	/// The offset to start with
	long	lOffset;
	/// Full size of the page
	long	lFullSize;
	/// true if printing is centered, false if not
	bool	bCenter;
	/// true to use the default DPI setting (300) instead of the original DPI
	bool	bForceDPI;

	/// Write the calculated data to the registry
	bool	WriteToRegistry(const PrintPageData& orig, LPCTSTR lpPrefix);
	/// Read the calculated data from the registry
	bool	ReadFromRegistry(const PrintPageData& orig, LPCTSTR lpPrefix);

	/// Calculates the factors
	void	FromData(const PrintPageData& orig, int nEnd, int nStart, double dEnd, double dStart, long lFull);
	/// Calculates the PDF location from the Excel location according to the factors
	long	FixLocation(double dPos, double dFullPage);
};

/**
	@param orig The matching page setup data
	@param lpPrefix Prefix to use for registry key
	@return true if read successfully, false if failed
*/
bool PrintCalcData::ReadFromRegistry(const PrintPageData& orig, LPCTSTR lpPrefix)
{
	// Get key name
	std::tstring sName = REGISTRY_DATA_PATH _T("\\PD_");
	sName += lpPrefix + orig.CreateName();

	// Read from (regular!) registry
	int nSize = sizeof(dFactor);
	unsigned char* pFactor;
	if (!CCRegistry::GetBinary(HKEY_CURRENT_USER, sName.c_str(), _T("Factor"), pFactor, nSize))
		return false;
	dFactor = *((double*)pFactor);
	delete [] pFactor;
	bool bSuccess;
	DWORD dw = CCRegistry::GetNumber(HKEY_CURRENT_USER, sName.c_str(), _T("Offset"), &bSuccess);
	if (!bSuccess)
		return false;
	lOffset = (long)dw;
	dw = CCRegistry::GetNumber(HKEY_CURRENT_USER, sName.c_str(), _T("FullSize"), &bSuccess);
	if (!bSuccess)
		return false;
	lFullSize = (long)dw;
	dw = CCRegistry::GetNumber(HKEY_CURRENT_USER, sName.c_str(), _T("ForceDPI"), &bSuccess);
	if (bSuccess)
		bForceDPI = (dw == 1);

	// Remember centering data
	bCenter = orig.bCenter;
	return true;
}

/**
	@param orig The matching page setup data
	@param lpPrefix Prefix to use for registry key
	@return true if saved successfully, false if failed
*/
bool PrintCalcData::WriteToRegistry(const PrintPageData& orig, LPCTSTR lpPrefix)
{
	// Get key name
	std::tstring sName = REGISTRY_DATA_PATH _T("\\PD_");
	sName += lpPrefix + orig.CreateName();

	// Write data
	if (!CCRegistry::SetBinary(HKEY_CURRENT_USER, sName.c_str(), _T("Factor"), (unsigned char*)&dFactor, sizeof(dFactor)))
		return false;
	if (!CCRegistry::SetNumber(HKEY_CURRENT_USER, sName.c_str(), _T("Offset"), (DWORD)lOffset))
		return false;
	if (!CCRegistry::SetNumber(HKEY_CURRENT_USER, sName.c_str(), _T("FullSize"), (DWORD)lFullSize))
		return false;
	if (!CCRegistry::SetNumber(HKEY_CURRENT_USER, sName.c_str(), _T("ForceDPI"), bForceDPI ? 1 : 0))
		return false;

	return true;
}

/**
	@param orig The matching page setup data
	@param nEnd Actual print end location
	@param nStart Actual print start location
	@param dEnd Excel's print end location
	@param dStart Excel's print start location
	@param lFull The full width of the page in the print
*/
void PrintCalcData::FromData(const PrintPageData& orig, int nEnd, int nStart, double dEnd, double dStart, long lFull)
{
	// Remember centering and size
	bCenter = orig.bCenter;
	lFullSize = lFull;
	// Calculate the multiplication factor
	dFactor = (nEnd - nStart) / (dEnd - dStart);
	// Calculate the offset (different for centered data)
	lOffset = bCenter ? (nStart + nEnd) / 2 : nStart;
}

/**
	@param dPos The Excel's (alleged) print position
	@param dFullPage Excel's full page size (for centering)
	@return The actual print location
*/
long PrintCalcData::FixLocation(double dPos, double dFullPage)
{
	// Start with the offset
	long lRet = lOffset;
	if (bCenter)
		// Centered, so the location is actually from the middle of the page)
		dPos -= (dFullPage / 2);
	// Calculate correct position
	return lRet + (dPos * dFactor);
}

/**
	@param hwndDialog Handle of dialog
	@param uMsg ID of the message
	@param wParam First message paramenter
	@param lParam Second message paramenter
	@return TRUE if the message was handled, FALSE otherwise
*/
INT_PTR CALLBACK ProcessingDlgFunc(HWND hwndDialog, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	// Do nothing
	return FALSE;
}

/**
	@param hWnd The window to center
	This function will center the window on the screen
*/
void CenterWindow(HWND hWnd)
{
	// get coordinates of the window relative to its parent
	RECT rcDlg;
	::GetWindowRect(hWnd, &rcDlg);
	RECT rcCenter, rcArea;
	// center within screen coordinates
	::SystemParametersInfo(SPI_GETWORKAREA, NULL, &rcArea, NULL);
	rcCenter = rcArea;

	int DlgWidth = rcDlg.right - rcDlg.left;
	int DlgHeight = rcDlg.bottom - rcDlg.top;

	// find dialog's upper left based on rcCenter
	int xLeft = (rcCenter.left + rcCenter.right) / 2 - DlgWidth / 2;
	int yTop = (rcCenter.top + rcCenter.bottom) / 2 - DlgHeight / 2;

	// if the dialog is outside the screen, move it inside
	if (xLeft < rcArea.left)
		xLeft = rcArea.left;
	else if(xLeft + DlgWidth > rcArea.right)
		xLeft = rcArea.right - DlgWidth;

	if(yTop < rcArea.top)
		yTop = rcArea.top;
	else if(yTop + DlgHeight > rcArea.bottom)
		yTop = rcArea.bottom - DlgHeight;

	// map screen coordinates to child coordinates
	::SetWindowPos(hWnd, NULL, xLeft, yTop, -1, -1, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

/**
	@param hPrinter Handle to the printer
	@param sPrinter Name of the printer (in Office, this includes the port)
*/
void CCCPDFExcelAddinObj::DoPrint(HANDLE hPrinter, const std::tstring& sPrinter)
{
	// Get the workbook
	CCPrintData data;
	CComPtr <MSExcel::_Workbook> workbook = m_spApp->GetActiveWorkbook();
	if (workbook == NULL)
	{
		MessageBox(NULL, _T("Cannot access the active worksheet"), PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
		return;
	}

	// Ensure we have the 'normal' style font, as this effects the print offset for some reason
	CComPtr<MSExcel::Styles> styles = workbook->GetStyles();
	CComPtr<MSExcel::Style> style = styles->GetItem(_T("Normal"));
	if (style == NULL)
	{
		MessageBox(NULL, _T("Cannot access the active worksheet's styles"), PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
		return;
	}
	CComPtr<MSExcel::Font> font = style->GetFont();
	long lFontSize = font->GetSize();
	bstr_t sFontName = (bstr_t) font->GetName();
	CComQIPtr<MSExcel::_Worksheet> worksheet = workbook->GetActiveSheet();
	if (worksheet == NULL)
	{
		MessageBox(NULL, _T("Cannot access the active worksheet"), PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
		return;
	}

	// Get the print range...
	CComPtr<MSExcel::Range> range = worksheet->GetUsedRange();
	int nColumns = range->GetColumns()->GetCount() + range->GetColumn(), nRows = range->GetRows()->GetCount() + range->GetRow();
	if ((nColumns < 1) || (nRows < 1))
	{
		// Nothing to print, stop here
		CleanupDialog();
		MessageBox(NULL, _T("No data in the active sheet"), PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
		return;
	}

	// Make sure we don't have a fit-to page settings, or we can't do this:
	CComPtr<MSExcel::PageSetup> page = worksheet->GetPageSetup();
	variant_t vZoom = page->GetZoom();
	bool bFitToPage = ((vZoom.vt == VT_BOOL) && (vZoom.boolVal == FALSE));
	if (bFitToPage)
	{
		// It's fit-to-page, we don't know how to calculate this (yet!)
		MessageBox(NULL, PRODUCT_NAME _T(" cannot make a PDF from a sheet which is set to 'fit-to'.\nPlease change the setting in Excel's Page Setup"), PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
		return;
	}

	// Show 'processing' dialog
	m_hDialog = CreateDialog(_Module.m_hInstResource, MAKEINTRESOURCE(IDD_PROCESSING), NULL, ProcessingDlgFunc);
	CenterWindow(m_hDialog);

	// Get the page settings for margins, DPI and centering
	PrintPageData pageHorz(lFontSize, sFontName), pageVert(lFontSize, sFontName);
	PrintCalcData dataHorz, dataVert;

	pageHorz.dStartMargin = page->GetLeftMargin();
	pageHorz.dEndMargin = page->GetRightMargin();
	pageHorz.dDPI = page->GetPrintQuality((long)1);
	pageHorz.bCenter = ((BOOL)page->GetCenterHorizontally()) ? true : false;

	pageVert.dStartMargin = page->GetTopMargin();
	pageVert.dEndMargin = page->GetBottomMargin();
	pageVert.dDPI = page->GetPrintQuality((long)2);
	pageVert.bCenter = ((BOOL)page->GetCenterVertically()) ? true : false;

	// Don't update the screen, we're working!
	m_spApp->PutScreenUpdating(0, VARIANT_FALSE);
	// Calculdate factors (read them from the registry if possible)
	if (!CalculatePrintData(pageHorz, pageVert, dataHorz, dataVert, data, hPrinter, sPrinter, page, worksheet))
	{
		m_spApp->PutScreenUpdating(0, VARIANT_TRUE);
		return;
	}

	// Update DPI if necessary (when the original DPI is not supported by OUR printer)
	if (dataHorz.bForceDPI)
		page->PutPrintQuality((long)1, 300.0);
	if (dataVert.bForceDPI)
		page->PutPrintQuality((long)2, 300.0);

	// Get the screen update back on
	m_spApp->PutScreenUpdating(0, VARIANT_TRUE);
	// Bring dialog back on top
	EnsureDialogVisible();

	// Do we print down first or right first?
	bool bDownfirst = (page->GetOrder() == MSExcel::xlDownThenOver);

	/* *****************************************
	** Calculating the row and column locations
	*******************************************/
	std::vector<CellInfo> arColumns, arRows;
	std::vector<double> arPageWidth, arPageHeight;
	arColumns.resize(nColumns+1);
	arRows.resize(nRows+1);
	double dSize = 0.0, dPageStart = 0.0;
	int nPage = 1;
	arPageWidth.push_back(0.0);
	CComPtr<MSExcel::Range> columns = worksheet->GetColumns();
	CComPtr<MSExcel::Range> rows = worksheet->GetRows();
	CComQIPtr<MSExcel::Range> temp;
	// Go over columns
	long i;
	for (i=1;i<=nColumns+1;i++)
	{
		// Get the column range object
		temp = columns->GetItem(i);
		// Get the width
		dSize = temp->GetWidth().dblVal;
		if (temp->GetPageBreak() != MSExcel::xlPageBreakNone)
		{
			// This is the first column in a new page
			dPageStart = temp->GetLeft();
			nPage++;
			// Add new page to width array
			arPageWidth.push_back(0.0);
		}
		// Remember the column location and width
		arColumns[i-1].Set(nPage, temp->GetLeft().dblVal - dPageStart, dSize);
		// Add this column's width to the page width
		arPageWidth[nPage - 1] += temp->GetWidth().dblVal;
	}
	// Remove the last column's width as this column is not used
	arPageWidth[nPage - 1] -= temp->GetWidth().dblVal;

	// Go over rows
	dSize = 0.0;
	dPageStart = 0.0;
	nPage = 1;
	arPageHeight.push_back(0.0);
	for (i=1;i<=nRows+1;i++)
	{
		// Get the row range object
		temp = rows->GetItem(i);
		// Get the height
		dSize = temp->GetHeight().dblVal;
		if (temp->GetPageBreak() != MSExcel::xlPageBreakNone)
		{
			// This is the first row in a new page
			dPageStart = temp->GetTop();
			nPage++;
			// Add new page to the height array
			arPageHeight.push_back(0.0);
		}
		// Remember the row location and height
		arRows[i-1].Set(nPage, temp->GetTop().dblVal - dPageStart, dSize);
		// Add this row's height to the page height
		arPageHeight[nPage - 1] += temp->GetHeight().dblVal;
	}
	// Remove the last row's height as this row is not used
	arPageHeight[nPage - 1] -= temp->GetHeight().dblVal;

	// Remember sheet name for internal links
	std::tstring sSheetName;
	LPCTSTR lpName = worksheet->GetName();
	if (lpName != NULL)
		sSheetName = worksheet->GetName();

	// Get data for links
	CComPtr<MSExcel::Hyperlinks> links = worksheet->GetHyperlinks();
	CComPtr<MSExcel::Names> names = workbook->GetNames();
	RECTL rect;

	// Run over the list of links and create the data
	for (i=1;i<=links->GetCount();i++)
	{
		// Get link
		CComPtr<MSExcel::Hyperlink> link = links->GetItem(i);
		try
		{
			// Does it have a location?
			temp = link->GetRange();
		}
		catch (_com_error e)
		{
			// No range: just continue
			continue;
		}

		// Get row and column location for this location
		const CellInfo& row = arRows[temp->GetRow() - 1], column = arColumns[temp->GetColumn() - 1];
		
		// Get the corrected printing location of the link
		rect.top = dataVert.FixLocation(row.dPos, arPageHeight[row.nPage - 1]);
		rect.left = dataHorz.FixLocation(column.dPos, arPageWidth[column.nPage - 1]);
		rect.bottom = dataVert.FixLocation(row.dEnd, arPageHeight[row.nPage - 1]);
		rect.right = dataHorz.FixLocation(column.dEnd, arPageWidth[column.nPage - 1]);

		// Get the page in which this cell will be printed
		nPage = bDownfirst ? (arRows.back().nPage * (column.nPage - 1)) + row.nPage : (arColumns.back().nPage * (row.nPage - 1)) + column.nPage;

		// OK, find the address
		std::tstring sAddress, sSubAddress;
		LPCTSTR lpAddress = link->GetAddress();
		LPCTSTR lpSubAddress = link->GetSubAddress();
		if (lpAddress != NULL)
			sAddress = link->GetAddress();
		if (lpSubAddress != NULL)
			sSubAddress = link->GetSubAddress();

		if (!sAddress.empty())
		{
			// Regular (outside) link:
			if (!sSubAddress.empty())
				sAddress += _T("#") + sSubAddress;

			// Add to link list
			data.AddLink(sAddress, rect, nPage, (LPCTSTR)link->GetScreenTip());
			continue;
		}

		// This is some kind of internal link, see if it's to this sheet
		if (sSubAddress.empty() || sSheetName.empty())
			continue;
		// Internal link... where does it go?
		try
		{
			// Is this a named link?
			CComPtr<MSExcel::Name> name = names->Item(sSubAddress.c_str());
			if (name != NULL)
			{
				// Yep, get the address it points to
				variant_t var = name->GetRefersToLocal();
				sSubAddress = (_bstr_t)var;
				if (!sSubAddress.empty())
				{
					if (sSubAddress[0] == '=')
						sSubAddress.erase(0, 1);
					else
						sSubAddress = _T("");
				}
			}
		}
		catch (...)
		{
			// Probably not a name, so lets continue
		}
		if (sSubAddress.empty())
			// Nothing to see, move along
			continue;

		// Link is defined as: [<SheetName>!][$]<Column>[$]<Row>

		// Find if this is linked to THIS sheet:
		std::tstring::size_type nBreak = sSubAddress.find_last_of('!');
		if (nBreak != std::tstring::npos)
		{
			// There's a sheet name, is this it?
			std::tstring sSheet = sSubAddress.substr(0, nBreak);
			if (sSheet != sSheetName)
				continue;
			// Yes, remove the name
			sSubAddress.erase(0, nBreak + 1);
		}
		// The rest is the location, so parse it
		nBreak = sSubAddress.find_first_of(_T("0123456789"));
		if (nBreak == std::tstring::npos)
			continue;
		std::tstring sColumn = sSubAddress.substr(0, nBreak);
		sSubAddress.erase(0, nBreak);
		if (sColumn.empty() || sSubAddress.empty())
			continue;
		if (sColumn[0] == '$')
			sColumn.erase(0, 1);
		if (sColumn.empty())
			continue;
		if (sColumn.at(sColumn.size() - 1) == '$')
			sColumn.erase(sColumn.size() - 1, 1);
		int nColumn = 0;
		while (!sColumn.empty())
		{
			nColumn = nColumn * 26 + (((unsigned char)sColumn[0]) - 'A');
			sColumn.erase(0, 1);
		}
		nColumn++;
		int nRow = _ttoi(sSubAddress.c_str());
		if (nRow < 1)
			continue;

		// Were is it?
		const CellInfo& linkRow = nRow <= arRows.size() ? arRows[nRow - 1] : arRows.back();
		const CellInfo& linkColumn = nColumn <= arColumns.size() ? arColumns[nColumn - 1] : arColumns.back();
		int nDestPage = bDownfirst ? (arRows.back().nPage * (linkColumn.nPage - 1)) + linkRow.nPage : (arColumns.back().nPage * (linkRow.nPage - 1)) + linkColumn.nPage;

		// Calculating the jump location:
		// 1. The jump location is in 72DPI, so we need to change from the print DPI to 72 (multiply by 72, divide by the DPI)
		// 2. The Y jump location IS FROM THE BOTTOM, so we first calculate the correct location by reducing it from the full page height
		// 3. We move a little backwards (-5 X, +5 Y) so we'll get a better view of the location
		double dX = dataHorz.FixLocation(linkColumn.dPos, arPageWidth[linkColumn.nPage - 1]);
		dX = dX * 72.0 / pageHorz.dDPI;
		double dY = dataVert.FixLocation(linkRow.dPos, arPageHeight[linkRow.nPage - 1]);
		dY = (dataVert.lFullSize - dY) * 72 / pageVert.dDPI;
		// OK, add the internal link
		data.AddLink(rect, nPage, nDestPage, dX - 5.0, dY + 5.0, (LPCTSTR)link->GetScreenTip());
	}

	// Write the link data to a file and tell the printer
	if (data.HasData())
	{
		if (!data.SaveProcessData(hPrinter))
		{
			CleanupDialog();
			MessageBox(NULL, ACCESS_PRINTER_ERROR_MSG, PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
			return;
		}
	}
	else
		data.CleanSaved(hPrinter);

	// Test: maybe we should sleep awhile for some reason
	::Sleep(2000);

	// Do print job
	CleanupDialog();
	CComVariant vEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR);
	worksheet->PrintOut(vEmpty, vEmpty, (variant_t)(long)1, VARIANT_FALSE, vEmpty, VARIANT_FALSE, VARIANT_FALSE);
}

/**
	@param pageHorz Horizontal page setup data
	@param pageVert Vertical page setup data
	@param dataHorz Horizontal calculated factors
	@param dataVert Vertical calculated factors
	@param data Print data save/load object to use
	@param hPrinter Handle to the printer
	@param sPrinter Name of the printer
	@param page Excel's page setup object
	@param worksheet Excel's worksheet object
	@return true if data calculated (or loaded from the registry), false if an error occured

	Gets the factors and other data
	In:	Width: DPI, left margin, right margin, center horz
		Height: DPI, top margin, bottom margin, center vert
	Out:Width: Factor, margin/center, full size, force (unsupported) DPI
		Height: Factor, margin/center, full size, force (unsupported) DPI
*/
bool CCCPDFExcelAddinObj::CalculatePrintData(const PrintPageData& pageHorz, const PrintPageData& pageVert, PrintCalcData& dataHorz, PrintCalcData& dataVert, CCPrintData& data, HANDLE hPrinter, const std::tstring& sPrinter, CComPtr<MSExcel::PageSetup>& page, CComQIPtr<MSExcel::_Worksheet>& worksheet)
{
	// Future handling: fit-to-page
	variant_t vZoom = page->GetZoom();
	bool bFitToPage = ((vZoom.vt == VT_BOOL) && (vZoom.boolVal == FALSE));

	if (!bFitToPage)
	{
		if (dataHorz.ReadFromRegistry(pageHorz, _T("Horz")) && dataVert.ReadFromRegistry(pageVert, _T("Vert")))
			// OK, got it
			return true;
	}

	// Prepare a dummy sheet with the same page settings
	double dTopLeftX, dTopLeftY, dBottomRightX, dBottomRightY;
	CComPtr<MSExcel::Workbooks> workbooks = m_spApp->GetWorkbooks();
	CComPtr<MSExcel::_Workbook> workbookTemp = workbooks->Add();
	try
	{
		// Set up the dummy sheet with the same page setup as we had
		workbookTemp->Activate();
		EnsureDialogVisible();
		CComQIPtr<MSExcel::_Worksheet> worksheetTemp = workbookTemp->GetActiveSheet();
		EnsureDialogVisible();
		m_spApp->PutActivePrinter(0, sPrinter.c_str());
		CComPtr<MSExcel::PageSetup> pageTemp = worksheetTemp->GetPageSetup();
		pageTemp->PutCenterHorizontally(pageHorz.bCenter ? VARIANT_TRUE : VARIANT_FALSE);
		pageTemp->PutCenterVertically(pageVert.bCenter ? VARIANT_TRUE : VARIANT_FALSE);
		pageTemp->PutDraft(page->GetDraft());
		pageTemp->PutFooterMargin(page->GetFooterMargin());
		pageTemp->PutHeaderMargin(page->GetHeaderMargin());
		pageTemp->PutLeftMargin(pageHorz.dStartMargin);
		pageTemp->PutRightMargin(pageHorz.dEndMargin);
		pageTemp->PutOrientation(page->GetOrientation());
		pageTemp->PutPaperSize(page->GetPaperSize());
		pageTemp->PutPrintGridlines(VARIANT_FALSE);
		pageTemp->PutPrintHeadings(page->GetPrintHeadings());
		pageTemp->PutPrintTitleColumns(page->GetPrintTitleColumns());
		pageTemp->PutPrintTitleRows(page->GetPrintTitleRows());
		pageTemp->PutTopMargin(pageVert.dStartMargin);
		pageTemp->PutBottomMargin(pageVert.dEndMargin);

		// Future: fit-to-page handling
		if (bFitToPage)
			pageTemp->PutZoom((variant_t)(long)100);
		else
			pageTemp->PutZoom(vZoom);
		try
		{
			// Set the horizontal DPI...
			pageTemp->PutPrintQuality((long)1, pageHorz.dDPI);
		}
		catch (...)
		{
			// Eek! We don't support this DPI, so override it
			dataHorz.bForceDPI = true;
			pageTemp->PutPrintQuality((long)1, 300.0);
		}
		try
		{
			// Set the vertical DPI
			pageTemp->PutPrintQuality((long)2, pageVert.dDPI);
		}
		catch (...)
		{
			// Eek! We don't support this DPI, so override it
			dataVert.bForceDPI = true;
			pageTemp->PutPrintQuality((long)2, 300.0);
		}

		// Set the 'normal' style font, as we need it for correct printing offset
		CComPtr<MSExcel::Styles> stylesTemp = workbookTemp->GetStyles();
		CComPtr<MSExcel::Style> styleTemp = stylesTemp->GetItem(_T("Normal"));
		CComPtr<MSExcel::Font> fontTemp = styleTemp->GetFont();
		fontTemp->PutSize(pageHorz.lFontSize);
		fontTemp->PutName(pageHorz.sFontName.c_str());

		// Write data for the first page so we'll have something to test
		CComPtr<MSExcel::Hyperlinks> links = worksheetTemp->GetHyperlinks();
		CComPtr<MSExcel::Range> range = worksheetTemp->GetRange((variant_t)_T("A1"));
		dTopLeftX = range->GetLeft();
		dTopLeftY = range->GetTop();
		CComVariant vEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR);
		CComPtr<MSExcel::Range> columns = worksheetTemp->GetColumns();
		CComPtr<MSExcel::Range> rows = worksheetTemp->GetRows();
		CComQIPtr<MSExcel::Range> temp;

		// 'a', linked, at the top-left cell
		links->Add(range, _T("http://topleft.com"), vEmpty, vEmpty, _T("a"));
		range->PutVerticalAlignment((long)MSExcel::xlTop);
		range->PutHorizontalAlignment((long)MSExcel::xlLeft);
		data.AddLink(_T("http://topleft.com"), _T("a"), 1);

		// Find the bottom-right cell of the first page
		long lRow = 2, lCol = 2;
		do
		{
			// Write in the next column
			range = worksheetTemp->GetRange((variant_t)CellName(1, lCol).c_str());
			range->PutFormula(_T("c"));
			temp = columns->GetItem(lCol);
			if (temp->GetPageBreak() != MSExcel::xlPageBreakNone)
				break;
			// Not yet a page break
			range->ClearContents();
			lCol++;
		} while (true);
		range->ClearContents();
		lCol--;

		do
		{
			// Write in the next row
			range = worksheetTemp->GetRange((variant_t)CellName(lRow, 1).c_str());
			range->PutFormula(_T("c"));
			temp = rows->GetItem(lRow);
			if (temp->GetPageBreak() != MSExcel::xlPageBreakNone)
				break;
			// Not yet a page break
			range->ClearContents();
			lRow++;
		} while (true);
		range->ClearContents();
		lRow--;

		// OK, write data in the bottom-right cell
		range = worksheetTemp->GetRange((variant_t)CellName(lRow, lCol).c_str());
		dBottomRightX = (double)range->GetLeft() + (double)range->GetWidth();
		dBottomRightY = (double)range->GetTop() + (double)range->GetHeight();
		links->Add(range, _T("http://bottomright.com"), vEmpty, vEmpty, _T("b"));
		// Make sure it's bottom/right aligned!
		range->PutHorizontalAlignment((long)MSExcel::xlRight);
		range->PutVerticalAlignment((long)MSExcel::xlBottom);
		data.AddLink(_T("http://bottomright.com"), _T("b"), 1);
		// This is a test only:
		data.SetTestPage();

		// If we use fit to page, we need to do it a little differently; future only!
		if (bFitToPage)
		{
			// Find the location of the end of the original page's range
			range = worksheet->GetUsedRange();
			long lColumns = range->GetColumns()->GetCount() + range->GetColumn(), lRows = range->GetRows()->GetCount() + range->GetRow();
			range = worksheet->GetRange((variant_t)CellName(lRows, lColumns).c_str());
			double dEndX = range->GetLeft().dblVal + range->GetWidth().dblVal, dEndY = range->GetTop().dblVal + range->GetWidth().dblVal;
			// Find where those are in the current sheet
			CComPtr<MSExcel::Range> columns = worksheetTemp->GetColumns();
			CComPtr<MSExcel::Range> rows = worksheetTemp->GetRows();
			lColumns = 1;
			CComQIPtr<MSExcel::Range> temp;
			while (true)
			{
				temp = columns->GetItem(lColumns);
				if (temp->GetLeft().dblVal + temp->GetWidth().dblVal >= dEndX)
					break;
				lColumns++;
			}
			lRows = 1;
			while (true)
			{
				temp = rows->GetItem(lRows);
				if (temp->GetTop().dblVal + temp->GetHeight().dblVal >= dEndY)
					break;
				lRows++;
			}
			range = worksheetTemp->GetRange((variant_t)CellName(lRows, lColumns).c_str());
			range->PutFormula(_T("xyz"));
			pageTemp->PutFitToPagesTall(page->GetFitToPagesTall());
			pageTemp->PutFitToPagesWide(page->GetFitToPagesWide());
		}

		// Print this page as a test page
		if (!data.SaveProcessData(hPrinter))
		{
			CleanupDialog();
			MessageBox(NULL, ACCESS_PRINTER_ERROR_MSG, PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
			return false;
		}
		try
		{
			worksheetTemp->PrintOut((variant_t)(long)1, (variant_t)(long)1, (variant_t)(long)1, VARIANT_FALSE, vEmpty, VARIANT_FALSE, VARIANT_FALSE);
		}
		catch (...)
		{
		}
	}
	catch (...)
	{
		// Crash for some reason, try to recover
		try
		{
			workbookTemp->Close(VARIANT_FALSE);
		}
		catch(...)
		{
		}
		throw;
	}
	// Finished with the temporary 
	workbookTemp->Close(VARIANT_FALSE);

	// Read back the data
	if (!data.ReloadProcessData(hPrinter) || (data.GetPageCount() != 1) || (data.GetPageData(1).size() != 2))
	{
		// Can't get the data back, fail
		CleanupDialog();
		MessageBox(NULL, WORKING_PRINTER_ERROR_MSG, PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
		return false;
	}

	// Calculate the multiplication numbers
	int nTopLeftX, nTopLeftY, nBottomRightX, nBottomRightY;
	if (data.GetPageCount() != 1)
	{
		// Can't get the data back, fail
		CleanupDialog();
		MessageBox(NULL, WORKING_PRINTER_ERROR_MSG, PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
		return false;
	}

	// Only one page data back
	const CCPrintData::PageData& pageData = data.GetPageData(1);
	for (CCPrintData::PageData::const_iterator iData = pageData.begin(); iData != pageData.end(); iData++)
	{
		const CCPrintData::LinkData& link = (*iData);
		// Is this the top-left link?
		if (link.sURL == _T("http://topleft.com"))
		{
			// Yes, get the data back
			nTopLeftX = link.rectLocation.left;
			nTopLeftY = link.rectLocation.top;
		}
		// Bottom-right link?
		else if (link.sURL == _T("http://bottomright.com"))
		{
			// Yes, get the data back
			nBottomRightX = link.rectLocation.right;
			nBottomRightY = link.rectLocation.bottom;
		}
		else
		{
			// Can't get the data back, fail
			CleanupDialog();
			MessageBox(NULL, WORKING_PRINTER_ERROR_MSG, PRINTER_ERROR_TITLE, MB_OK|MB_ICONERROR);
			return false;
		}
	}

	// Calculate print data
	dataHorz.FromData(pageHorz, nBottomRightX, nTopLeftX, dBottomRightX, dTopLeftX, pageData.szPage.cx);
	dataVert.FromData(pageVert, nBottomRightY, nTopLeftY, dBottomRightY, dTopLeftY, pageData.szPage.cy);

	if (!bFitToPage)
	{
		dataHorz.WriteToRegistry(pageHorz, _T("Horz"));
		dataVert.WriteToRegistry(pageVert, _T("Vert"));
	}

	// Create read link data
	data.CleanThis();
	return true;
}
