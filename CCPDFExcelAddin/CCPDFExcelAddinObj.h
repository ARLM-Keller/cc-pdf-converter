/**
	@file
	@brief Declaration of the CCCPDFExcelAddinObj
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

#ifndef __CCPDFEXCELADDINOBJ_H_
#define __CCPDFEXCELADDINOBJ_H_

#include "resource.h"       // main symbols
// This is taken from the COMMON FILES folder
#import "DESIGNER\MSADDNDR.TLB" raw_interfaces_only, raw_native_types, no_namespace, named_guids 
// This is taken from the appropriate Office folder
#import "EXCEL9.olb" rename("DialogBox","MyDialogBox"),rename("RGB","MyRGB"),named_guids, rename_namespace("MSExcel")
using namespace MSExcel;

#include "CCTChar.h"

/////////////////////////////////////////////////////////////////////////////
// CCCPDFExcelAddinObj
/**
    @brief Class for connecting to Excel and handling the Excel to PDF command
*/
class ATL_NO_VTABLE CCCPDFExcelAddinObj : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public ISupportErrorInfo,
#ifdef CC_PDF_CONVERTER
	public CComCoClass<CCCPDFExcelAddinObj, &CLSID_CCPDFExcelAddinObj>,
	public IDispatchImpl<ICCPDFExcelAddinObj, &IID_ICCPDFExcelAddinObj, &LIBID_CCPDFEXCELADDINLib>,
#elif EXCEL_TO_PDF
	public CComCoClass<CCCPDFExcelAddinObj, &CLSID_XL2PDFExcelAddinObj>,
	public IDispatchImpl<IXL2PDFExcelAddinObj, &IID_IXL2PDFExcelAddinObj, &LIBID_XL2PDFEXCELADDINLib>,
#else
#error "Please define one of the printer types"
#endif
	public IDispatchImpl<_IDTExtensibility2, &IID__IDTExtensibility2, &LIBID_AddInDesignerObjects>,
	public IDispEventSimpleImpl<1, CCCPDFExcelAddinObj, &__uuidof(MSExcel::AppEvents)>,
	public IDispEventSimpleImpl<2, CCCPDFExcelAddinObj, &__uuidof(Office2000::_CommandBarButtonEvents)>
{
public:
	/**
		@brief Default constructor
	*/
	CCCPDFExcelAddinObj() : m_hDialog(NULL)
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_CCPDFEXCELADDINOBJ)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CCCPDFExcelAddinObj)
#ifdef CC_PDF_CONVERTER
	COM_INTERFACE_ENTRY(ICCPDFExcelAddinObj)
	COM_INTERFACE_ENTRY2(IDispatch, ICCPDFExcelAddinObj)
#elif EXCEL_TO_PDF
	COM_INTERFACE_ENTRY(IXL2PDFExcelAddinObj)
	COM_INTERFACE_ENTRY2(IDispatch, IXL2PDFExcelAddinObj)
#else
#error "Please define one of the printer types"
#endif
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(_IDTExtensibility2)
END_COM_MAP()

BEGIN_SINK_MAP(CCCPDFExcelAddinObj)
	SINK_ENTRY_INFO(1, __uuidof(MSExcel::AppEvents), 0x620, OnWorkbookActivate, &DocumentOpenInfo)
	SINK_ENTRY_INFO(1, __uuidof(MSExcel::AppEvents), 0x622, OnWorkbookBeforeClose, &DocumentBeforeCloseInfo)
	SINK_ENTRY_INFO(1, __uuidof(MSExcel::AppEvents), 0x61d, OnNewWorkbook, &DocumentNew)
	SINK_ENTRY_INFO(2, __uuidof(Office2000::_CommandBarButtonEvents), /*dispid*/ 0x01, OnButtonClick, &OnClickButtonInfo)
END_SINK_MAP()


// ISupportsErrorInfo
	/// Checks if the requested interface is supported by the object
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// ICCPDFExcelAddinObj
public:
// _IDTExtensibility2
	/// Called when connected to Excel
	STDMETHOD(OnConnection)(IDispatch * Application, ext_ConnectMode ConnectMode, IDispatch * AddInInst, SAFEARRAY * * custom);
	/// Called when disconnected from Excel
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom);
	/**
		@brief Called when Excel's Addins list is updated
		@param custom Additional data (not used)
	*/
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom)
	{
		return E_NOTIMPL;
	}
	/**
		@brief Called when Excel's startup process is complete
		@param custom Additional data (not used)
	*/
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom)
	{
		return E_NOTIMPL;
	}
	/**
		@brief Called when Excel's shutdown process starts
		@param custom Additional data (not used)
	*/
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom)
	{
		return E_NOTIMPL;
	}

	// IDispEventSimpleImpl
	/// Called when a new workbook is opened in Excel
	void __stdcall OnNewWorkbook(IDispatch * workbook);
	/// Called when a workbook becomes active in Excel
	void __stdcall OnWorkbookActivate(IDispatch * workbook);
	/// Called before a workbook in Excel is closed
	void __stdcall OnWorkbookBeforeClose(IDispatch * workbook, VARIANT_BOOL * CancelDefault);
	/// Called when the user uses the Excel to PDF command (via menu item or toolbar button)
	void __stdcall OnButtonClick(IDispatch * button, VARIANT_BOOL * CancelDefault);

protected:
	/// Reference to the command added to the toolbar
	CComPtr<Office2000::_CommandBarButton> m_buttonToolbar;
	/// Reference to the command added to the menu
	CComPtr<Office2000::_CommandBarButton> m_buttonMenu;
	/// Reference to the Excel application
	CComPtr<MSExcel::_Application> m_spApp;
	/// Handle of the 'processing' dialog which is displayed when calculating link data
	HWND	m_hDialog;

	/// Enables the toolbar buttons and menu items
	bool	EnableButtons(bool bEnable);
	/// Creates a PDF from the current worksheet
	void	DoPrint(HANDLE hPrinter, const std::tstring& sPrinter);
	/// Updates the current printer (and sets the page breaks correctly)
	void	SetPrinterAndBreaks(const std::tstring& sPrinter);
	/// Calculates the location factors for specific print parameters for this document
	bool	CalculatePrintData(const struct PrintPageData& pageHorz, const PrintPageData& pageVert, struct PrintCalcData& dataHorz, PrintCalcData& dataVert, class CCPrintData& data, HANDLE hPrinter, const std::tstring& sPrinter, CComPtr<MSExcel::PageSetup>& page, CComQIPtr<MSExcel::_Worksheet>& worksheet);
	/**
		@brief Cleans up the 'processing' dialog, if it is currently displayed
	*/
	void	CleanupDialog() {if (m_hDialog != NULL) DestroyWindow(m_hDialog); m_hDialog = NULL;};
	/**
		@brief Makes sure that the 'processing' dialog is visible
	*/
	void	EnsureDialogVisible() {if (m_hDialog == NULL) return; ::BringWindowToTop(m_hDialog); UpdateWindow(m_hDialog);};

	/// Callback information for Excel button click
	static _ATL_FUNC_INFO OnClickButtonInfo;
	/// Callback information for Excel open document
	static _ATL_FUNC_INFO DocumentOpenInfo;
	/// Callback information for Excel new document
	static _ATL_FUNC_INFO DocumentNew;
	/// Callback information for Excel closing the current workbook
	static _ATL_FUNC_INFO DocumentBeforeCloseInfo;
};

#endif //__CCPDFEXCELADDINOBJ_H_
