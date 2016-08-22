/**
	@file
	@brief Main functions file for CCPDFConverter/XL2PDFConverter application
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

#include "iapi.h"
#include <shellapi.h>
#include <errno.h>
#include <stdio.h>
#include "Helpers.h"
#include <io.h>

#ifdef CC_PDF_CONVERTER
#define PRODUCT_NAME	"CC PDF Converter"
#elif EXCEL_TO_PDF
#define PRODUCT_NAME	"Excel to PDF Converter"
#else
#error "One of the printer types must be defined"
#endif

#ifdef _DEBUG
/// Debugging file pointer
FILE* pSave = NULL;
#endif
/// Input file pointer
FILE* fileInput;
/// Initial input buffer
char cBuffer[MAX_PATH * 2 + 1];
/// Length of data in initial buffer
int nBuffer = 0;
/// Current location in the initial buffer
int nInBuffer = 0;
/// Size of error string buffer
#define MAX_ERR		1023
/// Error string buffer
char cErr[MAX_ERR + 1];

#ifdef _DEBUG
/**
	@brief This function outputs an error via OutputDebugStringn
	@param pBefore Text to add before the error
	@param buf Error description
	@param len Size of error description
*/
static void WriteOutput(const char* pBefore, const char* buf, int len)
{
	// Calculate and create a large enough buffer for the error descriptionn
	int n = len + strlen(pBefore);
	char* pStr = new char[n + 1];
	// Fill it
	sprintf_s(pStr, n + 1, "%s%.*s", pBefore, len, buf);
	// Send it
	::OutputDebugString(pStr);
	// Cleanup
	delete [] pStr;
}
#endif

//////////////////////////////////////////////////////////////////////////
/**
	@param pPath Path to test
	@return true if a folder (not file) exists in the path
*/
bool ExistsAsFolder(LPCTSTR pPath)
{
	WIN32_FILE_ATTRIBUTE_DATA data;
	if (!::GetFileAttributesEx(pPath, GetFileExInfoStandard, &data))
		return false;
	if (data.dwFileAttributes & (FILE_ATTRIBUTE_OFFLINE|FILE_ATTRIBUTE_SPARSE_FILE|FILE_ATTRIBUTE_TEMPORARY))
		return false;

	return ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0);
}
//////////////////////////////////////////////////////////////////////////

#define TEMP_FILENAME "ccprint_"
#define TEMP_EXTENSION "pdf"

void CleanTempFiles ()
{
	char sTempFolder[MAX_PATH];
	GetTempPath(MAX_PATH, sTempFolder);

	char sFileToFind[MAX_PATH];
	sprintf_s (sFileToFind, "%s%s*.%s", sTempFolder, TEMP_FILENAME, TEMP_EXTENSION);

	struct _finddata64i32_t finddata;
	intptr_t hFind = _findfirst (sFileToFind, &finddata);
	int ret = (int)hFind;
	char sFullFilename[MAX_PATH];

	while (ret != -1) {
		sprintf_s (sFullFilename, "%s%s", sTempFolder, finddata.name);
		DeleteFile (sFullFilename);						// Deliberately ignore the return value
		ret = _findnext (hFind, &finddata);				// Find the next file
	}
	
	_findclose (hFind);
}


//////////////////////////////////////////////////////////////////////////

/**
	@brief Callback function used by GhostScript to retrieve more data from the input buffer; stops at newlines
	@param instance Pointer to the GhostScript instance (not used)
	@param buf Buffer to fill with data
	@param len Length of requested data
	@return Size of retrieved data (in bytes), 0 when there's no more data
*/
static int GSDLLCALL my_in(void *instance, char *buf, int len)
{
	// Initialize variables
    int ch;
    int count = 0;
	char* pStart = buf;
	// Read until we reached the wanted size...
    while (count < len) 
	{
		// Is there still data in the initial buffer?
		if (nBuffer > nInBuffer)
			// Yes, read from there
			ch = cBuffer[nInBuffer++];
		else
			// No, get more data
			ch = fgetc(fileInput);
		if (ch == EOF)
			// That's it
			return 0;
		// Put the character in the buffer and increate the countn
		*buf++ = ch;
		count++;
		if (ch == '\n')
			// Stop on newlines
			break;
    }
#ifdef _DEBUG
	// Leave a trace of the data (debug mode)
	WriteOutput("", pStart, count);
	if (pSave != NULL)
	{
		// Also save the data into the save file (debug mode)
		fwrite(pStart, 1, count, pSave);
	}
#endif
	// That's it
    return count;
}

/**
	@brief Callback function used by GhostScript to output notes and warnings
	@param instance Pointer to the GhostScript instance (not used)
	@param str String to output
	@param len Length of output
	@return Count of characters written
*/
static int GSDLLCALL my_out(void *instance, const char *str, int len)
{
#ifdef _DEBUG
	// Write to stdout (debug mode)
    fwrite(str, 1, len, stdout);
    fflush(stdout);
	// Trace also (debug mode)
	WriteOutput("OUT: ", str, len);
#endif

	// That's it
    return len;
}

/**
	@brief Callback function used by GhostScript to output errors
	@param instance Pointer to the GhostScript instance (not used)
	@param str Error string
	@param len Length of string
	@return Count of characters written
*/
static int GSDLLCALL my_err(void *instance, const char *str, int len)
{
#ifdef _DEBUG
	// Write to stderr (debug mode)
    fwrite(str, 1, len, stderr);
    fflush(stderr);
	// Trace too (debug mode)
	WriteOutput("ERR: ", str, len);
#endif
	// Keep the error in cErr for later handling
	int nAdd = min(len, (int)(MAX_ERR - strlen(cErr)));
	strncat_s(cErr, str, MAX_ERR);
	// OK
    return len;
}

/**
	Reads all the data from the input (so no error will be raised if application
	ends without sending the data to ghostscript)
*/
void CleanInput()
{
	char cBuffer[1024];
	while (fread(cBuffer, 1, 1024, fileInput) > 0)
		;
}

/**
	@brief This function will center the window on the screen
	@param hWnd The window to center
*/
void CenterWindow(HWND hWnd)
{
	// get coordinates of the window relative to the screen
	RECT rcWnd;
	::GetWindowRect(hWnd, &rcWnd);
	RECT rcCenter, rcArea;
	// center within screen coordinates
	::SystemParametersInfo(SPI_GETWORKAREA, NULL, &rcArea, NULL);
	rcCenter = rcArea;

	int WndWidth = rcWnd.right - rcWnd.left;
	int WndHeight = rcWnd.bottom - rcWnd.top;

	// find dialog's upper left based on rcCenter
	int xLeft = (rcCenter.left + rcCenter.right) / 2 - WndWidth / 2;
	int yTop = (rcCenter.top + rcCenter.bottom) / 2 - WndHeight / 2;

	// if the dialog is outside the screen, move it inside
	if (xLeft < rcArea.left)
		xLeft = rcArea.left;
	else if(xLeft + WndWidth > rcArea.right)
		xLeft = rcArea.right - WndWidth;

	if(yTop < rcArea.top)
		yTop = rcArea.top;
	else if(yTop + WndHeight > rcArea.bottom)
		yTop = rcArea.bottom - WndHeight;

	// map screen coordinates to child coordinates
	::SetWindowPos(hWnd, NULL, xLeft, yTop, -1, -1, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

/**
	@param hDlg Handle of the dialog
	@param uMsg ID of the message
	@param wParam First message paramenter
	@param lParam Second message paramenter
	@return TRUE if the message was handled, FALSE otherwise
*/
UINT_PTR CALLBACK SaveDlgCallback(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	// Which message?
	switch (uMsg)
	{
		case WM_NOTIFY:
			// Notification
			{
				LPNMHDR pNotify = (LPNMHDR)lParam;
				if (pNotify->code == CDN_INITDONE)
				{
					// Initial display: center and bring to top
					HWND hParent = ::GetParent(hDlg);
					::BringWindowToTop(hParent);
					SetForegroundWindow(hParent);
					CenterWindow(hParent);
				}
			}
			break;
	}
	return FALSE;
}

/// Command line options used by GhostScript
const char* ARGS[] =
{
	"PS2PDF",
	"-dNOPAUSE",
	"-dBATCH",
    "-dSAFER",
    "-sDEVICE=pdfwrite",
	"-sOutputFile=c:\\test.pdf",
	"-I.\\",
    "-c",
    ".setpdfwrite",
	"-"
};

/**
	@brief Main function
	@param hInstance Handle to the current instance
	@param hPrevInstance Handle to the previous running instance (not used)
	@param lpCmdLine Command line (not used)
	@param nCmdShow Initial window visibility and location flag (not used)
	@return 0 if all went well, other values upon errors
*/
int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	// Initialize stuff
	char cPath[MAX_PATH + 1];
	char cFile[MAX_PATH + 128];
	char cInclude[3 * MAX_PATH + 7];
	cErr[0] = '\0';

#ifdef _DEBUG
	// Save a record of the original PostScript data (debug mode)
	errno_t file_err = fopen_s (&pSave, "c:\\test.ps", "w+b");
#endif

	// Delete whichever temp files might exist
	CleanTempFiles();

	// Add the include directories to the command line flags we'll use with GhostScript:
	if (::GetModuleFileName(NULL, cPath, MAX_PATH))
	{
		// Should be next to the application
		char* pPos = strrchr(cPath, '\\');
		if (pPos != NULL)
			*(pPos) = '\0';
		else
			cPath[0] = '\0';
		// OK, add the fonts and lib folders:
		sprintf_s (cInclude, sizeof(cInclude), "-I%s\\urwfonts;%s\\lib", cPath, cPath);
		ARGS[6] = cInclude;
	}

#ifdef _DEBUG_CMD
	// Sample file debug mode: open a pre-existing file
	fileInput = fopen("c:\\test1.ps", "rb");
#else
	// Get the data from stdin (that's where the redmon port monitor sends it)
	fileInput = stdin;
#endif

	// Check if we have a filename to write to:
	cPath[0] = '\0';
	bool bAutoOpen = false;
	bool bMakeTemp = false;
	// Read the start of the file; if we have a filename and/or the auto-open flag, they must be there:
	nBuffer = fread(cBuffer, 1, MAX_PATH * 2, fileInput);
	cBuffer[nBuffer] = EOF;

	// Do we have a %%File: starting the buffer?
	if ((nBuffer > 8) && (strncmp(cBuffer, "%%File: ", 8) == 0))
	{
		// Yes, so read the filename
		char ch;
		int nCount = 0;
		nInBuffer += 8;
		do
		{
			ch = cBuffer[nInBuffer++];
			if (ch == EOF)
				break;
			if (ch == '\n')
				break;
			cPath[nCount++] = ch;
		} while (true);

		if (ch == EOF)
		{
			// If we didn't find a newline, something ain't right
			return 0;
		}

		// OK, found the page, so set it as a command line variable now
		cPath[nCount] = '\0';

		// Sometimes we don't want any output:
		if (strcmp(cPath, ":dropfile:") == 0)
		{
			// Nothing doing
			CleanInput();
			return 0;
		}

		sprintf_s(cFile, sizeof(cFile), "-sOutputFile=%s", cPath);
		ARGS[5] = cFile;
#ifdef _DEBUG
		// Trace it (debug mode)
		WriteOutput("FILENAME: ", cPath, nCount);
#endif
	}
	// Do we have an auto-file-open flag?
	if ((nBuffer - nInBuffer > 14) && ((!strncmp(cBuffer + nInBuffer, "%%FileAutoOpen", 14)) || (!strncmp(cBuffer + nInBuffer, "%%CreateAsTemp", 14))))
	{
		// Yes, found it, so jump over it until the newline
		if (!strncmp(cBuffer + nInBuffer, "%%CreateAsTemp", 14))
			bMakeTemp = true;

		nInBuffer += 14;
		bAutoOpen = true;
		while ((cBuffer[nInBuffer] != EOF) && (cBuffer[nInBuffer] != '\n'))
			nBuffer++;
		if (cBuffer[nInBuffer] == EOF)
		{
			// Nothing else, leave
			return 0;
		}
	}
	
	// Did we find a filename?
	if (cPath[0] == '\0')
	{
		// Do we make it a temp file?
		if (bMakeTemp) {
			char sTempFolder[MAX_PATH];
			GetTempPath(MAX_PATH, sTempFolder);
			sprintf_s (cPath, MAX_PATH, "%s%s%u.%s", sTempFolder, TEMP_FILENAME, GetTickCount(), TEMP_EXTENSION);
			
			HANDLE test = CreateFile (cPath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_DELETE_ON_CLOSE, NULL);
			if (test == INVALID_HANDLE_VALUE) {
				// If we can't write this file, for some reason:
				bMakeTemp = false;
			}
			else {
				CloseHandle (test);	
				sprintf_s (cFile, sizeof(cFile), "-sOutputFile=%s", cPath);
				ARGS[5] = cFile;
			}			
		}
		
		// It's possible that if something fails in the process of making a temp file, the bMakeTemp flag
		// will be disabled in the above block and then we want to run the following block as usual.
		if (!bMakeTemp) {
			// Ask the user for a file name:
			OPENFILENAME info;
			memset(&info, 0, sizeof(info));
			info.lStructSize = sizeof(info);
			info.hInstance = hInstance;
			info.lpstrFilter = "PDF Files (*.pdf)\0*.pdf\0All Files (*.*)\0*.*\0\0";
			info.lpstrFile = cPath;
			info.nMaxFile = MAX_PATH + 1;
			info.lpstrTitle = "Select a filename to write into";
			info.Flags = OFN_ENABLESIZING|OFN_EXPLORER|OFN_NOREADONLYRETURN|OFN_OVERWRITEPROMPT|OFN_PATHMUSTEXIST|OFN_ENABLEHOOK;
			info.lpstrDefExt = "pdf";
			// We want a hook function to center the window and bring it on top
			info.lpfnHook = SaveDlgCallback;

			if (GetSaveFileName(&info))
			{
				// OK, get a filename, write it up
				sprintf_s (cFile, sizeof(cFile), "-sOutputFile=%s", cPath);
				ARGS[5] = cFile;
#ifdef _DEBUG
				// Also trace it (debug mode)
				WriteOutput("FILENAME (USER): ", cPath, strlen(cPath));
#endif
			}
			else
			{
				// Continue reading until to end so we won't have a problem
				CleanInput();
				return 0;
			}
		}
	}

	// First try to initialize a new GhostScript instance
	void* pGS;
	if (gsapi_new_instance(&pGS, NULL) < 0)
	{
		// Error 
		return -1;
	}

	// Set up the callbacks
	if (gsapi_set_stdio(pGS, my_in, my_out, my_err) < 0)
	{
		// Failed...
		gsapi_delete_instance(pGS);
		return -2;
	}

	// Now run the GhostScript engine to transform PostScript into PDF
	int nRet = gsapi_init_with_args(pGS, sizeof(ARGS)/sizeof(char*), (char**)ARGS);

	gsapi_exit(pGS);
	gsapi_delete_instance(pGS);
		
#ifdef _DEBUG
	// Close the PostScript copy file (debug mode)
	fclose(pSave);
#endif
#ifdef _DEBUG_CMD
	// Close the sample file (sample file debug mode)
	fclose(fileInput);
#endif

	// Did we get an error?
	if (strlen(cErr) > 0)
	{
		// Yes, show it
		MessageBox(NULL, cErr, PRODUCT_NAME, MB_ICONERROR|MB_OK);
		return 0;
	}

	// Should we open the file (also make sure there's a handler for PDFs)
	if (bAutoOpen && CanOpenPDFFiles()) {
		// Yes, so open it
		ShellExecute(NULL, NULL, cPath, NULL, NULL, SW_NORMAL);
	}


	return 0;
}
