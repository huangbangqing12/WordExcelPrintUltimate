#include "stdafx.h"
#include "PrintWord.h"

void PrintWord::QuitMsDispatch()
{
	_wordApp.Quit(_vOpt, _vOpt, _vOpt);
	_wordApp.ReleaseDispatch();
}

PrintErrorStatus PrintWord::PrintOutWord(CString strSourceFile, CString strOutputFilename, CString strPrinterName, WdPaperSize dmPaperType, WdOrientation dmOrient, long lCopies)
{
	CString strDefaultPrinter = strPrinterName;
	
	if (_bIsMsWord)
	{
		return PrintOutMSWord(strSourceFile, strOutputFilename, strDefaultPrinter, dmPaperType, dmOrient, lCopies);
	}
	else
	{
		return PrintOutWPSWord(strSourceFile, strOutputFilename, strDefaultPrinter, dmPaperType, dmOrient, lCopies);
	}
}

void PrintWord::EndPrintOutWord()
{
	QuitMsDispatch();
}

PrintWord::PrintWord()
{
	_bIsMsWord = true;
	_vOpt = COleVariant((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
}

PrintWord::~PrintWord()
{
	EndPrintOutWord();
}

PrintErrorStatus PrintWord::BeginPrintOutWord()
{
	CCommonMethod::AvoidOtherProgrameRunning();

	if (!_wordApp.CreateDispatch(_T("Word.Application"), NULL))
	{
		_bIsMsWord = false;
		if (!_wordApp.CreateDispatch(_T("Wps.Application"), NULL))
		{
			_bIsMsWord = true;
			if (!_wordApp.CreateDispatch(_T("Kwps.Application"), NULL))
			{
				QuitMsDispatch();
				return ePrintNoneWpsOfficeInstall;
			}
		}
	}
	return ePrintOk;
}

PrintErrorStatus PrintWord::PrintOutMSWord(CString strSourceFile, CString strOutputFilename, CString strDefaultPrinter, WdPaperSize dmPaperType, WdOrientation dmOrient, long lCopies)
{
	CMsWordDocuments     docs;
	CMsWordDocument0     doc;

	LPDISPATCH pDoc = NULL;
	try
	{
		_wordApp.put_Visible(FALSE);
		_wordApp.put_ActivePrinter(strDefaultPrinter);
		docs = _wordApp.get_Documents();

		// 打开文档
		pDoc = docs.Open(COleVariant(strSourceFile)
			, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt);

		if (NULL == pDoc)
		{
			return ePrintDocumentBeenOpened;
		}

		doc = pDoc;
		doc.Select();

		// 设置打印横纵向及纸张大小
		CMsWordPageSetup pageSetup = doc.get_PageSetup();
		pageSetup.put_PaperSize(long(9));
		pageSetup.put_Orientation(dmOrient);

		// 执行打印
		if (strOutputFilename == L"")
		{
			doc.PrintOut(_vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, COleVariant(lCopies)
				, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt);
		}
		else
		{
			doc.PrintOut(_vOpt, _vOpt, _vOpt, COleVariant(strOutputFilename), _vOpt, _vOpt, _vOpt, COleVariant(lCopies)
				, _vOpt, _vOpt, COleVariant((short)true), _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, _vOpt);
		}

		// 释放资源
		pageSetup.ReleaseDispatch();
		doc.Close(_vOpt, _vOpt, _vOpt);
		doc.ReleaseDispatch();
		docs.ReleaseDispatch();
		//CCommonMethod::ResetDeafultPrinter(strDefaultPrinter);

		return ePrintOk;
	}
	catch (...)
	{
		// 释放资源
		docs.ReleaseDispatch();
		//CCommonMethod::ResetDeafultPrinter(strDefaultPrinter);
		return ePrintOfficeThrowException;
	}
}

PrintErrorStatus PrintWord::PrintOutWPSWord(CString strSourceFile, CString strOutputFilename, CString strDefaultPrinter, WdPaperSize dmPaperType, WdOrientation dmOrient, long lCopies)
{
	CWpsWordDocuments     docs;
	CWpsWordDocument0     doc;

	LPDISPATCH pDoc = NULL;
	try
	{
		_wordApp.put_Visible(FALSE);
		_wordApp.put_ActivePrinter(strDefaultPrinter);
		docs = _wordApp.get_Documents();
		pDoc = docs.Open(strSourceFile
			, FALSE, FALSE, FALSE, L"", L"", FALSE, L"", L"", 0, 0, TRUE, FALSE, 0, FALSE);

		if (NULL == pDoc)
		{
			return ePrintDocumentBeenOpened;
		}

		doc = pDoc;
		// 设置打印横纵向及纸张大小
		CWpsWordPageSetup pageSetup = doc.get_PageSetup();
		pageSetup.put_PaperSize(dmPaperType);
		pageSetup.put_Orientation(dmOrient);

		// 执行打印
		if (strOutputFilename == L"")
		{
			static BYTE parms[] = VTS_BOOL VTS_BOOL VTS_I4 VTS_BSTR VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BSTR VTS_I4 VTS_BOOL VTS_BOOL /*VTS_PVARIANT VTS_BOOL VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BOOL VTS_I4 VTS_BOOL VTS_I4*/;
			doc.InvokeHelper(0x1be, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FALSE, FALSE, 0, L"", 0, 0, 0, lCopies, L"", 0, FALSE, FALSE/*, _vOpt, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight, FlipPrint, PaperTray, CutterLine, PaperOrder*/);
		}
		else
		{
			static BYTE parms[] = VTS_BOOL VTS_BOOL VTS_I4 VTS_BSTR VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BSTR VTS_I4 VTS_BOOL VTS_BOOL /*VTS_PVARIANT VTS_BOOL VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_BOOL VTS_I4 VTS_BOOL VTS_I4*/;
			doc.InvokeHelper(0x1be, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FALSE, FALSE, 0, strOutputFilename, 0, 0, 0, lCopies, L"", 0, TRUE, FALSE/*, _vOpt, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight, FlipPrint, PaperTray, CutterLine, PaperOrder*/);
		}

		// 释放资源
		pageSetup.ReleaseDispatch();
		doc.Save();
		doc.Close(_vOpt, _vOpt, _vOpt);
		doc.ReleaseDispatch();
		docs.ReleaseDispatch();
		//CCommonMethod::ResetDeafultPrinter(strDefaultPrinter);

		return ePrintOk;
	}
	catch (...)
	{
		// 释放资源
		//CCommonMethod::ResetDeafultPrinter(strDefaultPrinter);
		return ePrintWPSVersionBelow2010;
	}
}
