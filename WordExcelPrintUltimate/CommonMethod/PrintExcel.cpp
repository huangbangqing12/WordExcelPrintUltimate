#include "stdafx.h"
#include "PrintExcel.h"


PrintExcel::PrintExcel()
{
	_bIsMsExcel = true;
	_vOpt = COleVariant((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
}

PrintExcel::~PrintExcel()
{
	EndPrintOutExcel();
}

PrintErrorStatus PrintExcel::BeginPrintOutExcel()
{
	CCommonMethod::AvoidOtherProgrameRunning();

	if (!_excelMsApp.CreateDispatch(_T("Excel.Application"), NULL))
	{
		_bIsMsExcel = false;
		if (!_excelMsApp.CreateDispatch(_T("et.Application"), NULL))
		{
			_bIsMsExcel = true;
			if (!_excelMsApp.CreateDispatch(_T("Ket.Application"), NULL))
			{
				QuitMsDispatch();
				return ePrintNoneWpsOfficeInstall;
			}
		}
	}

	if (!_bIsMsExcel)
	{
		_excelWpsApp.CreateDispatch(_T("et.Application"), NULL);
	}
	return ePrintOk;
}

void PrintExcel::QuitMsDispatch()
{
	_excelMsApp.Quit();
	_excelMsApp.ReleaseDispatch();
}

void PrintExcel::QuitWpsDispatch()
{
	_excelWpsApp.Quit();
	_excelWpsApp.ReleaseDispatch();
}

PrintErrorStatus PrintExcel::PrintOutExcel(CString strSourceFile, CString strOutputFilename, CString strPrinterName, XlPaperSize dmPaperType, XlPageOrientation dmOrient, long lCopies)
{
	CString strDefaultPrinter = CCommonMethod::SetPrinterPara(strPrinterName);

	if (_bIsMsExcel)
	{
		return PrintOutMsExcel(strSourceFile, strOutputFilename, strDefaultPrinter, dmPaperType, dmOrient, lCopies);
	}
	else
	{
		return PrintOutWpsExcel(strSourceFile, strOutputFilename, strDefaultPrinter, dmPaperType, dmOrient, lCopies);
	}
}

void PrintExcel::EndPrintOutExcel()
{
	if (_bIsMsExcel)
	{
		QuitMsDispatch();
	}
	else
	{
		QuitWpsDispatch();
	}
}

PrintErrorStatus PrintExcel::PrintOutMsExcel(CString strSourceFile, CString strOutputFilename, CString strDefaultPrinter, XlPaperSize dmPaperType, XlPageOrientation dmOrient, long lCopies)
{
	CMsExcelWorkbooks books;
	CMsExcelWorkbook book;
	CMsExcelWorksheets sheets;
	LPDISPATCH lpDisp = NULL;

	// 打开一个工作簿，如不存在，则新增一个工作簿
	try
	{
		_excelMsApp.put_Visible(FALSE);
		_excelMsApp.put_UserControl(FALSE);

		// 得到工作簿容器
		books.AttachDispatch(_excelMsApp.get_Workbooks());

		// 打开一个工作簿
		lpDisp = books.Open(strSourceFile,
			_vOpt, _vOpt, _vOpt, _vOpt, _vOpt,
			_vOpt, _vOpt, _vOpt, _vOpt, _vOpt,
			_vOpt, _vOpt, _vOpt, _vOpt);

		if (NULL == lpDisp)
		{
			return ePrintDocumentBeenOpened;
		}

		book.AttachDispatch(lpDisp);

		// 得到工作簿中的Sheet的容器
		sheets.AttachDispatch(book.get_Sheets());
		
		// 设置横纵向及纸张大小
		for (long l = 0; l < sheets.get_Count(); l++)
		{
			CMsExcelWorksheet sheet;
			lpDisp = sheets.get_Item(COleVariant(l + 1));
			sheet.AttachDispatch(lpDisp);
			lpDisp = sheet.get_PageSetup();
			CMsExcelPageSetup pageSetup;
			pageSetup.AttachDispatch(lpDisp);
			pageSetup.put_PaperSize(dmPaperType);
			pageSetup.put_Orientation(dmOrient);

			pageSetup.ReleaseDispatch();
			sheet.ReleaseDispatch();
		}

		// 打印表格
		if (strOutputFilename == L"")
		{
			sheets.PrintOut(_vOpt, _vOpt, COleVariant(lCopies), _vOpt, _vOpt, _vOpt, _vOpt, _vOpt);
		}
		else
		{
			sheets.PrintOut(_vOpt, _vOpt, COleVariant(lCopies), _vOpt, _vOpt, COleVariant((short)true), _vOpt, COleVariant(strOutputFilename));
		}

		// 释放资源
		sheets.ReleaseDispatch();
		book.Save();
		book.Close(_vOpt, _vOpt, _vOpt);
		book.ReleaseDispatch();
		books.Close();
		books.ReleaseDispatch();
		CCommonMethod::ResetDeafultPrinter(strDefaultPrinter);

		return ePrintOk;
	}
	catch (...)
	{
		// 释放资源
		sheets.ReleaseDispatch();
		book.Save();
		book.Close(_vOpt, _vOpt, _vOpt);
		book.ReleaseDispatch();
		books.Close();
		books.ReleaseDispatch();
		CCommonMethod::ResetDeafultPrinter(strDefaultPrinter);

		return ePrintOfficeThrowException;
	}
}

PrintErrorStatus PrintExcel::PrintOutWpsExcel(CString strSourceFile, CString strOutputFilename, CString strDefaultPrinter
	, XlPaperSize dmPaperType, XlPageOrientation dmOrient, long lCopies)
{
	CWpsExcelWorkbooks books;
	CWpsExcelWorkbook book;
	CWpsExcelWorksheets sheets;
	LPDISPATCH lpDisp = NULL;

	// 打开一个工作簿，如不存在，则新增一个工作簿
	try
	{
		_excelWpsApp.put_Visible(FALSE);

		// 得到工作簿容器
		books.AttachDispatch(_excelWpsApp.get_Workbooks());

		// 打开一个工作簿
		lpDisp = books.Open(COleVariant(strSourceFile),
			_vOpt, _vOpt, _vOpt, _vOpt, _vOpt,
			_vOpt, _vOpt, _vOpt, _vOpt, _vOpt,
			_vOpt, _vOpt);

		if (NULL == lpDisp)
		{
			return ePrintDocumentBeenOpened;
		}

		book.AttachDispatch(lpDisp);

		// 得到工作簿中的Sheet的容器
		sheets.AttachDispatch(book.get_Worksheets());

		// 设置横纵向及纸张大小
		for (long l = 0; l < sheets.get_Count(); l++)
		{
			CWpsExcelWorksheet sheet;
			lpDisp = sheets.get_Item(COleVariant(l + 1));
			sheet.AttachDispatch(lpDisp);
			lpDisp = sheet.get_PageSetup();
			CWpsExcelPageSetup pageSetup;
			pageSetup.AttachDispatch(lpDisp);
			pageSetup.put_PaperSize(long(dmPaperType));
			pageSetup.put_Orientation(COleVariant(long(dmOrient)));

			pageSetup.ReleaseDispatch();
			sheet.ReleaseDispatch();
		}

		// 打印表格
		if (strOutputFilename == L"")
		{
			sheets.PrintOut(_vOpt, _vOpt, COleVariant(lCopies), _vOpt, _vOpt, _vOpt, _vOpt, _vOpt, FALSE, 1, 1, 0, 0, FALSE, 0, FALSE, 0);
		}
		else
		{
			sheets.PrintOut(_vOpt, _vOpt, COleVariant(lCopies), _vOpt, _vOpt, COleVariant((short)true), _vOpt, COleVariant(strOutputFilename), FALSE, 1, 1, 0, 0, FALSE, 0, FALSE, 0);
		}

		// 释放资源
		sheets.ReleaseDispatch();
		book.Save();
		book.Close(_vOpt, _vOpt, _vOpt);
		book.ReleaseDispatch();
		books.Close(_vOpt);
		books.ReleaseDispatch();
		CCommonMethod::ResetDeafultPrinter(strDefaultPrinter);

		return ePrintOk;
	}
	catch (...)
	{
		// 释放资源
		sheets.ReleaseDispatch();
		book.Save();
		book.Close(_vOpt, _vOpt, _vOpt);
		book.ReleaseDispatch();
		books.Close(_vOpt);
		books.ReleaseDispatch();
		CCommonMethod::ResetDeafultPrinter(strDefaultPrinter);

		return ePrintWPSVersionBelow2010;
	}
}
