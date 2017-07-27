
#pragma once

#include "../MsExcel/CMsExcelWorksheet.h"
#include "../MsExcel/CMsExcelWorksheets.h"
#include "../MsExcel/CMsExcelWorkbook.h"
#include "../MsExcel/CMsExcelWorkbooks.h"
#include "../MsExcel/CMsExcelApplication.h"
#include "../MsExcel/CMsExcelPageSetup.h"

#include "../WpsExcel/CWpsExcelWorksheet.h"
#include "../WpsExcel/CWpsExcelWorksheets.h"
#include "../WpsExcel/CWpsExcelWorkbook.h"
#include "../WpsExcel/CWpsExcelWorkbooks.h"
#include "../WpsExcel/CWpsExcelApplication.h"
#include "../WpsExcel/CWpsExcelPageSetup.h"

#include "CommonMethod.h"

class PrintExcel
{
public:
	PrintExcel();
	~PrintExcel();
	
	/**
	*  @brief    创建word打印实例
	*
	*  @return   返回错误的状态
	*/
	PrintErrorStatus BeginPrintOutExcel();
	
	/**
	*  @brief    执行打印
	*
	*  @param    CString strSourceFile			要打印的源文件的完整路径
	*  @param    CString strOutputFilename		要输出到的指定目录完整路径，若直接在打印机打印，则输入L""即可
	*  @param    CString strPrinterName			指定打印机的名称，若按照系统默认打印机打印，则输入L""即可
	*  @param    XlPaperSize dmPaperType		打印的纸张大小
	*  @param    XlPageOrientation dmOrient		打印的方向（横向打印OR纵向打印）
	*  @param    long lCopies					打印的份数
	*  @return   PrintErrorStatus				返回错误的状态
	*/
	PrintErrorStatus PrintOutExcel(CString strSourceFile, CString strOutputFilename, CString strPrinterName, XlPaperSize dmPaperType, XlPageOrientation dmOrient, long lCopies);
	
	/**
	*  @brief    销毁实例
	*
	*  @return   void
	*/
	void EndPrintOutExcel();
private:
	
	/**
	*  @brief    打印微软 excel
	*
	*  @param    CString strSourceFile			要打印的源文件的完整路径
	*  @param    CString strOutputFilename		要输出到的指定目录完整路径，若直接在打印机打印，则输入L""即可
	*  @param    CString strDefaultPrinter		默认打印机的名称
	*  @param    XlPaperSize dmPaperType		打印的纸张大小
	*  @param    XlPageOrientation dmOrient		打印的方向（横向打印OR纵向打印）
	*  @param    long lCopies					打印的份数
	*  @return   PrintErrorStatus				返回错误的状态
	*/
	PrintErrorStatus PrintOutMsExcel(CString strSourceFile, CString strOutputFilename, CString strDefaultPrinter, XlPaperSize dmPaperType, XlPageOrientation dmOrient, long lCopies);
	
	/**
	*  @brief    打印金山 excel
	*
	*  @param    CString strSourceFile			要打印的源文件的完整路径
	*  @param    CString strOutputFilename		要输出到的指定目录完整路径，若直接在打印机打印，则输入L""即可
	*  @param    CString strDefaultPrinter		默认打印机的名称
	*  @param    XlPaperSize dmPaperType		打印的纸张大小
	*  @param    XlPageOrientation dmOrient		打印的方向（横向打印OR纵向打印）
	*  @param    long lCopies					打印的份数
	*  @return   PrintErrorStatus				返回错误的状态
	*/
	PrintErrorStatus PrintOutWpsExcel(CString strSourceFile, CString strOutputFilename, CString strDefaultPrinter, XlPaperSize dmPaperType, XlPageOrientation dmOrient, long lCopies);
	
	/**
	*  @brief    销毁实例
	*
	*  @return   void
	*/
	void QuitMsDispatch();
	
	/**
	*  @brief    销毁实例
	*
	*  @return   void
	*/
	void QuitWpsDispatch();

private:
	CMsExcelApplication _excelMsApp;
	CWpsExcelApplication _excelWpsApp;
	bool _bIsMsExcel;
	COleVariant _vOpt;
};