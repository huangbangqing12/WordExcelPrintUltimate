
#pragma once

#include "../MsWord/CMsWordApplication.h"
#include "../MsWord/CMsWordDocument0.h"
#include "../MsWord/CMsWordDocuments.h"
#include "../MsWord/CMSWordPageSetup.h"

#include "../WpsWord/CWpsWordApplication.h"
#include "../WpsWord/CWpsWordDocument0.h"
#include "../WpsWord/CWpsWordDocuments.h"
#include "../WpsWord/CWpsWordPageSetup.h"

#include "CommonMethod.h"


class PrintWord
{
public:
	PrintWord();
	~PrintWord();
	
	/**
	*  @brief    创建word打印实例
	*
	*  @return   返回错误的状态
	*/
	PrintErrorStatus BeginPrintOutWord();
	
	/**
	*  @brief    执行打印
	*
	*  @param    CString strSourceFile			要打印的源文件的完整路径
	*  @param    CString strOutputFilename		要输出到的指定目录完整路径，若直接在打印机打印，则输入L""即可
	*  @param    CString strPrinterName			指定打印机的名称，若按照系统默认打印机打印，则输入L""即可
	*  @param    WdPaperSize dmPaperType		打印的纸张大小
	*  @param    WdOrientation dmOrient			打印的方向（横向打印OR纵向打印）
	*  @param    long lCopies					打印的份数
	*  @return   PrintErrorStatus				返回错误的状态
	*/
	PrintErrorStatus PrintOutWord(CString strSourceFile, CString strOutputFilename, CString strPrinterName, WdPaperSize dmPaperType, WdOrientation dmOrient, long lCopies);
	
	/**
	*  @brief    销毁实例
	*
	*  @return   void
	*/
	void EndPrintOutWord();


private:
	
	/**
	*  @brief    打印微软 word
	*
	*  @param    CString strSourceFile				要打印的源文件的完整路径
	*  @param    CString strOutputFilename			要输出到的指定目录完整路径，若直接在打印机打印，则输入L""即可
	*  @param    CString strDefaultPrinter			默认打印机的名称
	*  @param    WdPaperSize dmPaperType			打印的纸张大小
	*  @param    WdOrientation dmOrient				打印的方向（横向打印OR纵向打印）
	*  @param    long lCopies						打印的份数
	*  @return   PrintErrorStatus					返回错误的状态
	*/
	PrintErrorStatus PrintOutMSWord(CString strSourceFile, CString strOutputFilename, CString strDefaultPrinter, WdPaperSize dmPaperType, WdOrientation dmOrient, long lCopies);
	
	/**
	*  @brief    打印金山 word
	*
	*  @param    CString strSourceFile				要打印的源文件的完整路径
	*  @param    CString strOutputFilename			要输出到的指定目录完整路径，若直接在打印机打印，则输入L""即可
	*  @param    CString strDefaultPrinter			默认打印机的名称
	*  @param    WdPaperSize dmPaperType			打印的纸张大小
	*  @param    WdOrientation dmOrient				打印的方向（横向打印OR纵向打印）
	*  @param    long lCopies						打印的份数
	*  @return   PrintErrorStatus					返回错误的状态
	*/
	PrintErrorStatus PrintOutWPSWord(CString strSourceFile, CString strOutputFilename, CString strDefaultPrinter, WdPaperSize dmPaperType, WdOrientation dmOrient, long lCopies);
	
	/**
	*  @brief    销毁实例
	*
	*  @return   void
	*/
	void QuitMsDispatch();

private:
	CMsWordApplication _wordApp;		//word实例
	bool _bIsMsWord;					//是否为微软word
	COleVariant _vOpt;
};