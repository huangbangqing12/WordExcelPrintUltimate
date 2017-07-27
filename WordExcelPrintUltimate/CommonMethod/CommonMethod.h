#pragma once

// 错误状态
enum PrintErrorStatus
{
	ePrintOk = 0,

	ePrintOfficeThrowException = 1,		// 打印过程中抛出异常
	ePrintDocumentBeenOpened = 2,	// 模板文档已被打开，请关闭文档！
	ePrintWPSVersionBelow2010 = 3,  // 请安装Wps office2010以上版本
	ePrintNoneWpsOfficeInstall = 4,		// 请安装Office或Wps
};

// word打印横纵向
enum WdOrientation
{
	wpsOrientPortrait = 0,		//纵向
	wpsOrientLandscape = 1		//横向
};

// word打印纸张大小
enum WdPaperSize
{
	wpsPaper10x14 = 0,
	wpsPaper11x17 = 1,
	wpsPaperLetter = 2,
	wpsPaperLetterSmall = 3,
	wpsPaperLegal = 4,
	wpsPaperExecutive = 5,
	wpsPaperA3 = 6,
	wpsPaperA4 = 7,
	wpsPaperA4Small = 8,
	wpsPaperA5 = 9,
	wpsPaperB4 = 10,
	wpsPaperB5 = 11,
	wpsPaperCSheet = 12,
	wpsPaperFanfoldLegalGerman = 13,
	wpsPaperFanfoldStdGerman = 14,
	wpsPaperFanfoldUS = 15,
	wpsPaperFolio = 16,
	wpsPaperLedger = 17,
	wpsPaperNote = 18,
	wpsPaperQuarto = 19,
	wpsPaperStatement = 20,
	wpsPaperTabloid = 21,
	wpsPaperEnvelope9 = 22,
	wpsPaperEnvelope10 = 23,
	wpsPaperEnvelope11 = 24,
	wpsPaperEnvelope14 = 25,
	wpsPaperEnvelopeB4 = 26,
	wpsPaperEnvelopeB5 = 27,
	wpsPaperEnvelopeB6 = 28,
	wpsPaperEnvelopeC3 = 29,
	wpsPaperEnvelopeC4 = 30,
	wpsPaperEnvelopeC5 = 31,
	wpsPaperEnvelopeC6 = 32,
	wpsPaperEnvelopeC65 = 33,
	wpsPaperEnvelopeDL = 34,
	wpsPaperEnvelopeItaly = 35,
	wpsPaperEnvelopeMonarch = 36,
	wpsPaperEnvelopePersonal = 37,
	wpsPaperCustom = 38
};

// excel打印方向
enum XlPageOrientation
{
	xlLandscape = 2,		//横向
	xlPortrait = 1			//纵向
};

// excel打印纸张大小
enum XlPaperSize
{
	xlPaper10x14 = 16,
	xlPaper11x17 = 17,
	xlPaperA3 = 8,
	xlPaperA4 = 9,
	xlPaperA4Small = 10,
	xlPaperA5 = 11,
	xlPaperB4 = 12,
	xlPaperB5 = 13,
	xlPaperCsheet = 24,
	xlPaperDsheet = 25,
	xlPaperEnvelope10 = 20,
	xlPaperEnvelope11 = 21,
	xlPaperEnvelope12 = 22,
	xlPaperEnvelope14 = 23,
	xlPaperEnvelope9 = 19,
	xlPaperEnvelopeB4 = 33,
	xlPaperEnvelopeB5 = 34,
	xlPaperEnvelopeB6 = 35,
	xlPaperEnvelopeC3 = 29,
	xlPaperEnvelopeC4 = 30,
	xlPaperEnvelopeC5 = 28,
	xlPaperEnvelopeC6 = 31,
	xlPaperEnvelopeC65 = 32,
	xlPaperEnvelopeDL = 27,
	xlPaperEnvelopeItaly = 36,
	xlPaperEnvelopeMonarch = 37,
	xlPaperEnvelopePersonal = 38,
	xlPaperEsheet = 26,
	xlPaperExecutive = 7,
	xlPaperFanfoldLegalGerman = 41,
	xlPaperFanfoldStdGerman = 40,
	xlPaperFanfoldUS = 39,
	xlPaperFolio = 14,
	xlPaperLedger = 4,
	xlPaperLegal = 5,
	xlPaperLetter = 1,
	xlPaperLetterSmall = 2,
	xlPaperNote = 18,
	xlPaperQuarto = 15,
	xlPaperStatement = 6,
	xlPaperTabloid = 3,
	xlPaperUser = 256
};

class CCommonMethod
{
public:
	
	/**
	*  @brief    获得exe路径
	*
	*  @return   CString
	*/
	static CString GetPath();

	
	/**
	*  @brief    解决Visual Studio2008中用excel automation读取Excel文档，有时会程序提示”由于另一个程序正在运行中,此操作无法完成的问题
	*
	*  @return   void
	*/
	static void AvoidOtherProgrameRunning();

	
	/**
	*  @brief    获得默认打印机
	*
	*  @return   CString
	*/
	static CString GetDeafultPrinter();

	
	/**
	*  @brief    设置默认打印机
	*
	*  @param    CString strActivePrinter
	*  @return   CString
	*/
	static CString SetPrinterPara(CString strActivePrinter);

	
	/**
	*  @brief    重置会默认打印机
	*
	*  @param    CString strDefaultPrinter
	*  @return   void
	*/
	static void ResetDeafultPrinter(CString strDefaultPrinter);

	static void Sample();
};