#include "stdafx.h"
#include "CommonMethod.h"
#include "WINSPOOL.H" 

CString CCommonMethod::GetPath()
{
	CString strPath;

	// 获取acad.exe文件路径
	TCHAR* pBuffer = strPath.GetBuffer(MAX_PATH);
	GetModuleFileName(GetModuleHandle(NULL), pBuffer, MAX_PATH);
	strPath.ReleaseBuffer();

	// 去除acad.exe文件名
	strPath = strPath.Left(strPath.ReverseFind(L'\\') + 1);

	return strPath;
}

void CCommonMethod::AvoidOtherProgrameRunning()
{
	// 解决Visual Studio2008中用excel automation读取Excel文档，有时会程序提示”由于另一个程序正在运行中,此操作无法完成的问题
	AfxOleGetMessageFilter()->EnableBusyDialog(FALSE);
	AfxOleGetMessageFilter()->SetBusyReply(SERVERCALL_RETRYLATER);
	AfxOleGetMessageFilter()->EnableNotRespondingDialog(TRUE);
	AfxOleGetMessageFilter()->SetMessagePendingDelay(-1);
}

CString CCommonMethod::GetDeafultPrinter()
{
	// 得到默认打印设备名称
	CString strDefualtDev = TEXT("");
	PRINTDLG pd;
	LPDEVMODE lpDevMode;
	if (AfxGetApp()->GetPrinterDeviceDefaults(&pd))
	{
		lpDevMode = (LPDEVMODE)GlobalLock(pd.hDevMode);
		if (lpDevMode)
		{
			strDefualtDev = lpDevMode->dmDeviceName;
		}
		GlobalUnlock(pd.hDevMode);
	}
	return strDefualtDev;
}

CString CCommonMethod::SetPrinterPara(CString strActivePrinter)
{
	if (strActivePrinter == L"") return L"";

	// 得到默认打印设备名称
	CString strDefualtDev = TEXT("");
	PRINTDLG pd;
	LPDEVMODE lpDevMode;
	if (AfxGetApp()->GetPrinterDeviceDefaults(&pd))
	{
		lpDevMode = (LPDEVMODE)GlobalLock(pd.hDevMode);
		if (lpDevMode)
		{
			strDefualtDev = lpDevMode->dmDeviceName;
		}
		GlobalUnlock(pd.hDevMode);
	}
	// 指定的打印机名称
	::SetDefaultPrinter(strActivePrinter);
	// 得到刚刚设定的打印机名称
	if (AfxGetApp()->GetPrinterDeviceDefaults(&pd))
	{
		lpDevMode = (LPDEVMODE)GlobalLock(pd.hDevMode);
		if (lpDevMode)
		{
			lpDevMode->dmPaperSize = DMPAPER_A3;    // 设定打印纸张幅面
			lpDevMode->dmOrientation = DMORIENT_LANDSCAPE; // 设定横向打印
			lpDevMode->dmPrintQuality = 600;     // 设定打印机分辨率
		}
		GlobalUnlock(pd.hDevMode);
	}

	return strDefualtDev;
}

void CCommonMethod::ResetDeafultPrinter(CString strDefaultPrinter)
{
	if (strDefaultPrinter != L"")
	{
		/// 还原默认的打印设备设定
		::SetDefaultPrinter(strDefaultPrinter);
	}
}

void CCommonMethod::Sample()
{
	UINT iBuffSize = DeviceCapabilities(L"Microsoft Print to PDF", NULL, DC_PAPERNAMES, NULL, NULL);
	if (iBuffSize > 0)
	{
		TCHAR  szPaperNameList[32][64];
		UINT iFlag = DeviceCapabilities(L"Microsoft Print to PDF", NULL, DC_PAPERNAMES, (LPWSTR)szPaperNameList, NULL);
		if (iFlag != -1)
		{
			for (int i = 0; i < iBuffSize; i++)
			{
				CString str = szPaperNameList[i];
				int j = 0;
				//m_vecPaperName.push_back(szPaperNameList[i]);
			}
		}
	}

	iBuffSize = DeviceCapabilities(L"Microsoft Print to PDF", NULL, DC_PAPERS, NULL, NULL);
	if (iBuffSize > 0)
	{
		WORD arrPaperSize[32];
		memset(arrPaperSize, 0, 32 * sizeof(WORD));
		UINT iFlag = DeviceCapabilities(L"Microsoft Print to PDF", NULL, DC_PAPERS, (LPWSTR)arrPaperSize, NULL);
		if (iFlag != -1)
		{
			for (int i = 0; i < iBuffSize; i++)
			{
				WORD str = arrPaperSize[i];
				int j = 0;
				//m_vecPaperSize.push_back(arrPaperSize[i]);
			}
		}
	}
}
