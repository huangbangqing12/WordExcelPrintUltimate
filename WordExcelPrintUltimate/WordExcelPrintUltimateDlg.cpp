
// WordExcelPrintUltimateDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "WordExcelPrintUltimate.h"
#include "WordExcelPrintUltimateDlg.h"
#include "DlgProxy.h"
#include "afxdialogex.h"
#include "CommonMethod/PrintWord.h"
#include "CommonMethod/PrintExcel.h"
#include "CommonMethod/CommonMethod.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CWordExcelPrintUltimateDlg 对话框


IMPLEMENT_DYNAMIC(CWordExcelPrintUltimateDlg, CDialogEx);

CWordExcelPrintUltimateDlg::CWordExcelPrintUltimateDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CWordExcelPrintUltimateDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
}

CWordExcelPrintUltimateDlg::~CWordExcelPrintUltimateDlg()
{
	// 如果该对话框有自动化代理，则
	//  将此代理指向该对话框的后向指针设置为 NULL，以便
	//  此代理知道该对话框已被删除。
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;
}

void CWordExcelPrintUltimateDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CWordExcelPrintUltimateDlg, CDialogEx)
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDWord, &CWordExcelPrintUltimateDlg::OnBnClickedWord)
	ON_BN_CLICKED(IDEXCEL, &CWordExcelPrintUltimateDlg::OnBnClickedExcel)
END_MESSAGE_MAP()


// CWordExcelPrintUltimateDlg 消息处理程序

BOOL CWordExcelPrintUltimateDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO:  在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CWordExcelPrintUltimateDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CWordExcelPrintUltimateDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 当用户关闭 UI 时，如果控制器仍保持着它的某个
//  对象，则自动化服务器不应退出。  这些
//  消息处理程序确保如下情形: 如果代理仍在使用，
//  则将隐藏 UI；但是在关闭对话框时，
//  对话框仍然会保留在那里。

void CWordExcelPrintUltimateDlg::OnClose()
{
	if (CanExit())
		CDialogEx::OnClose();
}

void CWordExcelPrintUltimateDlg::OnOK()
{
	if (CanExit())
		CDialogEx::OnOK();
}

void CWordExcelPrintUltimateDlg::OnCancel()
{
	if (CanExit())
		CDialogEx::OnCancel();
}

BOOL CWordExcelPrintUltimateDlg::CanExit()
{
	// 如果代理对象仍保留在那里，则自动化
	//  控制器仍会保持此应用程序。
	//  使对话框保留在那里，但将其 UI 隐藏起来。
	if (m_pAutoProxy != NULL)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}


void CWordExcelPrintUltimateDlg::OnBnClickedWord()
{
	CCommonMethod::Sample();
	PrintWord pw;
	//开始打印
	pw.BeginPrintOutWord();

	//执行打印
	CString strAppPath = CCommonMethod::GetPath();
	CString strSourceFile = strAppPath + L"设计总说明.doc";
	CString strOutputFilename = strAppPath + L"word.pdf";
	pw.PrintOutWord(strSourceFile, strOutputFilename, L"Microsoft Print to PDF", WdPaperSize::wpsPaperA3, WdOrientation::wpsOrientLandscape, 1);

	strOutputFilename = strAppPath + L"word1.pdf";
	pw.PrintOutWord(strSourceFile, strOutputFilename, L"Microsoft Print to PDF", WdPaperSize::wpsPaperA4, WdOrientation::wpsOrientPortrait, 1);

	//结束打印
	pw.EndPrintOutWord();

	CDialogEx::OnOK();
}


void CWordExcelPrintUltimateDlg::OnBnClickedExcel()
{
	PrintExcel pe;

	//开始打印
	pe.BeginPrintOutExcel();

	//执行打印
	CString strAppPath = CCommonMethod::GetPath();
	CString strSourceFile = strAppPath + L"设计总说明.xls";
	CString strOutputFilename = strAppPath + L"excel.pdf";

	pe.PrintOutExcel(strSourceFile, strOutputFilename, L"Microsoft Print to PDF", XlPaperSize::xlPaperA3, XlPageOrientation::xlLandscape, 1);

	strOutputFilename = strAppPath + L"excel1.pdf";
	pe.PrintOutExcel(strSourceFile, strOutputFilename, L"Microsoft Print to PDF", XlPaperSize::xlPaperA4, XlPageOrientation::xlPortrait, 1);

	//结束打印
	pe.EndPrintOutExcel();

	CDialogEx::OnOK();
}
