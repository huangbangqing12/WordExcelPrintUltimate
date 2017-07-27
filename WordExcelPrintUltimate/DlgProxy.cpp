
// DlgProxy.cpp : 实现文件
//

#include "stdafx.h"
#include "WordExcelPrintUltimate.h"
#include "DlgProxy.h"
#include "WordExcelPrintUltimateDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CWordExcelPrintUltimateDlgAutoProxy

IMPLEMENT_DYNCREATE(CWordExcelPrintUltimateDlgAutoProxy, CCmdTarget)

CWordExcelPrintUltimateDlgAutoProxy::CWordExcelPrintUltimateDlgAutoProxy()
{
	EnableAutomation();
	
	// 为使应用程序在自动化对象处于活动状态时一直保持 
	//	运行，构造函数调用 AfxOleLockApp。
	AfxOleLockApp();

	// 通过应用程序的主窗口指针
	//  来访问对话框。  设置代理的内部指针
	//  指向对话框，并设置对话框的后向指针指向
	//  该代理。
	ASSERT_VALID(AfxGetApp()->m_pMainWnd);
	if (AfxGetApp()->m_pMainWnd)
	{
		ASSERT_KINDOF(CWordExcelPrintUltimateDlg, AfxGetApp()->m_pMainWnd);
		if (AfxGetApp()->m_pMainWnd->IsKindOf(RUNTIME_CLASS(CWordExcelPrintUltimateDlg)))
		{
			m_pDialog = reinterpret_cast<CWordExcelPrintUltimateDlg*>(AfxGetApp()->m_pMainWnd);
			m_pDialog->m_pAutoProxy = this;
		}
	}
}

CWordExcelPrintUltimateDlgAutoProxy::~CWordExcelPrintUltimateDlgAutoProxy()
{
	// 为了在用 OLE 自动化创建所有对象后终止应用程序，
	//	析构函数调用 AfxOleUnlockApp。
	//  除了做其他事情外，这还将销毁主对话框
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CWordExcelPrintUltimateDlgAutoProxy::OnFinalRelease()
{
	// 释放了对自动化对象的最后一个引用后，将调用
	// OnFinalRelease。  基类将自动
	// 删除该对象。  在调用该基类之前，请添加您的
	// 对象所需的附加清理代码。

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CWordExcelPrintUltimateDlgAutoProxy, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CWordExcelPrintUltimateDlgAutoProxy, CCmdTarget)
END_DISPATCH_MAP()

// 注意: 我们添加了对 IID_IWordExcelPrintUltimate 的支持
//  以支持来自 VBA 的类型安全绑定。  此 IID 必须同附加到 .IDL 文件中的
//  调度接口的 GUID 匹配。

// {D835C838-DC60-4B8B-AD42-FB9E54417843}
static const IID IID_IWordExcelPrintUltimate =
{ 0xD835C838, 0xDC60, 0x4B8B, { 0xAD, 0x42, 0xFB, 0x9E, 0x54, 0x41, 0x78, 0x43 } };

BEGIN_INTERFACE_MAP(CWordExcelPrintUltimateDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CWordExcelPrintUltimateDlgAutoProxy, IID_IWordExcelPrintUltimate, Dispatch)
END_INTERFACE_MAP()

// IMPLEMENT_OLECREATE2 宏在此项目的 StdAfx.h 中定义
// {5CE108BC-836E-41CE-9946-02DD894E5A53}
IMPLEMENT_OLECREATE2(CWordExcelPrintUltimateDlgAutoProxy, "WordExcelPrintUltimate.Application", 0x5ce108bc, 0x836e, 0x41ce, 0x99, 0x46, 0x2, 0xdd, 0x89, 0x4e, 0x5a, 0x53)


// CWordExcelPrintUltimateDlgAutoProxy 消息处理程序
