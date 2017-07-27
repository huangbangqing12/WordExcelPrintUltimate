
// WordExcelPrintUltimateDlg.h : 头文件
//

#pragma once

class CWordExcelPrintUltimateDlgAutoProxy;


// CWordExcelPrintUltimateDlg 对话框
class CWordExcelPrintUltimateDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CWordExcelPrintUltimateDlg);
	friend class CWordExcelPrintUltimateDlgAutoProxy;

// 构造
public:
	CWordExcelPrintUltimateDlg(CWnd* pParent = NULL);	// 标准构造函数
	virtual ~CWordExcelPrintUltimateDlg();

// 对话框数据
	enum { IDD = IDD_WORDEXCELPRINTULTIMATE_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	CWordExcelPrintUltimateDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedWord();
	afx_msg void OnBnClickedExcel();
};
