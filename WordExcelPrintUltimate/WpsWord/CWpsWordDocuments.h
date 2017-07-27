// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

//#import "C:\\Program Files (x86)\\Kingsoft\\WPS Office Personal\\office6\\wpscore.dll" no_namespace
// CWpsWordDocuments 包装器类

class CWpsWordDocuments : public COleDispatchDriver
{
public:
	CWpsWordDocuments(){} // 调用 COleDispatchDriver 默认构造函数
	CWpsWordDocuments(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWpsWordDocuments(const CWpsWordDocuments& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// Documents 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void zimp_DispObj_Reserved1()
	{
		InvokeHelper(0xfffff01, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void zimp_DispObj_Reserved2()
	{
		InvokeHelper(0xfffff02, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void zimp_DispObj_Reserved3()
	{
		InvokeHelper(0xfffff03, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void zimp_DispObj_Reserved4()
	{
		InvokeHelper(0xfffff04, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void zimp_DispObj_Reserved5()
	{
		InvokeHelper(0xfffff05, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	CString get_Description()
	{
		CString result;
		InvokeHelper(0xfffff06, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(VARIANT * Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	void Close(VARIANT * SaveChanges, VARIANT * OriginalFormat, VARIANT * RouteDocument)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
		InvokeHelper(0x451, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SaveChanges, OriginalFormat, RouteDocument);
	}
	void Save(BOOL NoPrompt, LPCTSTR OriginalFormat)
	{
		static BYTE parms[] = VTS_BOOL VTS_BSTR;
		InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NoPrompt, OriginalFormat);
	}
	LPDISPATCH Add(VARIANT * Template, BOOL NewTemplate, long DocumentType, BOOL Visible)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT VTS_BOOL VTS_I4 VTS_BOOL;
		InvokeHelper(0xe, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Template, NewTemplate, DocumentType, Visible);
		return result;
	}
	LPDISPATCH Open(LPCTSTR FileName, BOOL ConfirmConversions, BOOL ReadOnly, BOOL AddToRecentFiles, LPCTSTR PasswordDocument, LPCTSTR PasswordTemplate, BOOL Revert, LPCTSTR WritePasswordDocument, LPCTSTR WritePasswordTemplate, long Format, long Encoding, BOOL Visible, BOOL OpenAndRepair, long DocumentDirection, BOOL NoEncodingDialog)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_BOOL VTS_BOOL VTS_BOOL VTS_BSTR VTS_BSTR VTS_BOOL VTS_BSTR VTS_BSTR VTS_I4 VTS_I4 VTS_BOOL VTS_BOOL VTS_I4 VTS_BOOL;
		InvokeHelper(0x12, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog);
		return result;
	}
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}

	// Documents 属性
public:

};
