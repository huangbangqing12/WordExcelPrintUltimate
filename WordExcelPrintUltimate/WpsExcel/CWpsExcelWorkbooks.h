// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

//#import "C:\\Program Files (x86)\\Kingsoft\\WPS Office Personal\\office6\\etapp.dll" no_namespace
// CWpsExcelWorkbooks 包装器类

class CWpsExcelWorkbooks : public COleDispatchDriver
{
public:
	CWpsExcelWorkbooks(){} // 调用 COleDispatchDriver 默认构造函数
	CWpsExcelWorkbooks(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWpsExcelWorkbooks(const CWpsExcelWorkbooks& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// Workbooks 方法
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
		InvokeHelper(0x221000, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Item(VARIANT& Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x221001, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH get__Default(VARIANT& Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH Add(VARIANT& Template)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x222000, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Template);
		return result;
	}
	void Close(VARIANT& SaveChanges)
	{
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x222001, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &SaveChanges);
	}
	LPDISPATCH Open(VARIANT& Filename, VARIANT& UpdateLinks, VARIANT& ReadOnly, VARIANT& Format, VARIANT& Password, VARIANT& WriteResPassword, VARIANT& IgnoreReadOnlyRecommended, VARIANT& Origin, VARIANT& Delimiter, VARIANT& Editable, VARIANT& Notify, VARIANT& Converter, VARIANT& AddToMru)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0x222002, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Filename, &UpdateLinks, &ReadOnly, &Format, &Password, &WriteResPassword, &IgnoreReadOnlyRecommended, &Origin, &Delimiter, &Editable, &Notify, &Converter, &AddToMru);
		return result;
	}
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH OpenText(LPCTSTR Filename, VARIANT& Origin, VARIANT& StartRow, VARIANT& DataType, VARIANT& ConsecutiveDelimiter, VARIANT& Tab, VARIANT& Semicolon, VARIANT& Comma, VARIANT& Space, VARIANT& Other, VARIANT& OtherChar, VARIANT& FieldInfo, VARIANT& TextVisualLayout, VARIANT& DecimalSeparator, VARIANT& ThousandsSeparator, VARIANT& TrailingMinusNumbers, VARIANT& Local)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0x222003, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Filename, &Origin, &StartRow, &DataType, &ConsecutiveDelimiter, &Tab, &Semicolon, &Comma, &Space, &Other, &OtherChar, &FieldInfo, &TextVisualLayout, &DecimalSeparator, &ThousandsSeparator, &TrailingMinusNumbers, &Local);
		return result;
	}
	LPDISPATCH OpenDatabase(LPCTSTR Filename, VARIANT& CommandText, VARIANT& CommandType, VARIANT& BackgroundQuery, VARIANT& ImportDataAs)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0x222004, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Filename, &CommandText, &CommandType, &BackgroundQuery, &ImportDataAs);
		return result;
	}

	// Workbooks 属性
public:

};
