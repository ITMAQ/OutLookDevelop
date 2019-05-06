// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CInspector 包装器类

class CInspector : public COleDispatchDriver
{
public:
	CInspector() {} // 调用 COleDispatchDriver 默认构造函数
	CInspector(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CInspector(const CInspector& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _Inspector 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0xf000, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_Class()
	{
		long result;
		InvokeHelper(0xf00a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Session()
	{
		LPDISPATCH result;
		InvokeHelper(0xf00b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0xf001, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_CommandBars()
	{
		LPDISPATCH result;
		InvokeHelper(0x2100, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_CurrentItem()
	{
		LPDISPATCH result;
		InvokeHelper(0x2102, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_EditorType()
	{
		long result;
		InvokeHelper(0x2110, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_ModifiedFormPages()
	{
		LPDISPATCH result;
		InvokeHelper(0x2106, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void Close(long SaveMode)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x2103, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, SaveMode);
	}
	void Display(VARIANT& Modal)
	{
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x2104, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &Modal);
	}
	void HideFormPage(LPCTSTR PageName)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x2108, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, PageName);
	}
	BOOL IsWordMail()
	{
		BOOL result;
		InvokeHelper(0x2105, DISPATCH_METHOD, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void SetCurrentFormPage(LPCTSTR PageName)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x210c, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, PageName);
	}
	void ShowFormPage(LPCTSTR PageName)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x2109, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, PageName);
	}
	LPDISPATCH get_HTMLEditor()
	{
		LPDISPATCH result;
		InvokeHelper(0x210e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_WordEditor()
	{
		LPDISPATCH result;
		InvokeHelper(0x210f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_Caption()
	{
		CString result;
		InvokeHelper(0x2111, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	long get_Height()
	{
		long result;
		InvokeHelper(0x2114, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_Height(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x2114, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_Left()
	{
		long result;
		InvokeHelper(0x2115, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_Left(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x2115, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_Top()
	{
		long result;
		InvokeHelper(0x2116, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_Top(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x2116, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_Width()
	{
		long result;
		InvokeHelper(0x2117, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_Width(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x2117, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_WindowState()
	{
		long result;
		InvokeHelper(0x2112, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_WindowState(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x2112, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	void Activate()
	{
		InvokeHelper(0x2113, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void SetControlItemProperty(LPDISPATCH Control, LPCTSTR PropertyName)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR;
		InvokeHelper(0xfac9, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Control, PropertyName);
	}
	LPDISPATCH NewFormRegion()
	{
		LPDISPATCH result;
		InvokeHelper(0xfbed, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH OpenFormRegion(LPCTSTR Path)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xfbff, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Path);
		return result;
	}
	void SaveFormRegion(LPDISPATCH Page, LPCTSTR FileName)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR;
		InvokeHelper(0xfc00, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Page, FileName);
	}
	LPDISPATCH get_AttachmentSelection()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc78, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void SetSchedulingStartTime(DATE Start)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0xfc87, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Start);
	}

	// _Inspector 属性
public:

};
