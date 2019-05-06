// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CExplorer 包装器类

class CExplorer : public COleDispatchDriver
{
public:
	CExplorer() {} // 调用 COleDispatchDriver 默认构造函数
	CExplorer(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CExplorer(const CExplorer& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _Explorer 方法
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
	LPDISPATCH get_CurrentFolder()
	{
		LPDISPATCH result;
		InvokeHelper(0x2101, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void putref_CurrentFolder(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0x2101, DISPATCH_PROPERTYPUTREF, VT_EMPTY, nullptr, parms, newValue);
	}
	void Close()
	{
		InvokeHelper(0x2103, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Display()
	{
		InvokeHelper(0x2104, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	CString get_Caption()
	{
		CString result;
		InvokeHelper(0x2111, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	VARIANT get_CurrentView()
	{
		VARIANT result;
		InvokeHelper(0x2200, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, nullptr);
		return result;
	}
	void put_CurrentView(VARIANT& newValue)
	{
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x2200, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, &newValue);
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
	LPDISPATCH get_Panes()
	{
		LPDISPATCH result;
		InvokeHelper(0x2201, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Selection()
	{
		LPDISPATCH result;
		InvokeHelper(0x2202, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
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
	BOOL IsPaneVisible(long Pane)
	{
		BOOL result;
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x2203, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Pane);
		return result;
	}
	void ShowPane(long Pane, BOOL Visible)
	{
		static BYTE parms[] = VTS_I4 VTS_BOOL;
		InvokeHelper(0x2204, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Pane, Visible);
	}
	LPDISPATCH get_Views()
	{
		LPDISPATCH result;
		InvokeHelper(0x3109, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_HTMLDocument()
	{
		LPDISPATCH result;
		InvokeHelper(0xfa92, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void SelectFolder(LPDISPATCH Folder)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfab1, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Folder);
	}
	void DeselectFolder(LPDISPATCH Folder)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfab2, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Folder);
	}
	BOOL IsFolderSelected(LPDISPATCH Folder)
	{
		BOOL result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfab3, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Folder);
		return result;
	}
	LPDISPATCH get_NavigationPane()
	{
		LPDISPATCH result;
		InvokeHelper(0xfbb3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void ClearSearch()
	{
		InvokeHelper(0xfbcd, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Search(LPCTSTR Query, long SearchScope)
	{
		static BYTE parms[] = VTS_BSTR VTS_I4;
		InvokeHelper(0xfa65, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Query, SearchScope);
	}
	BOOL IsItemSelectableInView(LPDISPATCH Item)
	{
		BOOL result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc35, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Item);
		return result;
	}
	void AddToSelection(LPDISPATCH Item)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc36, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Item);
	}
	void RemoveFromSelection(LPDISPATCH Item)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc37, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Item);
	}
	void SelectAllItems()
	{
		InvokeHelper(0xfc38, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void ClearSelection()
	{
		InvokeHelper(0xfc39, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	LPDISPATCH get_AccountSelector()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc71, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_AttachmentSelection()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc78, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_ActiveInlineResponse()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc93, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_ActiveInlineResponseWordEditor()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_DisplayMode()
	{
		long result;
		InvokeHelper(0xfc97, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_PreviewPane()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc9f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}

	// _Explorer 属性
public:

};
