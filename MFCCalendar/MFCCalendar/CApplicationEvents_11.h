// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CApplicationEvents_11 包装器类

class CApplicationEvents_11 : public COleDispatchDriver
{
public:
	CApplicationEvents_11() {} // 调用 COleDispatchDriver 默认构造函数
	CApplicationEvents_11(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplicationEvents_11(const CApplicationEvents_11& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// ApplicationEvents_11 方法
public:
	STDMETHOD(ItemSend)(LPDISPATCH Item, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xf002, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Item, Cancel);
		return result;
	}
	STDMETHOD(NewMail)()
	{
		HRESULT result;
		InvokeHelper(0xf003, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	STDMETHOD(Reminder)(LPDISPATCH Item)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf004, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Item);
		return result;
	}
	STDMETHOD(OptionsPagesAdd)(LPDISPATCH Pages)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf005, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Pages);
		return result;
	}
	STDMETHOD(Startup)()
	{
		HRESULT result;
		InvokeHelper(0xf006, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	STDMETHOD(Quit)()
	{
		HRESULT result;
		InvokeHelper(0xf007, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	STDMETHOD(AdvancedSearchComplete)(LPDISPATCH SearchObject)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfa6a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, SearchObject);
		return result;
	}
	STDMETHOD(AdvancedSearchStopped)(LPDISPATCH SearchObject)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfa6b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, SearchObject);
		return result;
	}
	STDMETHOD(MAPILogonComplete)()
	{
		HRESULT result;
		InvokeHelper(0xfa90, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	void NewMailEx(LPCTSTR EntryIDCollection)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xfab5, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, EntryIDCollection);
	}
	STDMETHOD(AttachmentContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH Attachments)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH;
		InvokeHelper(0xfb3e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, CommandBar, Attachments);
		return result;
	}
	void FolderContextMenuDisplay(LPDISPATCH CommandBar, LPDISPATCH Folder)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH;
		InvokeHelper(0xfb42, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, CommandBar, Folder);
	}
	void StoreContextMenuDisplay(LPDISPATCH CommandBar, LPDISPATCH Store)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH;
		InvokeHelper(0xfb43, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, CommandBar, Store);
	}
	void ShortcutContextMenuDisplay(LPDISPATCH CommandBar, LPDISPATCH Shortcut)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH;
		InvokeHelper(0xfb44, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, CommandBar, Shortcut);
	}
	void ViewContextMenuDisplay(LPDISPATCH CommandBar, LPDISPATCH View)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH;
		InvokeHelper(0xfb40, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, CommandBar, View);
	}
	void ItemContextMenuDisplay(LPDISPATCH CommandBar, LPDISPATCH Selection)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH;
		InvokeHelper(0xfb41, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, CommandBar, Selection);
	}
	void ContextMenuClose(long ContextMenu)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xfba6, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, ContextMenu);
	}
	void ItemLoad(LPDISPATCH Item)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfba7, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Item);
	}
	void BeforeFolderSharingDialog(LPDISPATCH FolderToShare, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xfc01, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, FolderToShare, Cancel);
	}

	// ApplicationEvents_11 属性
public:

};
