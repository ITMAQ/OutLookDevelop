// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CExplorerEvents_10 包装器类

class CExplorerEvents_10 : public COleDispatchDriver
{
public:
	CExplorerEvents_10() {} // 调用 COleDispatchDriver 默认构造函数
	CExplorerEvents_10(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CExplorerEvents_10(const CExplorerEvents_10& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// ExplorerEvents_10 方法
public:
	STDMETHOD(Activate)()
	{
		HRESULT result;
		InvokeHelper(0xf001, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	STDMETHOD(FolderSwitch)()
	{
		HRESULT result;
		InvokeHelper(0xf002, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	STDMETHOD(BeforeFolderSwitch)(LPDISPATCH NewFolder, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xf003, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, NewFolder, Cancel);
		return result;
	}
	STDMETHOD(ViewSwitch)()
	{
		HRESULT result;
		InvokeHelper(0xf004, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	STDMETHOD(BeforeViewSwitch)(VARIANT& NewView, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL;
		InvokeHelper(0xf005, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &NewView, Cancel);
		return result;
	}
	STDMETHOD(Deactivate)()
	{
		HRESULT result;
		InvokeHelper(0xf006, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	STDMETHOD(SelectionChange)()
	{
		HRESULT result;
		InvokeHelper(0xf007, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	STDMETHOD(Close)()
	{
		HRESULT result;
		InvokeHelper(0xf008, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	STDMETHOD(BeforeMaximize)(BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL;
		InvokeHelper(0xfa11, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Cancel);
		return result;
	}
	STDMETHOD(BeforeMinimize)(BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL;
		InvokeHelper(0xfa12, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Cancel);
		return result;
	}
	STDMETHOD(BeforeMove)(BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL;
		InvokeHelper(0xfa13, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Cancel);
		return result;
	}
	STDMETHOD(BeforeSize)(BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL;
		InvokeHelper(0xfa14, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Cancel);
		return result;
	}
	void BeforeItemCopy(BOOL * Cancel)
	{
		static BYTE parms[] = VTS_PBOOL;
		InvokeHelper(0xfa0e, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Cancel);
	}
	void BeforeItemCut(BOOL * Cancel)
	{
		static BYTE parms[] = VTS_PBOOL;
		InvokeHelper(0xfa0f, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Cancel);
	}
	void BeforeItemPaste(VARIANT * ClipboardContent, LPDISPATCH Target, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xfa10, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, ClipboardContent, Target, Cancel);
	}
	STDMETHOD(AttachmentSelectionChange)()
	{
		HRESULT result;
		InvokeHelper(0xfc79, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	void InlineResponse(LPDISPATCH Item)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc92, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Item);
	}
	void InlineResponseClose()
	{
		InvokeHelper(0xfc96, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	STDMETHOD(DisplayModeChange)(long DisplayMode)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xfc98, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, DisplayMode);
		return result;
	}

	// ExplorerEvents_10 属性
public:

};
