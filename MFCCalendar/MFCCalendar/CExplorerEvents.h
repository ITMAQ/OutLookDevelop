// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CExplorerEvents 包装器类

class CExplorerEvents : public COleDispatchDriver
{
public:
	CExplorerEvents() {} // 调用 COleDispatchDriver 默认构造函数
	CExplorerEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CExplorerEvents(const CExplorerEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// ExplorerEvents 方法
public:
	void Activate()
	{
		InvokeHelper(0xf001, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void FolderSwitch()
	{
		InvokeHelper(0xf002, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void BeforeFolderSwitch(LPDISPATCH NewFolder, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xf003, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, NewFolder, Cancel);
	}
	void ViewSwitch()
	{
		InvokeHelper(0xf004, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void BeforeViewSwitch(VARIANT& NewView, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL;
		InvokeHelper(0xf005, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &NewView, Cancel);
	}
	void Deactivate()
	{
		InvokeHelper(0xf006, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void SelectionChange()
	{
		InvokeHelper(0xf007, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Close()
	{
		InvokeHelper(0xf008, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// ExplorerEvents 属性
public:

};
