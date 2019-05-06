// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CFoldersEvents 包装器类

class CFoldersEvents : public COleDispatchDriver
{
public:
	CFoldersEvents() {} // 调用 COleDispatchDriver 默认构造函数
	CFoldersEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFoldersEvents(const CFoldersEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// FoldersEvents 方法
public:
	void FolderAdd(LPDISPATCH Folder)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf001, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Folder);
	}
	void FolderChange(LPDISPATCH Folder)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf002, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Folder);
	}
	void FolderRemove()
	{
		InvokeHelper(0xf003, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// FoldersEvents 属性
public:

};
