// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CMAPIFolderEvents_12 包装器类

class CMAPIFolderEvents_12 : public COleDispatchDriver
{
public:
	CMAPIFolderEvents_12() {} // 调用 COleDispatchDriver 默认构造函数
	CMAPIFolderEvents_12(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CMAPIFolderEvents_12(const CMAPIFolderEvents_12& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// MAPIFolderEvents_12 方法
public:
	void BeforeFolderMove(LPDISPATCH MoveTo, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xfba8, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, MoveTo, Cancel);
	}
	void BeforeItemMove(LPDISPATCH Item, LPDISPATCH MoveTo, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xfba9, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Item, MoveTo, Cancel);
	}

	// MAPIFolderEvents_12 属性
public:

};
