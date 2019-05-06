// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CNameSpaceEvents 包装器类

class CNameSpaceEvents : public COleDispatchDriver
{
public:
	CNameSpaceEvents() {} // 调用 COleDispatchDriver 默认构造函数
	CNameSpaceEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CNameSpaceEvents(const CNameSpaceEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// NameSpaceEvents 方法
public:
	void OptionsPagesAdd(LPDISPATCH Pages, LPDISPATCH Folder)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH;
		InvokeHelper(0xf005, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Pages, Folder);
	}
	void AutoDiscoverComplete()
	{
		InvokeHelper(0xfc2d, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// NameSpaceEvents 属性
public:

};
