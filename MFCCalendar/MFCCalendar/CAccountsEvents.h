// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CAccountsEvents 包装器类

class CAccountsEvents : public COleDispatchDriver
{
public:
	CAccountsEvents() {} // 调用 COleDispatchDriver 默认构造函数
	CAccountsEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAccountsEvents(const CAccountsEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// AccountsEvents 方法
public:
	void AutoDiscoverComplete(LPDISPATCH Account)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc6c, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Account);
	}

	// AccountsEvents 属性
public:

};
