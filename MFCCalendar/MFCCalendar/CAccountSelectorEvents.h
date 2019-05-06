// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CAccountSelectorEvents 包装器类

class CAccountSelectorEvents : public COleDispatchDriver
{
public:
	CAccountSelectorEvents() {} // 调用 COleDispatchDriver 默认构造函数
	CAccountSelectorEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAccountSelectorEvents(const CAccountSelectorEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// AccountSelectorEvents 方法
public:
	void SelectedAccountChange(LPDISPATCH SelectedAccount)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc73, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, SelectedAccount);
	}

	// AccountSelectorEvents 属性
public:

};
