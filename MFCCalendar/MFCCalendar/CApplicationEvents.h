// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CApplicationEvents 包装器类

class CApplicationEvents : public COleDispatchDriver
{
public:
	CApplicationEvents() {} // 调用 COleDispatchDriver 默认构造函数
	CApplicationEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplicationEvents(const CApplicationEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// ApplicationEvents 方法
public:
	void ItemSend(LPDISPATCH Item, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xf002, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Item, Cancel);
	}
	void NewMail()
	{
		InvokeHelper(0xf003, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Reminder(LPDISPATCH Item)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf004, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Item);
	}
	void OptionsPagesAdd(LPDISPATCH Pages)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf005, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Pages);
	}
	void Startup()
	{
		InvokeHelper(0xf006, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Quit()
	{
		InvokeHelper(0xf007, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// ApplicationEvents 属性
public:

};
