// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CApplicationEvents_10 包装器类

class CApplicationEvents_10 : public COleDispatchDriver
{
public:
	CApplicationEvents_10() {} // 调用 COleDispatchDriver 默认构造函数
	CApplicationEvents_10(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplicationEvents_10(const CApplicationEvents_10& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// ApplicationEvents_10 方法
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
	void AdvancedSearchComplete(LPDISPATCH SearchObject)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfa6a, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, SearchObject);
	}
	void AdvancedSearchStopped(LPDISPATCH SearchObject)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfa6b, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, SearchObject);
	}
	void MAPILogonComplete()
	{
		InvokeHelper(0xfa90, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// ApplicationEvents_10 属性
public:

};
