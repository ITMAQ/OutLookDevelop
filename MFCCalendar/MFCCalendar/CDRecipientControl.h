// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CDRecipientControl 包装器类

class CDRecipientControl : public COleDispatchDriver
{
public:
	CDRecipientControl() {} // 调用 COleDispatchDriver 默认构造函数
	CDRecipientControl(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDRecipientControl(const CDRecipientControl& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _DRecipientControl 方法
public:
	signed char get_Enabled()
	{
		signed char result;
		InvokeHelper(DISPID_ENABLED, DISPATCH_PROPERTYGET, VT_I1, (void*)&result, nullptr);
		return result;
	}
	void put_Enabled(signed char newValue)
	{
		static BYTE parms[] = VTS_I1;
		InvokeHelper(DISPID_ENABLED, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_BackColor()
	{
		long result;
		InvokeHelper(DISPID_BACKCOLOR, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_BackColor(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(DISPID_BACKCOLOR, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_ForeColor()
	{
		long result;
		InvokeHelper(DISPID_FORECOLOR, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_ForeColor(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(DISPID_FORECOLOR, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	signed char get_ReadOnly()
	{
		signed char result;
		InvokeHelper(0x8001f008, DISPATCH_PROPERTYGET, VT_I1, (void*)&result, nullptr);
		return result;
	}
	void put_ReadOnly(signed char newValue)
	{
		static BYTE parms[] = VTS_I1;
		InvokeHelper(0x8001f008, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_Font()
	{
		LPDISPATCH result;
		InvokeHelper(DISPID_FONT, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void put_Font(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(DISPID_FONT, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_SpecialEffect()
	{
		long result;
		InvokeHelper(0xc, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_SpecialEffect(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xc, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}

	// _DRecipientControl 属性
public:

};
