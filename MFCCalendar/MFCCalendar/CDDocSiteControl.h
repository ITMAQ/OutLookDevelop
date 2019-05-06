// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CDDocSiteControl 包装器类

class CDDocSiteControl : public COleDispatchDriver
{
public:
	CDDocSiteControl() {} // 调用 COleDispatchDriver 默认构造函数
	CDDocSiteControl(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDDocSiteControl(const CDDocSiteControl& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _DDocSiteControl 方法
public:
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
	signed char get_SuppressAttachments()
	{
		signed char result;
		InvokeHelper(0xfbe3, DISPATCH_PROPERTYGET, VT_I1, (void*)&result, nullptr);
		return result;
	}
	void put_SuppressAttachments(signed char newValue)
	{
		static BYTE parms[] = VTS_I1;
		InvokeHelper(0xfbe3, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}

	// _DDocSiteControl 属性
public:

};
