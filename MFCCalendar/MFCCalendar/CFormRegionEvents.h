// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CFormRegionEvents 包装器类

class CFormRegionEvents : public COleDispatchDriver
{
public:
	CFormRegionEvents() {} // 调用 COleDispatchDriver 默认构造函数
	CFormRegionEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFormRegionEvents(const CFormRegionEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// FormRegionEvents 方法
public:
	void Expanded(BOOL Expand)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfb38, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Expand);
	}
	void Close()
	{
		InvokeHelper(0xf004, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// FormRegionEvents 属性
public:

};
