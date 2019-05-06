// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CExplorersEvents 包装器类

class CExplorersEvents : public COleDispatchDriver
{
public:
	CExplorersEvents() {} // 调用 COleDispatchDriver 默认构造函数
	CExplorersEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CExplorersEvents(const CExplorersEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// ExplorersEvents 方法
public:
	void NewExplorer(LPDISPATCH Explorer)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf001, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Explorer);
	}

	// ExplorersEvents 属性
public:

};
