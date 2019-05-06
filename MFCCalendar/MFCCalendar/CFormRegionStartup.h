// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CFormRegionStartup 包装器类

class CFormRegionStartup : public COleDispatchDriver
{
public:
	CFormRegionStartup() {} // 调用 COleDispatchDriver 默认构造函数
	CFormRegionStartup(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFormRegionStartup(const CFormRegionStartup& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _FormRegionStartup 方法
public:
	VARIANT GetFormRegionStorage(LPCTSTR FormRegionName, LPDISPATCH Item, long LCID, long FormRegionMode, long FormRegionSize)
	{
		VARIANT result;
		static BYTE parms[] = VTS_BSTR VTS_DISPATCH VTS_I4 VTS_I4 VTS_I4;
		InvokeHelper(0xfb36, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, FormRegionName, Item, LCID, FormRegionMode, FormRegionSize);
		return result;
	}
	void BeforeFormRegionShow(LPDISPATCH FormRegion)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfb3d, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, FormRegion);
	}
	VARIANT GetFormRegionManifest(LPCTSTR FormRegionName, long LCID)
	{
		VARIANT result;
		static BYTE parms[] = VTS_BSTR VTS_I4;
		InvokeHelper(0xfc33, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, FormRegionName, LCID);
		return result;
	}
	VARIANT GetFormRegionIcon(LPCTSTR FormRegionName, long LCID, long Icon)
	{
		VARIANT result;
		static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4;
		InvokeHelper(0xfc34, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, FormRegionName, LCID, Icon);
		return result;
	}

	// _FormRegionStartup 属性
public:

};
