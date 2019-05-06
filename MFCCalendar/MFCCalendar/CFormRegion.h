// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CFormRegion 包装器类

class CFormRegion : public COleDispatchDriver
{
public:
	CFormRegion() {} // 调用 COleDispatchDriver 默认构造函数
	CFormRegion(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFormRegion(const CFormRegion& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _FormRegion 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0xf000, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_Class()
	{
		long result;
		InvokeHelper(0xf00a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Session()
	{
		LPDISPATCH result;
		InvokeHelper(0xf00b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0xf001, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	BOOL get_IsExpanded()
	{
		BOOL result;
		InvokeHelper(0xfb38, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Form()
	{
		LPDISPATCH result;
		InvokeHelper(0xfb39, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_Detail()
	{
		CString result;
		InvokeHelper(0xfb35, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Detail(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xfb35, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_Item()
	{
		LPDISPATCH result;
		InvokeHelper(0xfa60, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Inspector()
	{
		LPDISPATCH result;
		InvokeHelper(0xf03e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_FormRegionMode()
	{
		long result;
		InvokeHelper(0xfb34, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	long get_Language()
	{
		long result;
		InvokeHelper(0xfa41, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void Reflow()
	{
		InvokeHelper(0xfb3b, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	CString get_DisplayName()
	{
		CString result;
		InvokeHelper(0x3001, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	CString get_InternalName()
	{
		CString result;
		InvokeHelper(0x2102, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void SetControlItemProperty(LPDISPATCH Control, LPCTSTR PropertyName)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR;
		InvokeHelper(0xfac9, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Control, PropertyName);
	}
	BOOL get_SuppressControlReplacement()
	{
		BOOL result;
		InvokeHelper(0xfc06, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_SuppressControlReplacement(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfc06, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	void Select()
	{
		InvokeHelper(0x2109, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	BOOL get_EnableAutoLayout()
	{
		BOOL result;
		InvokeHelper(0xfc1a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_EnableAutoLayout(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfc1a, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_Visible()
	{
		BOOL result;
		InvokeHelper(0x2103, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_Visible(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x2103, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}

	// _FormRegion 属性
public:

};
