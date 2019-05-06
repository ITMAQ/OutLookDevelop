// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CAddressEntry 包装器类

class CAddressEntry : public COleDispatchDriver
{
public:
	CAddressEntry() {} // 调用 COleDispatchDriver 默认构造函数
	CAddressEntry(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAddressEntry(const CAddressEntry& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// AddressEntry 方法
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
	CString get_Address()
	{
		CString result;
		InvokeHelper(0x3003, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Address(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x3003, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_DisplayType()
	{
		long result;
		InvokeHelper(0x3900, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	CString get_ID()
	{
		CString result;
		InvokeHelper(0xf01e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Manager()
	{
		LPDISPATCH result;
		InvokeHelper(0x303, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPUNKNOWN get_MAPIOBJECT()
	{
		LPUNKNOWN result;
		InvokeHelper(0xf100, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, nullptr);
		return result;
	}
	void put_MAPIOBJECT(LPUNKNOWN newValue)
	{
		static BYTE parms[] = VTS_UNKNOWN;
		InvokeHelper(0xf100, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_Members()
	{
		LPDISPATCH result;
		InvokeHelper(0x304, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x3001, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Name(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x3001, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_Type()
	{
		CString result;
		InvokeHelper(0x3002, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Type(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x3002, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	void Delete()
	{
		InvokeHelper(0x302, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Details(VARIANT& HWnd)
	{
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x301, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &HWnd);
	}
	CString GetFreeBusy(DATE Start, long MinPerChar, VARIANT& CompleteFormat)
	{
		CString result;
		static BYTE parms[] = VTS_DATE VTS_I4 VTS_VARIANT;
		InvokeHelper(0x306, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Start, MinPerChar, &CompleteFormat);
		return result;
	}
	void Update(VARIANT& MakePermanent, VARIANT& Refresh)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0x300, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &MakePermanent, &Refresh);
	}
	void UpdateFreeBusy()
	{
		InvokeHelper(0x307, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	LPDISPATCH GetContact()
	{
		LPDISPATCH result;
		InvokeHelper(0xfaf0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetExchangeUser()
	{
		LPDISPATCH result;
		InvokeHelper(0xfaf1, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_AddressEntryUserType()
	{
		long result;
		InvokeHelper(0xfaf2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetExchangeDistributionList()
	{
		LPDISPATCH result;
		InvokeHelper(0xfaef, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_PropertyAccessor()
	{
		LPDISPATCH result;
		InvokeHelper(0xfafd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}

	// AddressEntry 属性
public:

};
