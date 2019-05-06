// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CCalendarSharing 包装器类

class CCalendarSharing : public COleDispatchDriver
{
public:
	CCalendarSharing() {} // 调用 COleDispatchDriver 默认构造函数
	CCalendarSharing(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCalendarSharing(const CCalendarSharing& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _CalendarSharing 方法
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
	void SaveAsICal(LPCTSTR Path)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xfb98, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Path);
	}
	LPDISPATCH ForwardAsICal(long MailFormat)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xfb99, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, MailFormat);
		return result;
	}
	long get_CalendarDetail()
	{
		long result;
		InvokeHelper(0xfb9a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_CalendarDetail(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xfb9a, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_EndDate()
	{
		DATE result;
		InvokeHelper(0xfb9b, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_EndDate(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0xfb9b, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_Folder()
	{
		LPDISPATCH result;
		InvokeHelper(0xfb9c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	BOOL get_IncludeAttachments()
	{
		BOOL result;
		InvokeHelper(0xfb9d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_IncludeAttachments(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfb9d, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_IncludePrivateDetails()
	{
		BOOL result;
		InvokeHelper(0xfb9e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_IncludePrivateDetails(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfb9e, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_RestrictToWorkingHours()
	{
		BOOL result;
		InvokeHelper(0xfb9f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_RestrictToWorkingHours(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfb9f, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_StartDate()
	{
		DATE result;
		InvokeHelper(0xfba0, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_StartDate(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0xfba0, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_IncludeWholeCalendar()
	{
		BOOL result;
		InvokeHelper(0xfba1, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_IncludeWholeCalendar(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfba1, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}

	// _CalendarSharing 属性
public:

};
