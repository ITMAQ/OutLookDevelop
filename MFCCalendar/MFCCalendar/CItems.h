// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CItems 包装器类

class CItems : public COleDispatchDriver
{
public:
	CItems() {} // 调用 COleDispatchDriver 默认构造函数
	CItems(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CItems(const CItems& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _Items 方法
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
	long get_Count()
	{
		long result;
		InvokeHelper(0x50, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH Item(VARIANT& Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x51, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPUNKNOWN get_RawTable()
	{
		LPUNKNOWN result;
		InvokeHelper(0x5a, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, nullptr);
		return result;
	}
	BOOL get_IncludeRecurrences()
	{
		BOOL result;
		InvokeHelper(0xce, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_IncludeRecurrences(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xce, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH Add(VARIANT& Type)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x5f, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Type);
		return result;
	}
	LPDISPATCH Find(LPCTSTR Filter)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x62, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Filter);
		return result;
	}
	LPDISPATCH FindNext()
	{
		LPDISPATCH result;
		InvokeHelper(0x63, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetFirst()
	{
		LPDISPATCH result;
		InvokeHelper(0x56, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetLast()
	{
		LPDISPATCH result;
		InvokeHelper(0x58, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetNext()
	{
		LPDISPATCH result;
		InvokeHelper(0x57, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetPrevious()
	{
		LPDISPATCH result;
		InvokeHelper(0x59, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void Remove(long Index)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x54, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Index);
	}
	void ResetColumns()
	{
		InvokeHelper(0x5d, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	LPDISPATCH Restrict(LPCTSTR Filter)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x64, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Filter);
		return result;
	}
	void SetColumns(LPCTSTR Columns)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x5c, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Columns);
	}
	void Sort(LPCTSTR Property, VARIANT& Descending)
	{
		static BYTE parms[] = VTS_BSTR VTS_VARIANT;
		InvokeHelper(0x61, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Property, &Descending);
	}

	// _Items 属性
public:

};
