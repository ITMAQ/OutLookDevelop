// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CAddressEntries ��װ����

class CAddressEntries : public COleDispatchDriver
{
public:
	CAddressEntries() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CAddressEntries(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAddressEntries(const CAddressEntries& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// AddressEntries ����
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
	LPDISPATCH Add(LPCTSTR Type, VARIANT& Name, VARIANT& Address)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0x5f, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Type, &Name, &Address);
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
	void Sort(VARIANT& Property, VARIANT& Order)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0x61, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &Property, &Order);
	}

	// AddressEntries ����
public:

};
