// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CColumns0 ��װ����

class CColumns0 : public COleDispatchDriver
{
public:
	CColumns0() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CColumns0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CColumns0(const CColumns0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _Columns ����
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
	LPDISPATCH Add(LPCTSTR Name)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x5f, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name);
		return result;
	}
	void Remove(VARIANT& Index)
	{
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x52, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &Index);
	}
	void RemoveAll()
	{
		InvokeHelper(0x53, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// _Columns ����
public:

};
