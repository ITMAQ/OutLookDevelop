// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CActions ��װ����

class CActions : public COleDispatchDriver
{
public:
	CActions() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CActions(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CActions(const CActions& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// Actions ����
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
	LPDISPATCH Add()
	{
		LPDISPATCH result;
		InvokeHelper(0x64, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void Remove(long Index)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x52, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Index);
	}

	// Actions ����
public:

};
