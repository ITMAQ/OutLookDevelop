// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CFolders ��װ����

class CFolders : public COleDispatchDriver
{
public:
	CFolders() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CFolders(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFolders(const CFolders& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _Folders ����
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
	LPDISPATCH Add(LPCTSTR Name, VARIANT& Type)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT;
		InvokeHelper(0x5f, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, &Type);
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

	// _Folders ����
public:

};
