// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CAddressList ��װ����

class CAddressList : public COleDispatchDriver
{
public:
	CAddressList() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CAddressList(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAddressList(const CAddressList& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// AddressList ����
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
	LPDISPATCH get_AddressEntries()
	{
		LPDISPATCH result;
		InvokeHelper(0x100, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_ID()
	{
		CString result;
		InvokeHelper(0xf01e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	long get_Index()
	{
		long result;
		InvokeHelper(0x5b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	BOOL get_IsReadOnly()
	{
		BOOL result;
		InvokeHelper(0xf017, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x3001, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetContactsFolder()
	{
		LPDISPATCH result;
		InvokeHelper(0xfafa, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_AddressListType()
	{
		long result;
		InvokeHelper(0xfaed, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	long get_ResolutionOrder()
	{
		long result;
		InvokeHelper(0xfaee, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	BOOL get_IsInitialAddressList()
	{
		BOOL result;
		InvokeHelper(0xfaf8, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_PropertyAccessor()
	{
		LPDISPATCH result;
		InvokeHelper(0xfafd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}

	// AddressList ����
public:

};
