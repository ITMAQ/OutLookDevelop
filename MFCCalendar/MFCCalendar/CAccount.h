// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CAccount ��װ����

class CAccount : public COleDispatchDriver
{
public:
	CAccount() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CAccount(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAccount(const CAccount& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _Account ����
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
	long get_AccountType()
	{
		long result;
		InvokeHelper(0xfad2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	CString get_DisplayName()
	{
		CString result;
		InvokeHelper(0x3001, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	CString get_UserName()
	{
		CString result;
		InvokeHelper(0xfad3, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	CString get_SmtpAddress()
	{
		CString result;
		InvokeHelper(0xfad4, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	long get_AutoDiscoverConnectionMode()
	{
		long result;
		InvokeHelper(0xfc6f, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_CurrentUser()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc6e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_AutoDiscoverXml()
	{
		CString result;
		InvokeHelper(0xfc70, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_DeliveryStore()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc66, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_ExchangeConnectionMode()
	{
		long result;
		InvokeHelper(0xfc67, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	CString get_ExchangeMailboxServerName()
	{
		CString result;
		InvokeHelper(0xfc68, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	CString get_ExchangeMailboxServerVersion()
	{
		CString result;
		InvokeHelper(0xfc69, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetAddressEntryFromID(LPCTSTR ID)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xfc6a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, ID);
		return result;
	}
	LPDISPATCH GetRecipientFromID(LPCTSTR EntryID)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xfc6b, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, EntryID);
		return result;
	}
	LPUNKNOWN get_IOlkAccount()
	{
		LPUNKNOWN result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, nullptr);
		return result;
	}

	// _Account ����
public:

};
