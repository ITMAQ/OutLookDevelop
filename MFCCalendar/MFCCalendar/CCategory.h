// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CCategory ��װ����

class CCategory : public COleDispatchDriver
{
public:
	CCategory() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CCategory(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCategory(const CCategory& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _Category ����
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
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x2102, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Name(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x2102, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_Color()
	{
		long result;
		InvokeHelper(0xfba3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_Color(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xfba3, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_ShortcutKey()
	{
		long result;
		InvokeHelper(0xfba4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_ShortcutKey(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xfba4, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_CategoryID()
	{
		CString result;
		InvokeHelper(0xfbd0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	unsigned long get_CategoryBorderColor()
	{
		unsigned long result;
		InvokeHelper(0xfc1b, DISPATCH_PROPERTYGET, VT_UI4, (void*)&result, nullptr);
		return result;
	}
	unsigned long get_CategoryGradientTopColor()
	{
		unsigned long result;
		InvokeHelper(0xfc1c, DISPATCH_PROPERTYGET, VT_UI4, (void*)&result, nullptr);
		return result;
	}
	unsigned long get_CategoryGradientBottomColor()
	{
		unsigned long result;
		InvokeHelper(0xfc1d, DISPATCH_PROPERTYGET, VT_UI4, (void*)&result, nullptr);
		return result;
	}

	// _Category ����
public:

};
