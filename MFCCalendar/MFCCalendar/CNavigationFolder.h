// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CNavigationFolder ��װ����

class CNavigationFolder : public COleDispatchDriver
{
public:
	CNavigationFolder() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CNavigationFolder(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CNavigationFolder(const CNavigationFolder& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _NavigationFolder ����
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
	LPDISPATCH get_Folder()
	{
		LPDISPATCH result;
		InvokeHelper(0xfbc4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	BOOL get_IsSelected()
	{
		BOOL result;
		InvokeHelper(0xfbc5, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_IsSelected(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfbc5, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_IsSideBySide()
	{
		BOOL result;
		InvokeHelper(0xfbc6, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_IsSideBySide(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfbc6, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_Position()
	{
		long result;
		InvokeHelper(0xfbc7, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_Position(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xfbc7, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_IsRemovable()
	{
		BOOL result;
		InvokeHelper(0xfbc8, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	CString get_DisplayName()
	{
		CString result;
		InvokeHelper(0x2102, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}

	// _NavigationFolder ����
public:

};
