// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CNavigationFolders ��װ����

class CNavigationFolders : public COleDispatchDriver
{
public:
	CNavigationFolders() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CNavigationFolders(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CNavigationFolders(const CNavigationFolders& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _NavigationFolders ����
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
	LPDISPATCH Add(LPDISPATCH Folder)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfbc2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Folder);
		return result;
	}
	void Remove(LPDISPATCH RemovableFolder)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfbc3, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, RemovableFolder);
	}

	// _NavigationFolders ����
public:

};
