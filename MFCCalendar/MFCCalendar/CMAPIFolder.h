// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CMAPIFolder 包装器类

class CMAPIFolder : public COleDispatchDriver
{
public:
	CMAPIFolder() {} // 调用 COleDispatchDriver 默认构造函数
	CMAPIFolder(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CMAPIFolder(const CMAPIFolder& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// MAPIFolder 方法
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
	long get_DefaultItemType()
	{
		long result;
		InvokeHelper(0x3106, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	CString get_DefaultMessageClass()
	{
		CString result;
		InvokeHelper(0x3107, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	CString get_Description()
	{
		CString result;
		InvokeHelper(0x3004, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Description(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x3004, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_EntryID()
	{
		CString result;
		InvokeHelper(0xf01e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Folders()
	{
		LPDISPATCH result;
		InvokeHelper(0x2103, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Items()
	{
		LPDISPATCH result;
		InvokeHelper(0x3100, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x3001, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Name(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x3001, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_StoreID()
	{
		CString result;
		InvokeHelper(0x3108, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	long get_UnReadItemCount()
	{
		long result;
		InvokeHelper(0x3603, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH CopyTo(LPDISPATCH DestinationFolder)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf032, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, DestinationFolder);
		return result;
	}
	void Delete()
	{
		InvokeHelper(0xf045, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Display()
	{
		InvokeHelper(0x3104, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	LPDISPATCH GetExplorer(VARIANT& DisplayMode)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x3101, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &DisplayMode);
		return result;
	}
	void MoveTo(LPDISPATCH DestinationFolder)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf034, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, DestinationFolder);
	}
	LPDISPATCH get_UserPermissions()
	{
		LPDISPATCH result;
		InvokeHelper(0x3111, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	BOOL get_WebViewOn()
	{
		BOOL result;
		InvokeHelper(0x3112, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_WebViewOn(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x3112, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_WebViewURL()
	{
		CString result;
		InvokeHelper(0x3113, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_WebViewURL(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x3113, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_WebViewAllowNavigation()
	{
		BOOL result;
		InvokeHelper(0x3114, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_WebViewAllowNavigation(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x3114, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	void AddToPFFavorites()
	{
		InvokeHelper(0x3115, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	CString get_AddressBookName()
	{
		CString result;
		InvokeHelper(0xfa6e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_AddressBookName(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xfa6e, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_ShowAsOutlookAB()
	{
		BOOL result;
		InvokeHelper(0xfa6f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_ShowAsOutlookAB(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfa6f, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_FolderPath()
	{
		CString result;
		InvokeHelper(0xfa78, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void AddToFavorites(VARIANT& fNoUI, VARIANT& Name)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0xfa61, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &fNoUI, &Name);
	}
	BOOL get_InAppFolderSyncObject()
	{
		BOOL result;
		InvokeHelper(0xfa4b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_InAppFolderSyncObject(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfa4b, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_CurrentView()
	{
		LPDISPATCH result;
		InvokeHelper(0x2200, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	BOOL get_CustomViewsOnly()
	{
		BOOL result;
		InvokeHelper(0xfa46, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_CustomViewsOnly(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfa46, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_Views()
	{
		LPDISPATCH result;
		InvokeHelper(0x3109, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPUNKNOWN get_MAPIOBJECT()
	{
		LPUNKNOWN result;
		InvokeHelper(0xf100, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, nullptr);
		return result;
	}
	CString get_FullFolderPath()
	{
		CString result;
		InvokeHelper(0xfa91, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	BOOL get_IsSharePointFolder()
	{
		BOOL result;
		InvokeHelper(0xfab6, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	long get_ShowItemCount()
	{
		long result;
		InvokeHelper(0xfac2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_ShowItemCount(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xfac2, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_Store()
	{
		LPDISPATCH result;
		InvokeHelper(0xfad9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetStorage(LPCTSTR StorageIdentifier, long StorageIdentifierType)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_I4;
		InvokeHelper(0xfb08, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, StorageIdentifier, StorageIdentifierType);
		return result;
	}
	LPDISPATCH GetTable(VARIANT& Filter, VARIANT& TableContents)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0xfb1d, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Filter, &TableContents);
		return result;
	}
	LPDISPATCH get_PropertyAccessor()
	{
		LPDISPATCH result;
		InvokeHelper(0xfafd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetCalendarExporter()
	{
		LPDISPATCH result;
		InvokeHelper(0xfba2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_UserDefinedProperties()
	{
		LPDISPATCH result;
		InvokeHelper(0xf816, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetCustomIcon()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc3c, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void SetCustomIcon(LPDISPATCH Picture)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc3d, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Picture);
	}

	// MAPIFolder 属性
public:

};
