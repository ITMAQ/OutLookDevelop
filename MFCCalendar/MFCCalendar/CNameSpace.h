// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CNameSpace 包装器类

class CNameSpace : public COleDispatchDriver
{
public:
	CNameSpace() {} // 调用 COleDispatchDriver 默认构造函数
	CNameSpace(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CNameSpace(const CNameSpace& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _NameSpace 方法
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
	LPDISPATCH get_CurrentUser()
	{
		LPDISPATCH result;
		InvokeHelper(0x2101, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Folders()
	{
		LPDISPATCH result;
		InvokeHelper(0x2103, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_Type()
	{
		CString result;
		InvokeHelper(0x2104, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_AddressLists()
	{
		LPDISPATCH result;
		InvokeHelper(0x210d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH CreateRecipient(LPCTSTR RecipientName)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x210a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, RecipientName);
		return result;
	}
	LPDISPATCH GetDefaultFolder(long FolderType)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x210b, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FolderType);
		return result;
	}
	LPDISPATCH GetFolderFromID(LPCTSTR EntryIDFolder, VARIANT& EntryIDStore)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT;
		InvokeHelper(0x2108, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, EntryIDFolder, &EntryIDStore);
		return result;
	}
	LPDISPATCH GetItemFromID(LPCTSTR EntryIDItem, VARIANT& EntryIDStore)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT;
		InvokeHelper(0x2109, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, EntryIDItem, &EntryIDStore);
		return result;
	}
	LPDISPATCH GetRecipientFromID(LPCTSTR EntryID)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x2107, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, EntryID);
		return result;
	}
	LPDISPATCH GetSharedDefaultFolder(LPDISPATCH Recipient, long FolderType)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_I4;
		InvokeHelper(0x210c, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Recipient, FolderType);
		return result;
	}
	void Logoff()
	{
		InvokeHelper(0x2106, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Logon(VARIANT& Profile, VARIANT& Password, VARIANT& ShowDialog, VARIANT& NewSession)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0x2105, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &Profile, &Password, &ShowDialog, &NewSession);
	}
	LPDISPATCH PickFolder()
	{
		LPDISPATCH result;
		InvokeHelper(0x210e, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void RefreshRemoteHeaders()
	{
		InvokeHelper(0x2117, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	LPDISPATCH get_SyncObjects()
	{
		LPDISPATCH result;
		InvokeHelper(0x2118, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void AddStore(VARIANT& Store)
	{
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0x2119, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &Store);
	}
	void RemoveStore(LPDISPATCH Folder)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0x211a, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Folder);
	}
	BOOL get_Offline()
	{
		BOOL result;
		InvokeHelper(0xfa4c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void Dial(VARIANT& ContactItem)
	{
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0xfa0d, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &ContactItem);
	}
	LPUNKNOWN get_MAPIOBJECT()
	{
		LPUNKNOWN result;
		InvokeHelper(0xf100, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, nullptr);
		return result;
	}
	long get_ExchangeConnectionMode()
	{
		long result;
		InvokeHelper(0xfac1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void AddStoreEx(VARIANT& Store, long Type)
	{
		static BYTE parms[] = VTS_VARIANT VTS_I4;
		InvokeHelper(0xfac5, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &Store, Type);
	}
	LPDISPATCH get_Accounts()
	{
		LPDISPATCH result;
		InvokeHelper(0xfad0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_CurrentProfileName()
	{
		CString result;
		InvokeHelper(0xfad5, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Stores()
	{
		LPDISPATCH result;
		InvokeHelper(0xfad8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetSelectNamesDialog()
	{
		LPDISPATCH result;
		InvokeHelper(0xfae1, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void SendAndReceive(BOOL showProgressDialog)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfad7, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, showProgressDialog);
	}
	LPDISPATCH get_DefaultStore()
	{
		LPDISPATCH result;
		InvokeHelper(0xfaec, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetAddressEntryFromID(LPCTSTR ID)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xfb04, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, ID);
		return result;
	}
	LPDISPATCH GetGlobalAddressList()
	{
		LPDISPATCH result;
		InvokeHelper(0xfb05, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetStoreFromID(LPCTSTR ID)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xfb06, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, ID);
		return result;
	}
	LPDISPATCH get_Categories()
	{
		LPDISPATCH result;
		InvokeHelper(0xfba5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH OpenSharedFolder(LPCTSTR Path, VARIANT& Name, VARIANT& DownloadAttachments, VARIANT& UseTTL)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0xfbf6, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Path, &Name, &DownloadAttachments, &UseTTL);
		return result;
	}
	LPDISPATCH OpenSharedItem(LPCTSTR Path)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xfbf7, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Path);
		return result;
	}
	LPDISPATCH CreateSharingItem(VARIANT& Context, VARIANT& Provider)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0xfbe4, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Context, &Provider);
		return result;
	}
	CString get_ExchangeMailboxServerName()
	{
		CString result;
		InvokeHelper(0xfc05, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	CString get_ExchangeMailboxServerVersion()
	{
		CString result;
		InvokeHelper(0xfc04, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	BOOL CompareEntryIDs(LPCTSTR FirstEntryID, LPCTSTR SecondEntryID)
	{
		BOOL result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR;
		InvokeHelper(0xfbfc, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, FirstEntryID, SecondEntryID);
		return result;
	}
	CString get_AutoDiscoverXml()
	{
		CString result;
		InvokeHelper(0xfc03, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	long get_AutoDiscoverConnectionMode()
	{
		long result;
		InvokeHelper(0xfc2e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH CreateContactCard(LPDISPATCH AddressEntry)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc85, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, AddressEntry);
		return result;
	}

	// _NameSpace 属性
public:

};
