// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CApplication 包装器类

class CApplication : public COleDispatchDriver
{
public:
	CApplication() {} // 调用 COleDispatchDriver 默认构造函数
	CApplication(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplication(const CApplication& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _Application 方法
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
	LPDISPATCH get_Assistant()
	{
		LPDISPATCH result;
		InvokeHelper(0x114, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x3001, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	CString get_Version()
	{
		CString result;
		InvokeHelper(0x116, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH ActiveExplorer()
	{
		LPDISPATCH result;
		InvokeHelper(0x111, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH ActiveInspector()
	{
		LPDISPATCH result;
		InvokeHelper(0x112, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH CreateItem(long ItemType)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x10a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, ItemType);
		return result;
	}
	LPDISPATCH CreateItemFromTemplate(LPCTSTR TemplatePath, VARIANT& InFolder)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT;
		InvokeHelper(0x10b, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, TemplatePath, &InFolder);
		return result;
	}
	LPDISPATCH CreateObject(LPCTSTR ObjectName)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x115, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, ObjectName);
		return result;
	}
	LPDISPATCH GetNamespace(LPCTSTR Type)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x110, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Type);
		return result;
	}
	void Quit()
	{
		InvokeHelper(0x113, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	LPDISPATCH get_COMAddIns()
	{
		LPDISPATCH result;
		InvokeHelper(0x118, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Explorers()
	{
		LPDISPATCH result;
		InvokeHelper(0x119, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Inspectors()
	{
		LPDISPATCH result;
		InvokeHelper(0x11a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_LanguageSettings()
	{
		LPDISPATCH result;
		InvokeHelper(0x11b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_ProductCode()
	{
		CString result;
		InvokeHelper(0x11c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_AnswerWizard()
	{
		LPDISPATCH result;
		InvokeHelper(0x11d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_FeatureInstall()
	{
		long result;
		InvokeHelper(0x11e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_FeatureInstall(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x11e, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH ActiveWindow()
	{
		LPDISPATCH result;
		InvokeHelper(0x11f, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH CopyFile(LPCTSTR FilePath, LPCTSTR DestFolderPath)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR;
		InvokeHelper(0xfa62, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FilePath, DestFolderPath);
		return result;
	}
	LPDISPATCH AdvancedSearch(LPCTSTR Scope, VARIANT& Filter, VARIANT& SearchSubFolders, VARIANT& Tag)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0xfa65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Scope, &Filter, &SearchSubFolders, &Tag);
		return result;
	}
	BOOL IsSearchSynchronous(LPCTSTR LookInFolders)
	{
		BOOL result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xfa6c, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, LookInFolders);
		return result;
	}
	void GetNewNickNames(VARIANT * pvar)
	{
		static BYTE parms[] = VTS_PVARIANT;
		InvokeHelper(0xfa48, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, pvar);
	}
	LPDISPATCH get_Reminders()
	{
		LPDISPATCH result;
		InvokeHelper(0xfa99, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_DefaultProfileName()
	{
		CString result;
		InvokeHelper(0xfad6, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	BOOL get_IsTrusted()
	{
		BOOL result;
		InvokeHelper(0xfbf3, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetObjectReference(LPDISPATCH Item, long ReferenceType)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_I4;
		InvokeHelper(0xfbd6, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Item, ReferenceType);
		return result;
	}
	LPDISPATCH get_Assistance()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc08, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_TimeZones()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc29, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_PickerDialog()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc65, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void RefreshFormRegionDefinition(LPCTSTR RegionName)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xfc7f, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, RegionName);
	}

	// _Application 属性
public:

};
