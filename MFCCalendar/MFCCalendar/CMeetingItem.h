// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CMeetingItem 包装器类

class CMeetingItem : public COleDispatchDriver
{
public:
	CMeetingItem() {} // 调用 COleDispatchDriver 默认构造函数
	CMeetingItem(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CMeetingItem(const CMeetingItem& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _MeetingItem 方法
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
	LPDISPATCH get_Actions()
	{
		LPDISPATCH result;
		InvokeHelper(0xf817, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Attachments()
	{
		LPDISPATCH result;
		InvokeHelper(0xf815, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_BillingInformation()
	{
		CString result;
		InvokeHelper(0x8535, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_BillingInformation(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x8535, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_Body()
	{
		CString result;
		InvokeHelper(0x9100, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Body(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x9100, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_Categories()
	{
		CString result;
		InvokeHelper(0x9001, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Categories(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x9001, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_Companies()
	{
		CString result;
		InvokeHelper(0x853b, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Companies(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x853b, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_ConversationIndex()
	{
		CString result;
		InvokeHelper(0xfac0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	CString get_ConversationTopic()
	{
		CString result;
		InvokeHelper(0x70, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	DATE get_CreationTime()
	{
		DATE result;
		InvokeHelper(0x3007, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	CString get_EntryID()
	{
		CString result;
		InvokeHelper(0xf01e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_FormDescription()
	{
		LPDISPATCH result;
		InvokeHelper(0xf095, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_GetInspector()
	{
		LPDISPATCH result;
		InvokeHelper(0xf03e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_Importance()
	{
		long result;
		InvokeHelper(0x17, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_Importance(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x17, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_LastModificationTime()
	{
		DATE result;
		InvokeHelper(0x3008, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	LPUNKNOWN get_MAPIOBJECT()
	{
		LPUNKNOWN result;
		InvokeHelper(0xf100, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, nullptr);
		return result;
	}
	CString get_MessageClass()
	{
		CString result;
		InvokeHelper(0x1a, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_MessageClass(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x1a, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_Mileage()
	{
		CString result;
		InvokeHelper(0x8534, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Mileage(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x8534, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_NoAging()
	{
		BOOL result;
		InvokeHelper(0x850e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_NoAging(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x850e, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_OutlookInternalVersion()
	{
		long result;
		InvokeHelper(0x8552, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	CString get_OutlookVersion()
	{
		CString result;
		InvokeHelper(0x8554, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	BOOL get_Saved()
	{
		BOOL result;
		InvokeHelper(0xf0a3, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	long get_Sensitivity()
	{
		long result;
		InvokeHelper(0x36, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_Sensitivity(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x36, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_Size()
	{
		long result;
		InvokeHelper(0xe08, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	CString get_Subject()
	{
		CString result;
		InvokeHelper(0x37, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Subject(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x37, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_UnRead()
	{
		BOOL result;
		InvokeHelper(0xf01c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_UnRead(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xf01c, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_UserProperties()
	{
		LPDISPATCH result;
		InvokeHelper(0xf816, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void Close(long SaveMode)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xf023, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, SaveMode);
	}
	LPDISPATCH Copy()
	{
		LPDISPATCH result;
		InvokeHelper(0xf032, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void Delete()
	{
		InvokeHelper(0xf04a, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Display(VARIANT& Modal)
	{
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0xf0a6, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &Modal);
	}
	LPDISPATCH Move(LPDISPATCH DestFldr)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf034, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, DestFldr);
		return result;
	}
	void PrintOut()
	{
		InvokeHelper(0xf033, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Save()
	{
		InvokeHelper(0xf048, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void SaveAs(LPCTSTR Path, VARIANT& Type)
	{
		static BYTE parms[] = VTS_BSTR VTS_VARIANT;
		InvokeHelper(0xf051, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Path, &Type);
	}
	BOOL get_AutoForwarded()
	{
		BOOL result;
		InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_AutoForwarded(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x5, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_DeferredDeliveryTime()
	{
		DATE result;
		InvokeHelper(0xf, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_DeferredDeliveryTime(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0xf, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_DeleteAfterSubmit()
	{
		BOOL result;
		InvokeHelper(0xe01, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_DeleteAfterSubmit(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xe01, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_ExpiryTime()
	{
		DATE result;
		InvokeHelper(0x15, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_ExpiryTime(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0x15, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_FlagDueBy()
	{
		DATE result;
		InvokeHelper(0x30, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_FlagDueBy(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0x30, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_FlagRequest()
	{
		CString result;
		InvokeHelper(0x8530, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_FlagRequest(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x8530, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_FlagStatus()
	{
		long result;
		InvokeHelper(0x1090, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_FlagStatus(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x1090, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_OriginatorDeliveryReportRequested()
	{
		BOOL result;
		InvokeHelper(0x23, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_OriginatorDeliveryReportRequested(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x23, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_ReceivedTime()
	{
		DATE result;
		InvokeHelper(0xe06, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_ReceivedTime(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0xe06, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_Recipients()
	{
		LPDISPATCH result;
		InvokeHelper(0xf814, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	BOOL get_ReminderSet()
	{
		BOOL result;
		InvokeHelper(0x8503, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_ReminderSet(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x8503, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_ReminderTime()
	{
		DATE result;
		InvokeHelper(0x8502, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_ReminderTime(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0x8502, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_ReplyRecipients()
	{
		LPDISPATCH result;
		InvokeHelper(0xf013, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_SaveSentMessageFolder()
	{
		LPDISPATCH result;
		InvokeHelper(0xf401, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void putref_SaveSentMessageFolder(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf401, DISPATCH_PROPERTYPUTREF, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_SenderName()
	{
		CString result;
		InvokeHelper(0xc1a, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	BOOL get_Sent()
	{
		BOOL result;
		InvokeHelper(0xf402, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	DATE get_SentOn()
	{
		DATE result;
		InvokeHelper(0x39, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	BOOL get_Submitted()
	{
		BOOL result;
		InvokeHelper(0xf403, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH Forward()
	{
		LPDISPATCH result;
		InvokeHelper(0xf813, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetAssociatedAppointment(BOOL AddToCalendar)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xf760, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, AddToCalendar);
		return result;
	}
	LPDISPATCH Reply()
	{
		LPDISPATCH result;
		InvokeHelper(0xf810, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH ReplyAll()
	{
		LPDISPATCH result;
		InvokeHelper(0xf811, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void Send()
	{
		InvokeHelper(0xf075, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	LPDISPATCH get_Links()
	{
		LPDISPATCH result;
		InvokeHelper(0xf405, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_DownloadState()
	{
		long result;
		InvokeHelper(0xfa4d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void ShowCategoriesDialog()
	{
		InvokeHelper(0xfa0b, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	LPDISPATCH get_ItemProperties()
	{
		LPDISPATCH result;
		InvokeHelper(0xfa09, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_MarkForDownload()
	{
		long result;
		InvokeHelper(0x8571, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_MarkForDownload(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x8571, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_IsConflict()
	{
		BOOL result;
		InvokeHelper(0xfaa4, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	CString get_MeetingWorkspaceURL()
	{
		CString result;
		InvokeHelper(0x8209, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	long get_FlagIcon()
	{
		long result;
		InvokeHelper(0xfacc, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_FlagIcon(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xfacc, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_AutoResolvedWinner()
	{
		BOOL result;
		InvokeHelper(0xfaba, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Conflicts()
	{
		LPDISPATCH result;
		InvokeHelper(0xfabb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_SenderEmailAddress()
	{
		CString result;
		InvokeHelper(0xc1f, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	CString get_SenderEmailType()
	{
		CString result;
		InvokeHelper(0xc1e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_PropertyAccessor()
	{
		LPDISPATCH result;
		InvokeHelper(0xfafd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetConversation()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc54, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString get_ConversationID()
	{
		CString result;
		InvokeHelper(0xfc75, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_SendUsingAccount()
	{
		LPDISPATCH result;
		InvokeHelper(0xfad1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void put_SendUsingAccount(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfad1, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_IsLatestVersion()
	{
		BOOL result;
		InvokeHelper(0xfc82, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	VARIANT get_RTFBody()
	{
		VARIANT result;
		InvokeHelper(0xfc84, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, nullptr);
		return result;
	}
	void put_RTFBody(VARIANT& newValue)
	{
		static BYTE parms[] = VTS_VARIANT;
		InvokeHelper(0xfc84, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, &newValue);
	}
	DATE get_RetentionExpirationDate()
	{
		DATE result;
		InvokeHelper(0xfc8a, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	CString get_RetentionPolicyName()
	{
		CString result;
		InvokeHelper(0xfc8b, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}

	// _MeetingItem 属性
public:

};
