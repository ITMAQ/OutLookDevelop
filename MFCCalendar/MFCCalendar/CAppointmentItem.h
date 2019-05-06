// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CAppointmentItem 包装器类

class CAppointmentItem : public COleDispatchDriver
{
public:
	CAppointmentItem() {} // 调用 COleDispatchDriver 默认构造函数
	CAppointmentItem(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAppointmentItem(const CAppointmentItem& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _AppointmentItem 方法
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
	BOOL get_AllDayEvent()
	{
		BOOL result;
		InvokeHelper(0x8215, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_AllDayEvent(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x8215, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_BusyStatus()
	{
		long result;
		InvokeHelper(0x8205, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_BusyStatus(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x8205, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_Duration()
	{
		long result;
		InvokeHelper(0x8213, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_Duration(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x8213, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_End()
	{
		DATE result;
		InvokeHelper(0x820e, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_End(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0x820e, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_IsOnlineMeeting()
	{
		BOOL result;
		InvokeHelper(0x8240, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_IsOnlineMeeting(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x8240, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_IsRecurring()
	{
		BOOL result;
		InvokeHelper(0x8223, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	CString get_Location()
	{
		CString result;
		InvokeHelper(0x8208, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Location(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x8208, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_MeetingStatus()
	{
		long result;
		InvokeHelper(0x8217, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_MeetingStatus(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x8217, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_NetMeetingAutoStart()
	{
		BOOL result;
		InvokeHelper(0x8244, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_NetMeetingAutoStart(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x8244, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_NetMeetingOrganizerAlias()
	{
		CString result;
		InvokeHelper(0x8243, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_NetMeetingOrganizerAlias(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x8243, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_NetMeetingServer()
	{
		CString result;
		InvokeHelper(0x8242, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_NetMeetingServer(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x8242, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_NetMeetingType()
	{
		long result;
		InvokeHelper(0x8241, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_NetMeetingType(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x8241, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_OptionalAttendees()
	{
		CString result;
		InvokeHelper(0xe03, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_OptionalAttendees(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xe03, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_Organizer()
	{
		CString result;
		InvokeHelper(0x42, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH get_Recipients()
	{
		LPDISPATCH result;
		InvokeHelper(0xf814, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	long get_RecurrenceState()
	{
		long result;
		InvokeHelper(0xf545, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	long get_ReminderMinutesBeforeStart()
	{
		long result;
		InvokeHelper(0x8501, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_ReminderMinutesBeforeStart(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x8501, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_ReminderOverrideDefault()
	{
		BOOL result;
		InvokeHelper(0x851c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_ReminderOverrideDefault(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x851c, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_ReminderPlaySound()
	{
		BOOL result;
		InvokeHelper(0x851e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_ReminderPlaySound(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x851e, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
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
	CString get_ReminderSoundFile()
	{
		CString result;
		InvokeHelper(0x851f, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_ReminderSoundFile(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x851f, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_ReplyTime()
	{
		DATE result;
		InvokeHelper(0x8220, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_ReplyTime(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0x8220, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_RequiredAttendees()
	{
		CString result;
		InvokeHelper(0xe04, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_RequiredAttendees(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xe04, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_Resources()
	{
		CString result;
		InvokeHelper(0xe02, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_Resources(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0xe02, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	BOOL get_ResponseRequested()
	{
		BOOL result;
		InvokeHelper(0x63, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_ResponseRequested(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x63, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_ResponseStatus()
	{
		long result;
		InvokeHelper(0x8218, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	DATE get_Start()
	{
		DATE result;
		InvokeHelper(0x820d, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_Start(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0x820d, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	void ClearRecurrencePattern()
	{
		InvokeHelper(0xf0a5, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	LPDISPATCH ForwardAsVcal()
	{
		LPDISPATCH result;
		InvokeHelper(0xf547, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetRecurrencePattern()
	{
		LPDISPATCH result;
		InvokeHelper(0xf0a4, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH Respond(long Response, VARIANT& fNoUI, VARIANT& fAdditionalTextDialog)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT;
		InvokeHelper(0xf502, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Response, &fNoUI, &fAdditionalTextDialog);
		return result;
	}
	void Send()
	{
		InvokeHelper(0xf075, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	CString get_NetMeetingDocPathName()
	{
		CString result;
		InvokeHelper(0x8247, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_NetMeetingDocPathName(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x8247, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_NetShowURL()
	{
		CString result;
		InvokeHelper(0x8248, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_NetShowURL(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x8248, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_Links()
	{
		LPDISPATCH result;
		InvokeHelper(0xf405, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	BOOL get_ConferenceServerAllowExternal()
	{
		BOOL result;
		InvokeHelper(0x8246, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_ConferenceServerAllowExternal(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x8246, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	CString get_ConferenceServerPassword()
	{
		CString result;
		InvokeHelper(0x8249, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	void put_ConferenceServerPassword(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x8249, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_ItemProperties()
	{
		LPDISPATCH result;
		InvokeHelper(0xfa09, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
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
	long get_InternetCodepage()
	{
		long result;
		InvokeHelper(0x3fde, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_InternetCodepage(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x3fde, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
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
	LPDISPATCH get_PropertyAccessor()
	{
		LPDISPATCH result;
		InvokeHelper(0xfafd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
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
	CString get_GlobalAppointmentID()
	{
		CString result;
		InvokeHelper(0xfbfd, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	BOOL get_ForceUpdateToAllAttendees()
	{
		BOOL result;
		InvokeHelper(0xf53a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, nullptr);
		return result;
	}
	void put_ForceUpdateToAllAttendees(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xf53a, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_StartUTC()
	{
		DATE result;
		InvokeHelper(0x8268, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_StartUTC(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0x8268, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_EndUTC()
	{
		DATE result;
		InvokeHelper(0x8269, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_EndUTC(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0x8269, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_StartInStartTimeZone()
	{
		DATE result;
		InvokeHelper(0x8266, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_StartInStartTimeZone(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0x8266, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	DATE get_EndInEndTimeZone()
	{
		DATE result;
		InvokeHelper(0x8267, DISPATCH_PROPERTYGET, VT_DATE, (void*)&result, nullptr);
		return result;
	}
	void put_EndInEndTimeZone(DATE newValue)
	{
		static BYTE parms[] = VTS_DATE;
		InvokeHelper(0x8267, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_StartTimeZone()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc27, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void put_StartTimeZone(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc27, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	LPDISPATCH get_EndTimeZone()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc28, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void put_EndTimeZone(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc28, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
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
	LPDISPATCH CopyTo(LPDISPATCH DestinationFolder, long CopyOptions)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_I4;
		InvokeHelper(0xfb12, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, DestinationFolder, CopyOptions);
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
	LPDISPATCH GetOrganizer()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc88, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}

	// _AppointmentItem 属性
public:

};
