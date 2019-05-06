// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CConversation 包装器类

class CConversation : public COleDispatchDriver
{
public:
	CConversation() {} // 调用 COleDispatchDriver 默认构造函数
	CConversation(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CConversation(const CConversation& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _Conversation 方法
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
	LPDISPATCH GetTable()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc4f, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetChildren(LPDISPATCH Item)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc50, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Item);
		return result;
	}
	LPDISPATCH GetParent(LPDISPATCH Item)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc52, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Item);
		return result;
	}
	LPDISPATCH GetRootItems()
	{
		LPDISPATCH result;
		InvokeHelper(0xfc53, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	CString GetAlwaysAssignCategories(LPDISPATCH Store)
	{
		CString result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc5a, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Store);
		return result;
	}
	long GetAlwaysDelete(LPDISPATCH Store)
	{
		long result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc5b, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Store);
		return result;
	}
	LPDISPATCH GetAlwaysMoveToFolder(LPDISPATCH Store)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc5c, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Store);
		return result;
	}
	void MarkAsRead()
	{
		InvokeHelper(0xfc5d, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void MarkAsUnread()
	{
		InvokeHelper(0xfc5e, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void SetAlwaysAssignCategories(LPCTSTR Categories, LPDISPATCH Store)
	{
		static BYTE parms[] = VTS_BSTR VTS_DISPATCH;
		InvokeHelper(0xfc5f, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Categories, Store);
	}
	void SetAlwaysDelete(long AlwaysDelete, LPDISPATCH Store)
	{
		static BYTE parms[] = VTS_I4 VTS_DISPATCH;
		InvokeHelper(0xfc60, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, AlwaysDelete, Store);
	}
	void SetAlwaysMoveToFolder(LPDISPATCH MoveToFolder, LPDISPATCH Store)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH;
		InvokeHelper(0xfc61, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, MoveToFolder, Store);
	}
	void ClearAlwaysAssignCategories(LPDISPATCH Store)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc62, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Store);
	}
	void StopAlwaysDelete(LPDISPATCH Store)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc63, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Store);
	}
	void StopAlwaysMoveToFolder(LPDISPATCH Store)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc64, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Store);
	}
	CString get_ConversationID()
	{
		CString result;
		InvokeHelper(0xfc75, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}

	// _Conversation 属性
public:

};
