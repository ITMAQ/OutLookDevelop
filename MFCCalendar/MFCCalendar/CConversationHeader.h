// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CConversationHeader ��װ����

class CConversationHeader : public COleDispatchDriver
{
public:
	CConversationHeader() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CConversationHeader(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CConversationHeader(const CConversationHeader& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _ConversationHeader ����
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
	CString get_ConversationTopic()
	{
		CString result;
		InvokeHelper(0xfc86, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, nullptr);
		return result;
	}
	LPDISPATCH GetItems()
	{
		LPDISPATCH result;
		InvokeHelper(0x2109, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}

	// _ConversationHeader ����
public:

};
