// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CApplicationEvents ��װ����

class CApplicationEvents : public COleDispatchDriver
{
public:
	CApplicationEvents() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CApplicationEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplicationEvents(const CApplicationEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// ApplicationEvents ����
public:
	void ItemSend(LPDISPATCH Item, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xf002, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Item, Cancel);
	}
	void NewMail()
	{
		InvokeHelper(0xf003, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Reminder(LPDISPATCH Item)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf004, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Item);
	}
	void OptionsPagesAdd(LPDISPATCH Pages)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf005, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Pages);
	}
	void Startup()
	{
		InvokeHelper(0xf006, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Quit()
	{
		InvokeHelper(0xf007, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// ApplicationEvents ����
public:

};
