// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CAccountsEvents ��װ����

class CAccountsEvents : public COleDispatchDriver
{
public:
	CAccountsEvents() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CAccountsEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAccountsEvents(const CAccountsEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// AccountsEvents ����
public:
	void AutoDiscoverComplete(LPDISPATCH Account)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc6c, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Account);
	}

	// AccountsEvents ����
public:

};
