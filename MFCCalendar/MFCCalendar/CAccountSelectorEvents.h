// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CAccountSelectorEvents ��װ����

class CAccountSelectorEvents : public COleDispatchDriver
{
public:
	CAccountSelectorEvents() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CAccountSelectorEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAccountSelectorEvents(const CAccountSelectorEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// AccountSelectorEvents ����
public:
	void SelectedAccountChange(LPDISPATCH SelectedAccount)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfc73, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, SelectedAccount);
	}

	// AccountSelectorEvents ����
public:

};
