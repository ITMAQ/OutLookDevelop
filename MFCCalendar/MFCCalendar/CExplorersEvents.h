// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CExplorersEvents ��װ����

class CExplorersEvents : public COleDispatchDriver
{
public:
	CExplorersEvents() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CExplorersEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CExplorersEvents(const CExplorersEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// ExplorersEvents ����
public:
	void NewExplorer(LPDISPATCH Explorer)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf001, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Explorer);
	}

	// ExplorersEvents ����
public:

};
