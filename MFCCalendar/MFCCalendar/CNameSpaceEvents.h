// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CNameSpaceEvents ��װ����

class CNameSpaceEvents : public COleDispatchDriver
{
public:
	CNameSpaceEvents() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CNameSpaceEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CNameSpaceEvents(const CNameSpaceEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// NameSpaceEvents ����
public:
	void OptionsPagesAdd(LPDISPATCH Pages, LPDISPATCH Folder)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH;
		InvokeHelper(0xf005, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Pages, Folder);
	}
	void AutoDiscoverComplete()
	{
		InvokeHelper(0xfc2d, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// NameSpaceEvents ����
public:

};
