// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CFormRegionEvents ��װ����

class CFormRegionEvents : public COleDispatchDriver
{
public:
	CFormRegionEvents() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CFormRegionEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFormRegionEvents(const CFormRegionEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// FormRegionEvents ����
public:
	void Expanded(BOOL Expand)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0xfb38, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Expand);
	}
	void Close()
	{
		InvokeHelper(0xf004, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// FormRegionEvents ����
public:

};
