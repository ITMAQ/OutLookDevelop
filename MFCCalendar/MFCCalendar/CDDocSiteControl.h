// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CDDocSiteControl ��װ����

class CDDocSiteControl : public COleDispatchDriver
{
public:
	CDDocSiteControl() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CDDocSiteControl(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDDocSiteControl(const CDDocSiteControl& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _DDocSiteControl ����
public:
	signed char get_ReadOnly()
	{
		signed char result;
		InvokeHelper(0x8001f008, DISPATCH_PROPERTYGET, VT_I1, (void*)&result, nullptr);
		return result;
	}
	void put_ReadOnly(signed char newValue)
	{
		static BYTE parms[] = VTS_I1;
		InvokeHelper(0x8001f008, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	signed char get_SuppressAttachments()
	{
		signed char result;
		InvokeHelper(0xfbe3, DISPATCH_PROPERTYGET, VT_I1, (void*)&result, nullptr);
		return result;
	}
	void put_SuppressAttachments(signed char newValue)
	{
		static BYTE parms[] = VTS_I1;
		InvokeHelper(0xfbe3, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}

	// _DDocSiteControl ����
public:

};
