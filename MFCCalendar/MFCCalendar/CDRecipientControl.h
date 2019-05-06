// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CDRecipientControl ��װ����

class CDRecipientControl : public COleDispatchDriver
{
public:
	CDRecipientControl() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CDRecipientControl(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDRecipientControl(const CDRecipientControl& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _DRecipientControl ����
public:
	signed char get_Enabled()
	{
		signed char result;
		InvokeHelper(DISPID_ENABLED, DISPATCH_PROPERTYGET, VT_I1, (void*)&result, nullptr);
		return result;
	}
	void put_Enabled(signed char newValue)
	{
		static BYTE parms[] = VTS_I1;
		InvokeHelper(DISPID_ENABLED, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_BackColor()
	{
		long result;
		InvokeHelper(DISPID_BACKCOLOR, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_BackColor(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(DISPID_BACKCOLOR, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_ForeColor()
	{
		long result;
		InvokeHelper(DISPID_FORECOLOR, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_ForeColor(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(DISPID_FORECOLOR, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
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
	LPDISPATCH get_Font()
	{
		LPDISPATCH result;
		InvokeHelper(DISPID_FONT, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, nullptr);
		return result;
	}
	void put_Font(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(DISPID_FONT, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}
	long get_SpecialEffect()
	{
		long result;
		InvokeHelper(0xc, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, nullptr);
		return result;
	}
	void put_SpecialEffect(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xc, DISPATCH_PROPERTYPUT, VT_EMPTY, nullptr, parms, newValue);
	}

	// _DRecipientControl ����
public:

};
