// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CApplicationEvents_10 ��װ����

class CApplicationEvents_10 : public COleDispatchDriver
{
public:
	CApplicationEvents_10() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CApplicationEvents_10(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplicationEvents_10(const CApplicationEvents_10& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// ApplicationEvents_10 ����
public:
	STDMETHOD(ItemSend)(LPDISPATCH Item, BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xf002, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Item, Cancel);
		return result;
	}
	STDMETHOD(NewMail)()
	{
		HRESULT result;
		InvokeHelper(0xf003, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	STDMETHOD(Reminder)(LPDISPATCH Item)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf004, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Item);
		return result;
	}
	STDMETHOD(OptionsPagesAdd)(LPDISPATCH Pages)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf005, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Pages);
		return result;
	}
	STDMETHOD(Startup)()
	{
		HRESULT result;
		InvokeHelper(0xf006, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	STDMETHOD(Quit)()
	{
		HRESULT result;
		InvokeHelper(0xf007, DISPATCH_METHOD, VT_HRESULT, (void*)&result, nullptr);
		return result;
	}
	void AdvancedSearchComplete(LPDISPATCH SearchObject)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfa6a, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, SearchObject);
	}
	void AdvancedSearchStopped(LPDISPATCH SearchObject)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xfa6b, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, SearchObject);
	}
	void MAPILogonComplete()
	{
		InvokeHelper(0xfa90, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// ApplicationEvents_10 ����
public:

};
