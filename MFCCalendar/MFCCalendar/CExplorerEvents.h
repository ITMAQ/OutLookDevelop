// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CExplorerEvents ��װ����

class CExplorerEvents : public COleDispatchDriver
{
public:
	CExplorerEvents() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CExplorerEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CExplorerEvents(const CExplorerEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// ExplorerEvents ����
public:
	void Activate()
	{
		InvokeHelper(0xf001, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void FolderSwitch()
	{
		InvokeHelper(0xf002, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void BeforeFolderSwitch(LPDISPATCH NewFolder, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xf003, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, NewFolder, Cancel);
	}
	void ViewSwitch()
	{
		InvokeHelper(0xf004, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void BeforeViewSwitch(VARIANT& NewView, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_VARIANT VTS_PBOOL;
		InvokeHelper(0xf005, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, &NewView, Cancel);
	}
	void Deactivate()
	{
		InvokeHelper(0xf006, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void SelectionChange()
	{
		InvokeHelper(0xf007, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}
	void Close()
	{
		InvokeHelper(0xf008, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// ExplorerEvents ����
public:

};
