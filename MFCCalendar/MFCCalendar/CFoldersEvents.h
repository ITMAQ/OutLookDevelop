// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CFoldersEvents ��װ����

class CFoldersEvents : public COleDispatchDriver
{
public:
	CFoldersEvents() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CFoldersEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFoldersEvents(const CFoldersEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// FoldersEvents ����
public:
	void FolderAdd(LPDISPATCH Folder)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf001, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Folder);
	}
	void FolderChange(LPDISPATCH Folder)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xf002, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Folder);
	}
	void FolderRemove()
	{
		InvokeHelper(0xf003, DISPATCH_METHOD, VT_EMPTY, nullptr, nullptr);
	}

	// FoldersEvents ����
public:

};
