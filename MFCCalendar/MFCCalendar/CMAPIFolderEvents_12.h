// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB" no_namespace
// CMAPIFolderEvents_12 ��װ����

class CMAPIFolderEvents_12 : public COleDispatchDriver
{
public:
	CMAPIFolderEvents_12() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CMAPIFolderEvents_12(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CMAPIFolderEvents_12(const CMAPIFolderEvents_12& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// MAPIFolderEvents_12 ����
public:
	void BeforeFolderMove(LPDISPATCH MoveTo, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xfba8, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, MoveTo, Cancel);
	}
	void BeforeItemMove(LPDISPATCH Item, LPDISPATCH MoveTo, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xfba9, DISPATCH_METHOD, VT_EMPTY, nullptr, parms, Item, MoveTo, Cancel);
	}

	// MAPIFolderEvents_12 ����
public:

};
