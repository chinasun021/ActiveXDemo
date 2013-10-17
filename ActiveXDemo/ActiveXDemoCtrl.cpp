// ActiveXDemoCtrl.cpp : Implementation of the CActiveXDemoCtrl ActiveX Control class.

#include "stdafx.h"
#include "ActiveXDemo.h"
#include "ActiveXDemoCtrl.h"
#include "ActiveXDemoPropPage.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#ifndef WM_MYWINSOCK 
#define WM_MYWINSOCK WM_USER+1888
#endif

IMPLEMENT_DYNCREATE(CActiveXDemoCtrl, COleControl)



// Message map

BEGIN_MESSAGE_MAP(CActiveXDemoCtrl, COleControl)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()



// Dispatch map

BEGIN_DISPATCH_MAP(CActiveXDemoCtrl, COleControl)
	DISP_FUNCTION_ID(CActiveXDemoCtrl, "Connect", dispidConnect, Connect, VT_I4, VTS_VARIANT VTS_I4)
	DISP_FUNCTION_ID(CActiveXDemoCtrl, "Disconnect", dispidDisconnect, Disconnect, VT_I4, VTS_NONE)
END_DISPATCH_MAP()



// Event map

BEGIN_EVENT_MAP(CActiveXDemoCtrl, COleControl)
END_EVENT_MAP()



// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CActiveXDemoCtrl, 1)
	PROPPAGEID(CActiveXDemoPropPage::guid)
END_PROPPAGEIDS(CActiveXDemoCtrl)



// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CActiveXDemoCtrl, "ACTIVEXDEMO.ActiveXDemoCtrl.1",
	0xa6c57f0c, 0xc1ef, 0x4e94, 0xac, 0xc, 0xcd, 0xc9, 0xd5, 0x7a, 0xe4, 0x2c)



// Type library ID and version

IMPLEMENT_OLETYPELIB(CActiveXDemoCtrl, _tlid, _wVerMajor, _wVerMinor)



// Interface IDs

const IID BASED_CODE IID_DActiveXDemo =
		{ 0x768223FA, 0x5D02, 0x489A, { 0xB4, 0x3E, 0x85, 0x1F, 0x22, 0x99, 0xC4, 0xBD } };
const IID BASED_CODE IID_DActiveXDemoEvents =
		{ 0x5D558BC6, 0x9E0C, 0x4A74, { 0x94, 0x1A, 0xB5, 0x30, 0xE9, 0x35, 0xB6, 0x64 } };



// Control type information

static const DWORD BASED_CODE _dwActiveXDemoOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CActiveXDemoCtrl, IDS_ACTIVEXDEMO, _dwActiveXDemoOleMisc)



// CActiveXDemoCtrl::CActiveXDemoCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CActiveXDemoCtrl

BOOL CActiveXDemoCtrl::CActiveXDemoCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Verify that your control follows apartment-model threading rules.
	// Refer to MFC TechNote 64 for more information.
	// If your control does not conform to the apartment-model rules, then
	// you must modify the code below, changing the 6th parameter from
	// afxRegApartmentThreading to 0.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_ACTIVEXDEMO,
			IDB_ACTIVEXDEMO,
			afxRegApartmentThreading,
			_dwActiveXDemoOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}



// CActiveXDemoCtrl::CActiveXDemoCtrl - Constructor

CActiveXDemoCtrl::CActiveXDemoCtrl()
{
	InitializeIIDs(&IID_DActiveXDemo, &IID_DActiveXDemoEvents);
	// TODO: Initialize your control's instance data here.
	isOnlyConnect = false;
}



// CActiveXDemoCtrl::~CActiveXDemoCtrl - Destructor

CActiveXDemoCtrl::~CActiveXDemoCtrl()
{
	// TODO: Cleanup your control's instance data here.
	if(isOnlyConnect)//���������ѽ���,��������û�������Ͽ����˳�����,�ؼ��Զ��Ͽ�����,�ͷ���Դ
	{
		shutdown(OnlySock,0x02);
		closesocket(OnlySock);//�ͷ�ռ�е�SOCK��Դ
		isOnlyConnect = false;
	}
}



// CActiveXDemoCtrl::OnDraw - Drawing function

void CActiveXDemoCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	if (!pdc)
		return;

	// TODO: Replace the following code with your own drawing code.
	pdc->FillRect(rcBounds, CBrush::FromHandle((HBRUSH)GetStockObject(WHITE_BRUSH)));
	pdc->Ellipse(rcBounds);
}



// CActiveXDemoCtrl::DoPropExchange - Persistence support

void CActiveXDemoCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: Call PX_ functions for each persistent custom property.
}



// CActiveXDemoCtrl::OnResetState - Reset control to default state

void CActiveXDemoCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}



// CActiveXDemoCtrl message handlers

LONG CActiveXDemoCtrl::Connect(VARIANT &ip, LONG port)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your dispatch handler code here
	char *IPAddress;
	CString temp(ip.bstrVal);
	IPAddress=temp.GetBuffer(temp.GetLength());
	temp.ReleaseBuffer();
	if(isOnlyConnect)//�������ѽ���,��û�жϿ�
		return 0;
	struct hostent *host = 0;
	struct sockaddr_in addr;
	ULONG dotIP = inet_addr(IPAddress);

	OnlySock = socket(AF_INET, SOCK_STREAM, 0);
	// �ж��Ƿ�Ϊ��IP��ַ��ʽ
    if (OnlySock == INVALID_SOCKET)
	{
		shutdown(OnlySock, 0x02);
		closesocket(OnlySock);//�ͷ�ռ�е�SOCK��Դ
		return 0;
	}

	// �趨 SOCKADDR_IN �ṹ������
    // ���ͨѶЭ����ѡ��IP Protocol���Ǵ�ֵ�̶�ΪAF_INET
    // AF_INET �� PF_INET ����������ֵ��ͬ
	memset(&addr,0,sizeof(addr));
    addr.sin_family = AF_INET;
    addr.sin_port = htons(port);
    addr.sin_addr.s_addr = dotIP;
	 //��ʼ����
    if (connect(OnlySock, (LPSOCKADDR)&addr, sizeof(SOCKADDR))<0)
    {
        shutdown(OnlySock, 0x02);
		closesocket(OnlySock);//�ͷ�ռ�е�SOCK��Դ
		return 0;
    }
	isOnlyConnect = true;

	char sendline[1000]="hello MFC ActiveX Socket!";
	if( send(OnlySock, sendline, strlen(sendline), 0) < 0)
    {
		shutdown(OnlySock, 0x02);
		closesocket(OnlySock);//�ͷ�ռ�е�SOCK��Դ
		return 0;
	}
	return 1;
}

LONG CActiveXDemoCtrl::Disconnect(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your dispatch handler code here
	if(!isOnlyConnect)//��ǰû���κ�����
		return 0;
		
	shutdown(OnlySock,0x02);
	closesocket(OnlySock);//�ͷ�ռ�е�SOCK��Դ
	isOnlyConnect = false;

	return 1;
}
