// ActiveXDemoPropPage.cpp : Implementation of the CActiveXDemoPropPage property page class.

#include "stdafx.h"
#include "ActiveXDemo.h"
#include "ActiveXDemoPropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


IMPLEMENT_DYNCREATE(CActiveXDemoPropPage, COlePropertyPage)



// Message map

BEGIN_MESSAGE_MAP(CActiveXDemoPropPage, COlePropertyPage)
END_MESSAGE_MAP()



// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CActiveXDemoPropPage, "ACTIVEXDEMO.ActiveXDemoPropPage.1",
	0x40d0cfc1, 0x9768, 0x48b0, 0xa7, 0xeb, 0x90, 0xf5, 0xa7, 0x1, 0xb1, 0xf6)



// CActiveXDemoPropPage::CActiveXDemoPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CActiveXDemoPropPage

BOOL CActiveXDemoPropPage::CActiveXDemoPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_ACTIVEXDEMO_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}



// CActiveXDemoPropPage::CActiveXDemoPropPage - Constructor

CActiveXDemoPropPage::CActiveXDemoPropPage() :
	COlePropertyPage(IDD, IDS_ACTIVEXDEMO_PPG_CAPTION)
{
}



// CActiveXDemoPropPage::DoDataExchange - Moves data between page and properties

void CActiveXDemoPropPage::DoDataExchange(CDataExchange* pDX)
{
	DDP_PostProcessing(pDX);
}



// CActiveXDemoPropPage message handlers
