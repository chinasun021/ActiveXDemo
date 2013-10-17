// ActiveXDemo.cpp : Implementation of CActiveXDemoApp and DLL registration.

#include "stdafx.h"
#include "ActiveXDemo.h"
#include  <comcat.h>     
#include  <objsafe.h> 
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


CActiveXDemoApp theApp;

const GUID CDECL BASED_CODE _tlid =
		{ 0x41F29E03, 0xE746, 0x4F49, { 0x82, 0x12, 0x8F, 0x5E, 0xD9, 0x21, 0xAC, 0x70 } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;



// CActiveXDemoApp::InitInstance - DLL initialization

BOOL CActiveXDemoApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();

	if (bInit)
	{
		// TODO: Add your own module initialization code here.
		if (!AfxSocketInit())
		{
			AfxMessageBox(_T("�޷���ʼ��Socket,����!"));
			return FALSE;
		}
		WSADATA wsaData;
		WORD wVersion = MAKEWORD(1, 1);//�趨ΪWinsock 1.1��
		int errCode;		
		errCode = WSAStartup(wVersion, &wsaData);//����Socket����
		if (errCode)
		{
			AfxMessageBox(_T("�޷��ҵ�����ʹ�õ� WSOCK32.DLL"));
			return FALSE;
		}
	}

	return bInit;
}



// CActiveXDemoApp::ExitInstance - DLL termination

int CActiveXDemoApp::ExitInstance()
{
	// TODO: Add your own module termination code here.
	WSACleanup();//�����������

	return COleControlModule::ExitInstance();
}



//  �����������     
HRESULT CreateComponentCategory(CATID catid, WCHAR *  catDescription) 
{    
    ICatRegister *  pcr  =  NULL ;    
    HRESULT hr  =  S_OK ;    
    hr  =  CoCreateInstance(CLSID_StdComponentCategoriesMgr, NULL, CLSCTX_INPROC_SERVER, IID_ICatRegister, ( void ** ) & pcr);    
     if  (FAILED(hr))  return  hr;    
     //  Make sure the HKCR\Component Categories\{..catid...}    
     //  key is registered.     
    CATEGORYINFO catinfo;    
    catinfo.catid  =  catid;    
    catinfo.lcid  =   0x0409  ;  //  english    
     //  Make sure the provided description is not too long.    
     //  Only copy the first 127 characters if it is.     
     int  len  =  wcslen(catDescription);    
     if  (len > 127 ) len  =   127 ;    
    wcsncpy(catinfo.szDescription, catDescription, len);    
     //  Make sure the description is null terminated.     
    catinfo.szDescription[len]  =   ' \0 ' ;    
    hr  =  pcr -> RegisterCategories( 1 ,  & catinfo);    
    pcr -> Release();    
     return  hr;    
}  

//  ע���������     
HRESULT RegisterCLSIDInCategory(REFCLSID clsid, CATID catid)
 {    
     //  Register your component categories information.     
    ICatRegister *  pcr  =  NULL ;    
    HRESULT hr  =  S_OK ;    
    hr  =  CoCreateInstance(CLSID_StdComponentCategoriesMgr, NULL, CLSCTX_INPROC_SERVER, IID_ICatRegister, ( void ** ) & pcr);    
     if  (SUCCEEDED(hr)) {    
       //  Register this category as being "implemented" by the class.     
      CATID rgcatid[ 1 ];    
      rgcatid[ 0 ]  =  catid;    
      hr  =  pcr -> RegisterClassImplCategories(clsid,  1 , rgcatid);    
    }    
     if  (pcr  !=  NULL) pcr -> Release();    
     return  hr;    
}    
//  ж���������     
HRESULT UnRegisterCLSIDInCategory(REFCLSID clsid, CATID catid)
 {    
    ICatRegister *  pcr  =  NULL ;    
    HRESULT hr  =  S_OK ;    
    hr  =  CoCreateInstance(CLSID_StdComponentCategoriesMgr,    
            NULL, CLSCTX_INPROC_SERVER, IID_ICatRegister, ( void ** ) & pcr);    
     if  (SUCCEEDED(hr)) {    
       //  Unregister this category as being "implemented" by the class.     
      CATID rgcatid[ 1 ] ;    
      rgcatid[ 0 ]  =  catid;    
      hr  =  pcr -> UnRegisterClassImplCategories(clsid,  1 , rgcatid);    
    }    
     if  (pcr  !=  NULL) pcr -> Release();    
     return  hr;    
}    
STDAPI DllRegisterServer( void ) 
{    
    HRESULT hr;    
    AFX_MANAGE_STATE(_afxModuleAddrThis);    
     if  ( ! AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))    
         return  ResultFromScode(SELFREG_E_TYPELIB);    
     if  ( ! COleObjectFactoryEx::UpdateRegistryAll(TRUE))    
         return  ResultFromScode(SELFREG_E_CLASS);    
     //  ��ǿؼ���ʼ����ȫ.    
     //  ������ʼ����ȫ�������     
    hr  =  CreateComponentCategory(CATID_SafeForInitializing, L" Controls safely initializable from persistent data! " );    
     if  (FAILED(hr))  return  hr;    
     //  ע���ʼ����ȫ     
    hr  =  RegisterCLSIDInCategory(BASED_CODE _tlid , CATID_SafeForInitializing);    
     if  (FAILED(hr))  return  hr;    
     //  ��ǿؼ��ű���ȫ    
     //  �����ű���ȫ�������     
    hr  =  CreateComponentCategory(CATID_SafeForScripting, L" Controls safely scriptable! " );    
     if  (FAILED(hr))  return  hr;    
     //  ע��ű���ȫ�������     
    hr  =  RegisterCLSIDInCategory(BASED_CODE _tlid , CATID_SafeForScripting);    
     if  (FAILED(hr))  return  hr;    
     return  NOERROR;    
}    
//  DllUnregisterServer - Removes entries from the system registry     
STDAPI DllUnregisterServer( void ) 
{    
    HRESULT hr;    
    AFX_MANAGE_STATE(_afxModuleAddrThis);    
     if  ( ! AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))    
         return  ResultFromScode(SELFREG_E_TYPELIB);    
     if  ( ! COleObjectFactoryEx::UpdateRegistryAll(FALSE))    
         return  ResultFromScode(SELFREG_E_CLASS);    
     //  ɾ���ؼ���ʼ����ȫ���.     
    hr = UnRegisterCLSIDInCategory(BASED_CODE _tlid , CATID_SafeForInitializing);    
     if  (FAILED(hr))  return  hr;    
     //  ɾ���ؼ��ű���ȫ���     
    hr = UnRegisterCLSIDInCategory(BASED_CODE _tlid , CATID_SafeForScripting);    
     if  (FAILED(hr))  return  hr;    
     return  NOERROR;    
}  