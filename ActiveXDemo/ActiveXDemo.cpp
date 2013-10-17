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
			AfxMessageBox(_T("无法初始化Socket,请检查!"));
			return FALSE;
		}
		WSADATA wsaData;
		WORD wVersion = MAKEWORD(1, 1);//设定为Winsock 1.1版
		int errCode;		
		errCode = WSAStartup(wVersion, &wsaData);//启动Socket服务
		if (errCode)
		{
			AfxMessageBox(_T("无法找到可以使用的 WSOCK32.DLL"));
			return FALSE;
		}
	}

	return bInit;
}



// CActiveXDemoApp::ExitInstance - DLL termination

int CActiveXDemoApp::ExitInstance()
{
	// TODO: Add your own module termination code here.
	WSACleanup();//结束网络服务

	return COleControlModule::ExitInstance();
}



//  创建组件种类     
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

//  注册组件种类     
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
//  卸载组件种类     
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
     //  标记控件初始化安全.    
     //  创建初始化安全组件种类     
    hr  =  CreateComponentCategory(CATID_SafeForInitializing, L" Controls safely initializable from persistent data! " );    
     if  (FAILED(hr))  return  hr;    
     //  注册初始化安全     
    hr  =  RegisterCLSIDInCategory(BASED_CODE _tlid , CATID_SafeForInitializing);    
     if  (FAILED(hr))  return  hr;    
     //  标记控件脚本安全    
     //  创建脚本安全组件种类     
    hr  =  CreateComponentCategory(CATID_SafeForScripting, L" Controls safely scriptable! " );    
     if  (FAILED(hr))  return  hr;    
     //  注册脚本安全组件种类     
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
     //  删除控件初始化安全入口.     
    hr = UnRegisterCLSIDInCategory(BASED_CODE _tlid , CATID_SafeForInitializing);    
     if  (FAILED(hr))  return  hr;    
     //  删除控件脚本安全入口     
    hr = UnRegisterCLSIDInCategory(BASED_CODE _tlid , CATID_SafeForScripting);    
     if  (FAILED(hr))  return  hr;    
     return  NOERROR;    
}  