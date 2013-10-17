#pragma once

// ActiveXDemoCtrl.h : Declaration of the CActiveXDemoCtrl ActiveX Control class.


// CActiveXDemoCtrl : See ActiveXDemoCtrl.cpp for implementation.

class CActiveXDemoCtrl : public COleControl
{
	DECLARE_DYNCREATE(CActiveXDemoCtrl)

// Constructor
public:
	CActiveXDemoCtrl();
	SOCKET OnlySock;//建立的唯一Socket,不允许重复建立多个

	bool isOnlyConnect;//是否建立了连接

// Overrides
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();

// Implementation
protected:
	~CActiveXDemoCtrl();

	DECLARE_OLECREATE_EX(CActiveXDemoCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CActiveXDemoCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CActiveXDemoCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CActiveXDemoCtrl)		// Type name and misc status

// Message maps
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	DECLARE_DISPATCH_MAP()

// Event maps
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	enum {
		dispidDisconnect = 2L,
		dispidConnect = 1L
	};
protected:
	LONG Connect(VARIANT &ip, LONG port);
	LONG Disconnect(void);
};

