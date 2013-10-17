#pragma once

// ActiveXDemoPropPage.h : Declaration of the CActiveXDemoPropPage property page class.


// CActiveXDemoPropPage : See ActiveXDemoPropPage.cpp for implementation.

class CActiveXDemoPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CActiveXDemoPropPage)
	DECLARE_OLECREATE_EX(CActiveXDemoPropPage)

// Constructor
public:
	CActiveXDemoPropPage();

// Dialog Data
	enum { IDD = IDD_PROPPAGE_ACTIVEXDEMO };

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	DECLARE_MESSAGE_MAP()
};

