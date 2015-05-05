#if !defined(AFX_DLGDEVICESELECT_H__038F294C_C275_40AD_B992_A3D0A30B9AED__INCLUDED_)
#define AFX_DLGDEVICESELECT_H__038F294C_C275_40AD_B992_A3D0A30B9AED__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DlgDeviceSelect.h : header file
//
#include "DataStruct.h"
/////////////////////////////////////////////////////////////////////////////
// CDlgDeviceSelect dialog

class CDlgDeviceSelect : public CDialog
{
// Construction
public:
	CDlgDeviceSelect(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CDlgDeviceSelect)
	enum { IDD = IDD_DIALOG_SELECTDEVICE };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDlgDeviceSelect)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CDlgDeviceSelect)
	virtual BOOL OnInitDialog();
	afx_msg void OnSelchangeComboDevice();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
public:
	DeviceTypeList m_TypeList;
	int m_iSelect;
	void InitDevice(DeviceTypeList TypeList,int iSelect);
public:
	afx_msg void OnBnClickedOk();
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGDEVICESELECT_H__038F294C_C275_40AD_B992_A3D0A30B9AED__INCLUDED_)
