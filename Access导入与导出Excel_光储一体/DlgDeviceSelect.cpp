// DlgDeviceSelect.cpp : implementation file
//

#include "stdafx.h"
#include "ProjectX.h"
#include "DlgDeviceSelect.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDlgDeviceSelect dialog


CDlgDeviceSelect::CDlgDeviceSelect(CWnd* pParent /*=NULL*/)
	: CDialog(CDlgDeviceSelect::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDlgDeviceSelect)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CDlgDeviceSelect::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDlgDeviceSelect)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CDlgDeviceSelect, CDialog)
	//{{AFX_MSG_MAP(CDlgDeviceSelect)
	ON_CBN_SELCHANGE(IDC_COMBO_DEVICE, OnSelchangeComboDevice)
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDOK, &CDlgDeviceSelect::OnBnClickedOk)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDlgDeviceSelect message handlers

BOOL CDlgDeviceSelect::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// TODO: Add extra initialization here
	for (int i=0;i<(int)m_TypeList.size();i++)
	{
		((CComboBox *)GetDlgItem(IDC_COMBO_DEVICE))->AddString(m_TypeList[i].cDeviceName);
	}
	((CComboBox *)GetDlgItem(IDC_COMBO_DEVICE))->SetCurSel(m_iSelect);
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CDlgDeviceSelect::InitDevice( DeviceTypeList TypeList,int iSelect )
{
	m_TypeList = TypeList;
	m_iSelect = iSelect;
}

void CDlgDeviceSelect::OnSelchangeComboDevice() 
{
	// TODO: Add your control notification handler code here
	m_iSelect = ((CComboBox *)GetDlgItem(IDC_COMBO_DEVICE))->GetCurSel();
}

void CDlgDeviceSelect::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	OnOK();
}
