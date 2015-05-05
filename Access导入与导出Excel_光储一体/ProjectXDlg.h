// ProjectXDlg.h : header file
//
//{{AFX_INCLUDES()
#include "msflexgrid.h"
//}}AFX_INCLUDES

#if !defined(AFX_PROJECTXDLG_H__E42F8ACB_5720_4BE5_9C39_C7161B8F1815__INCLUDED_)
#define AFX_PROJECTXDLG_H__E42F8ACB_5720_4BE5_9C39_C7161B8F1815__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "DataStruct.h"

/////////////////////////////////////////////////////////////////////////////
// CProjectXDlg dialog

class CProjectXDlg : public CDialog
{
// Construction
public:
	CProjectXDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CProjectXDlg)
	enum { IDD = IDD_PROJECTX_DIALOG };
	CMSFlexGrid	m_FlexGrid;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CProjectXDlg)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation

//工具条和状态栏
protected:  // control bar embedded members

protected:
	HICON m_hIcon;

    CRect m_rect;                   //对话框原始大小（调整控件大小用）

	_ConnectionPtr m_pConnection; 	//智能指针
	_RecordsetPtr m_pRecordset; 

	CString strDBFile;              //access绝对地址
	CString strExcleFile;
    void Refresh();                 //刷新函数
                                    
    CString strValue1, strValue2, strValue3, strValue4, strValue5, strValue6, strValue7;  //读入Excel单元格数据
	
	// Generated message map functions
	//{{AFX_MSG(CProjectXDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	virtual void OnOK();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnExit();
	afx_msg void OnAbout();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnShowCmd();
	afx_msg void OnShowParam();
	afx_msg void OnShowCoeft();
	afx_msg void OnShowLogtype();
	afx_msg void OnShowLog();
	afx_msg void OnShowMainparam();
	afx_msg void OnShowMainswith();
	afx_msg void OnShowWaveparam();
	afx_msg void OnSelectDevice();
	afx_msg void OnShowDevice();
	afx_msg void OnLoadinNow();
	afx_msg void OnExportoutNow();
	afx_msg void OnExportAll();
	afx_msg void OnExportDevicetype();
	afx_msg void OnLoadinDevicetype();
	afx_msg void OnLoadinAll();
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
  
public:                              
	DeviceTypeList m_DeviceTypeList;//所有设备信息
	ST_DeviceType  m_StDeviceSelectNow;//当前选择的设备
	int m_iExcleType;//当前查看的表格
public:
	void InitDeviceTypeList();
	void InitExcleType(int iType,BOOL bShowtip = TRUE);
	int	 GetExcleCnt(int iType);
	void SetExcleColWith(int iType);
	void LoadInfoFormDataBase(int iType,BOOL bFresh = FALSE);
	CString GetSelectSentence(int iType);
	CString GetDeleteSentence(int iType);
	void ProgressShow(int oPera,CString strName,int iVal);
	BOOL ExportExcle(BOOL bTip = FALSE,CString strPath = "");
	void ExportDeviceExcle(CString strPath = "");
	CString SHBrowseForFolder_DirectoryDlg();
	void LoadInDeviceExcle(CString strPath = "");
	BOOL LoadInExcle(CString strPath = "");
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PROJECTXDLG_H__E42F8ACB_5720_4BE5_9C39_C7161B8F1815__INCLUDED_)
