

//���ߣ�Ԭ��     QQ:����
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 1������Ĭ�ϵĻ����Ի��򹤳̣�������ΪProjectX��
// 2����stdafx.h����ӵ���ADO��
// 3����ProjectXDlg.h����ӱ���������ָ������ͱ�Ǽ�¼�������ı�����
// 4����ProjectXDlg.cpp����ӳ�ʼ������ȡ�
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ���MSFlexGrid�ؼ�
// 1��Ctrl+W�����򵼣�����->���ӵ�����->Components and Contols->Registered ActiveX Controls->Microsoft FlexGrid Control ,version6.0 ->Insert
// 2��ΪIDC_DATAGRID1��������m_FlexGrid1
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ���Excel
// 1��Ctrl+W�����򵼣��½�һ���࣬ѡ���Type Library��ӡ������Office 2003����ӵ���Office��װ·���µ�Excel.exe (��Office 2000��������ӵ�Ӧ����Excel9.OLB) �� 
//    �ڵ�����Confirm Classes��ѡ��_Application��Workbooks��_Workbook��Worksheets ��_Worksheet��Range ��Font �⼸���࣬
//    ��ȷ�������ɵ�.CPP��.h�ļ�������ΪExcel.cpp��Excel.h��Ȼ��ȷ����
// 2����ProjectXDlg.cpp�����ͷ�ļ����ã�#include "Excel.h"
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ��Ҫ��ʾ��
// 1����ʱ������ڴ�Ϊд�Ĵ��󣬽���������齨->ȫ���ؽ�
// 2�������Excel�ļ����������úõı�׼��ͨѶ¼ģ��
// 3������ʵ���Ͼ��ǲ��룬��������������ݣ�ֻҪ�����е����ݴ�Excel����ɾ�����ɡ�
// 4��Access���ݿ�����Ϊ:111111
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// ProjectXDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ProjectX.h"
#include "ProjectXDlg.h"
#include "DlgDeviceSelect.h"
#include <shlwapi.h>
#pragma comment(lib,"Shlwapi.lib")

#include "excel.h"              //����Excel 
using namespace excel9;

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//==����======================================================B
#define TABLE_TYPE_CNT 9
#define TABLE_DEVICE_CNT 6 //��Ŀ���ǲ�����������,����˵���
#define TABLE_PARAM_CNT 10
#define TABLE_COEFT_CNT 3
#define TABLE_CMD_CNT 5 
#define	TABLE_LOGTYPE_CNT 2
#define	TABLE_LOG_CNT 4 //��������������ASC��Ȼ����������ASC(DESC����)
#define TABLE_MAINPARAM_CNT 2
#define TABLE_MAINSWITH_CNT 2
#define TABLE_WAVEPARAM_CNT 5

#define CLIENT_DATABASE_DEF 2	//0��photovol.dll��1��photovol_����.dll��2��photovol_�ͻ�ר��.dll
enum TableType
{
	Table_DeciceType = 0,//�豸����
	Table_Param,//������Ϣ
	Table_Coeft,//���ϵ��
	Table_Cmd,//�����ַ
	Table_LogType,//������־����
	Table_Log,//��־
	Table_MainParam,//�����������Ϣ
	Table_MainSwith,//�����濪��״̬
	Table_WaveParam//������Ϣ
};

enum ProGressType
{
	ProgressType_LOOK=0,
	ProgressType_IN,
	ProgressType_OUT,
};

CString g_strExcleTitle[TABLE_TYPE_CNT] = {"�豸����","������Ϣ","���ϵ��","�����ַ","������־����","������־","��������ʾ����","�����濪�ص�ַ","���ν���ͨ������"};//����

CString g_strTableDeciceType[TABLE_DEVICE_CNT] = {"�豸��������","�豸�ͺ�","����","��ʵֵ��ʾ","��ǰ������ʾ","����������Ե�"};//�豸����
CString g_strTableParam[TABLE_PARAM_CNT] = {"��ַ","ͨ�����","ϵͳ��","��ע��","���뵽FPGAֵ����ʵ��ֵ*���","���뵽FPGAֵ�������","��ʾλ","��������","���","��λ"};//������Ϣ��һ�н���Ÿ�Ϊ��ַ��λ�ű�Ϊ��ʾλ
CString g_strTableCoeft[TABLE_COEFT_CNT] = {"ϵͳ��","��ע��","��ַ"};//���ϵ��
CString g_strTableCmd[TABLE_CMD_CNT] = {"�����ַ","��������0","��������1","����λ��","�������"};//�����ַ
CString g_strTableLogType[TABLE_LOGTYPE_CNT] = {"����","��������"};//������־����
CString g_strTableLog[TABLE_LOG_CNT] = {"����","������","��־��","�Ĵ���λ"};//��־
CString g_strTableMainParam[TABLE_MAINPARAM_CNT] = {"����","��ַ"};//�����������Ϣ
CString g_strTableMainSwith[TABLE_MAINSWITH_CNT] = {"��ַ","��ֵ"};//�����濪��״̬
CString g_strTableWaveParam[TABLE_WAVEPARAM_CNT] = {"�����","ͨ����","ͨ��ֵ","���ϵ����ַ","����"};//������Ϣ
//==����======================================================E

// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	afx_msg void OnPaint();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

public:                              

};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

//���ñ���ɫ
void CAboutDlg::OnPaint() 
{
	CPaintDC dc(this); // device context for painting
	
	// TODO: Add your message handler code here

	CRect rect;                                                                     
    GetClientRect(rect); 
	dc.FillSolidRect(rect,RGB(50,130,200));

	// Do not call CDialog::OnPaint() for painting messages
}


HBRUSH CAboutDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor) 
{
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);
	
	// TODO: Change any attributes of the DC here

	// TODO: Return a different brush if the default is not desired
   if( nCtlColor == CTLCOLOR_STATIC)              //ʵ�־�̬�ı���͸����ʾ
	{   
       pDC->SetBkMode(TRANSPARENT);                
	   return   
	   HBRUSH(GetStockObject(HOLLOW_BRUSH));   
	}

	return hbr;
}

//��ʼ��
BOOL CAboutDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// TODO: Add extra initialization here


	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

//����Esc��
BOOL CAboutDlg::PreTranslateMessage(MSG* pMsg) 
{
	// TODO: Add your specialized code here and/or call the base class
	
	if(pMsg->message==WM_KEYDOWN) 
    { 
          switch(pMsg-> wParam) 
           { 
			  case VK_ESCAPE:                                   
                   return TRUE;  
            } 
     } 

	return CDialog::PreTranslateMessage(pMsg);
}

void CAboutDlg::OnOK() 
{
	// TODO: Add extra validation here

	 CDialog::OnOK();
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
	ON_WM_PAINT()
	ON_WM_CTLCOLOR()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CProjectXDlg dialog

CProjectXDlg::CProjectXDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CProjectXDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CProjectXDlg)
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CProjectXDlg::DoDataExchange(CDataExchange* pDX)  //���ݽ���
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CProjectXDlg)
	DDX_Control(pDX, IDC_MSFLEXGRID1, m_FlexGrid);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CProjectXDlg, CDialog)
	//{{AFX_MSG_MAP(CProjectXDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_SIZE()
	ON_WM_MOUSEWHEEL()
	ON_COMMAND(IDR_EXIT, OnExit)
	ON_COMMAND(IDR_ABOUT, OnAbout)
	ON_WM_CTLCOLOR()
	ON_COMMAND(ID_SHOW_CMD, OnShowCmd)
	ON_COMMAND(ID_SHOW_PARAM, OnShowParam)
	ON_COMMAND(ID_SHOW_LOGTYPE, OnShowLogtype)
	ON_COMMAND(ID_SHOW_LOG, OnShowLog)
	ON_COMMAND(ID_SHOW_MAINPARAM, OnShowMainparam)
	ON_COMMAND(ID_SHOW_MAINSWITH, OnShowMainswith)
	ON_COMMAND(ID_SHOW_WAVEPARAM, OnShowWaveparam)
	ON_COMMAND(IDR_SELECT_DEVICE, OnSelectDevice)
	ON_COMMAND(IDR_SHOW_DEVICE, OnShowDevice)
	ON_COMMAND(ID_LOADIN_NOW, OnLoadinNow)
	ON_COMMAND(ID_EXPORTOUT_NOW, OnExportoutNow)
	ON_COMMAND(ID_EXPORT_ALL, OnExportAll)
	ON_COMMAND(ID_EXPORT_DEVICETYPE, OnExportDevicetype)
	ON_COMMAND(ID_LOADIN_DEVICETYPE, OnLoadinDevicetype)
	ON_WM_RBUTTONDOWN()
	ON_WM_HSCROLL()
	ON_WM_VSCROLL()
	ON_WM_CTLCOLOR()
	ON_COMMAND(ID_LOADIN_ALL, OnLoadinAll)
	//}}AFX_MSG_MAP

	ON_COMMAND(ID_SHOW_COEFT, &CProjectXDlg::OnShowCoeft)
END_MESSAGE_MAP()

////////////////////////////////////
// CProjectXDlg message handlers
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//��ʼ��
BOOL CProjectXDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here

	char path[MAX_PATH];
	GetModuleFileName(NULL, path, MAX_PATH);        //��ȡ������·���磺E:\Tools\qq.exe
	*strrchr(path,'\\') = '\0';
	
    strDBFile = path;
	strExcleFile = strDBFile ;

	if(0==CLIENT_DATABASE_DEF)
		strDBFile += "\\photovol.dll";
	else if(1==CLIENT_DATABASE_DEF)
		strDBFile += "\\photovol_����.dll";
	else
		strDBFile += "\\photovol_�ͻ�ר��.dll";

	InitDeviceTypeList();//��ʼ����������
	m_iExcleType = Table_DeciceType;
	if ((int)m_DeviceTypeList.size()!=0)
	{
		memcpy(&m_StDeviceSelectNow,&m_DeviceTypeList[0],sizeof(ST_DeviceType));
		InitExcleType(m_iExcleType);
	}
    ((CProgressCtrl *)GetDlgItem(IDC_PROGRESS_LOAD))->SetPos(0);
    CenterWindow();                 //���������ʾ
    ShowWindow(SW_MAXIMIZE);        //�����ʾ

	return TRUE;   // return TRUE  unless you set the focus to a control
}

void CProjectXDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CProjectXDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CRect rect;                                      //���ñ���ɫ                          
        CPaintDC dc(this); 
        GetClientRect(rect); 
        dc.FillSolidRect(rect,RGB(210,240,255));
     
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CProjectXDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

BEGIN_EVENTSINK_MAP(CProjectXDlg, CDialog)
    //{{AFX_EVENTSINK_MAP(CProjectXDlg)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

//����Esc��
BOOL CProjectXDlg::PreTranslateMessage(MSG* pMsg)                  
{
	// TODO: Add your specialized code here and/or call the base class

	if(pMsg->message==WM_KEYDOWN) 
    { 
        switch(pMsg-> wParam) 
        { 
          	  
	       case VK_ESCAPE:                                       
               return TRUE;  
        }   
	} 

	return CDialog::PreTranslateMessage(pMsg);
}

//ע��CDialog::OnOK(),������FlexGrid��ʱ����Enter�������˳��Ի��� 
void CProjectXDlg::OnOK()                                        
{
	// TODO: Add extra validation here
	
	//CDialog::OnOK();                   
}

//�����ռ��С
void CProjectXDlg::OnSize(UINT nType, int cx, int cy)   //�Ի����С
{
	CDialog::OnSize(nType, cx, cy);
	
	// TODO: Add your message handler code here
    CWnd *pWnd; 
    
	//����FlexGrid1λ�úʹ�С
	if(nType==1) return;                            //���������С���򷵻�
	pWnd = GetDlgItem(IDC_MSFLEXGRID1);             //��ȡ�ؼ����
	if(pWnd)                                        //�ж��Ƿ�Ϊ�գ���Ϊ�Ի��򴴽�ʱ����ô˺���������ʱ�ؼ���δ����
    {
      CRect rect;                                   //�仯ǰ��С
      pWnd->GetWindowRect(&rect);
      
	  ScreenToClient(&rect);                        
      
      rect.left=rect.left*cx/m_rect.Width();        //���������С
      rect.right=rect.right*cx/m_rect.Width();    
      rect.top=rect.top*cy/m_rect.Height();         //���������С
      rect.bottom=rect.bottom*cy/m_rect.Height();
      
	  rect.bottom = cy -24;
	  pWnd->MoveWindow(rect);                       //���ÿؼ���С
	  GetDlgItem(IDC_PROGRESS_LOAD)->MoveWindow(m_rect.left+2,cy-22,600,20);
	  GetDlgItem(IDC_STATIC_1)->MoveWindow(m_rect.left+615,cy-20,500,20);
    }

	GetClientRect(&m_rect);                         //���仯��ĶԻ�����Ϊԭʼ��С	
}

//��Ӧ����
BOOL CProjectXDlg::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	// TODO: Add your message handler code here and/or call default
	                              
	int nPos = m_FlexGrid.GetScrollPos(SB_VERT);         
    int nMax = m_FlexGrid.GetRows();/*m_FlexGrid.GetScrollLimit(SB_VERT);*/       
   
	if(zDelta < 0)                                     
    {
        nPos += 5;
        (nPos >= nMax) ? (nPos = nMax-1) : NULL;
    }
    else
    {
        nPos -= 5;
        (nPos <= 1) ? (nPos = 1) : NULL;
    }
       
	m_FlexGrid.SetTopRow(nPos);                         

	
	return CDialog::OnMouseWheel(nFlags, zDelta, pt);
}


//ˢ�º���
void CProjectXDlg:: Refresh()
{
	LoadInfoFormDataBase(m_iExcleType,TRUE);
}   

//�˳�
void CProjectXDlg::OnExit() 
{
	// TODO: Add your command handler code here
	
	CDialog::OnOK(); 
}

//����
void CProjectXDlg::OnAbout() 
{
	// TODO: Add your command handler code here
	CAboutDlg dlg;
    dlg.DoModal();
}            

void CProjectXDlg::InitDeviceTypeList()
{
	m_DeviceTypeList.clear();

	//�������ݿ�                                                         
	CString strConnection;
	strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="+strDBFile;
	m_pConnection.CreateInstance(__uuidof(Connection));
	m_pConnection->CursorLocation = adUseClient;                
	m_pRecordset.CreateInstance(__uuidof(Recordset)); 
	m_pConnection->Open((LPCTSTR)strConnection, "", "", adModeUnknown);
	
	//�����ݱ�
	
	CString strText;
	strText.Format("select * from %s order by ���",g_strExcleTitle[Table_DeciceType]);
	m_pRecordset->PutCursorLocation(adUseClient);                 
	m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	
	while(!m_pRecordset->EndOfFile)                                //���û�е���¼�����
	{
		CString strValue;
		_variant_t vstr;
		ST_DeviceType oParam;
		
		for(int i = 0; i < TABLE_DEVICE_CNT+1; i++)                                 //Ϊÿ�е�Ԫ��ֵ
		{
			vstr = m_pRecordset->GetCollect(_variant_t(long(i)));
			if(vstr.vt!=VT_NULL)
				strValue = (LPCSTR)_bstr_t(vstr);
			else
				strValue = "";
			strValue.TrimLeft();
			strValue.TrimRight();
			if (i==0)
			{
				oParam.iID = atoi(strValue);
			}
			else if (i==1)
			{
				strcpy(oParam.cDeviceName,strValue);
			}
			else if (i==2)
			{
				oParam.iType= atoi(strValue);
			}
			else if (i==3)
			{
				strcpy(oParam.cDescribe,strValue);
			}
		}
		m_DeviceTypeList.push_back(oParam);
		m_pRecordset->MoveNext();                                  //����һ��                                                 //�����Լ�1
	}
	
	m_pRecordset->Close();                                         //�رն���
	m_pConnection->Close();
	m_pRecordset.Release();                                        //�ͷŶ���
	m_pConnection.Release();
}

void CProjectXDlg::InitExcleType( int iType,BOOL bShowtip /*= TRUE*/ )
{
	//����FlexGrid 
	int iArrayCnt = GetExcleCnt(iType);//��ȡ����
	
	m_FlexGrid.Clear();
	m_FlexGrid.SetCols(iArrayCnt+1);                                        //����FlexGridΪ9��
	
	m_FlexGrid.SetRows(2);
	
	m_FlexGrid.SetBackColorFixed(RGB(50,120, 180));               //���ù̶��к��е���ɫ
	
	//m_FlexGrid.SetBackColor(RGB(170,230,255));                  //���ñ���ɫ
	//m_FlexGrid.SetForeColor(RGB(0,0,0));                        //����ǰ��ɫ
	SetExcleColWith(iType); //�����п�
	
	m_FlexGrid.SetAllowUserResizing(3);                           //����ͨ����������иߺ��п�
	
	
	for(int k = 0; k < iArrayCnt; k++)
	{	
		m_FlexGrid.SetRow(0);                                      //���õ�0��
		m_FlexGrid.SetCol(k+1);                                  //�ӵ�1�п�ʼ����Ԫ��ֵ
		m_FlexGrid.SetCellAlignment(4);                            //���õ�Ԫ�������ʾ
		
		switch(iType)
		{
		case Table_DeciceType :
			m_FlexGrid.SetText(g_strTableDeciceType[k]);                       //���ַ��������ֵ������Ԫ��
			break;
		case Table_Param:
			m_FlexGrid.SetText(g_strTableParam[k]);                       //���ַ��������ֵ������Ԫ��
			break;
		case Table_Coeft:
			m_FlexGrid.SetText(g_strTableCoeft[k]);                       //���ַ��������ֵ������Ԫ��
			break;
		case Table_Cmd:
			m_FlexGrid.SetText(g_strTableCmd[k]);                       //���ַ��������ֵ������Ԫ��
			break;
		case Table_LogType:
			m_FlexGrid.SetText(g_strTableLogType[k]);                       //���ַ��������ֵ������Ԫ��
			break;
		case Table_Log:
			m_FlexGrid.SetText(g_strTableLog[k]);                       //���ַ��������ֵ������Ԫ��
			break;
		case Table_MainParam:
			m_FlexGrid.SetText(g_strTableMainParam[k]);                       //���ַ��������ֵ������Ԫ��
			break;
		case Table_MainSwith:
			m_FlexGrid.SetText(g_strTableMainSwith[k]);                       //���ַ��������ֵ������Ԫ��
			break;
		case Table_WaveParam:
			m_FlexGrid.SetText(g_strTableWaveParam[k]);                       //���ַ��������ֵ������Ԫ��
			break;
		default:
			m_FlexGrid.SetText(g_strTableDeciceType[k]);                       //���ַ��������ֵ������Ԫ��
			break;
		}
	}
	
	LoadInfoFormDataBase(iType,bShowtip);
}

int CProjectXDlg::GetExcleCnt( int iType )
{
	int iArrayCnt = TABLE_DEVICE_CNT;
	switch(iType)
	{
	case Table_DeciceType :
		iArrayCnt = TABLE_DEVICE_CNT;                     
		break;
	case Table_Param:
        iArrayCnt = TABLE_PARAM_CNT;          
		break;
	case Table_Coeft:
		iArrayCnt = TABLE_COEFT_CNT;
		break;
	case Table_Cmd:
        iArrayCnt = TABLE_CMD_CNT;               
		break;
	case Table_LogType:
		iArrayCnt = TABLE_LOGTYPE_CNT;            
		break;
	case Table_Log:
        iArrayCnt = TABLE_LOG_CNT;         
		break;
	case Table_MainParam:
        iArrayCnt = TABLE_MAINPARAM_CNT;        
		break;
	case Table_MainSwith:
        iArrayCnt = TABLE_MAINSWITH_CNT;           
		break;
	case Table_WaveParam:
        iArrayCnt = TABLE_WAVEPARAM_CNT;            
		break;
	default:
		iArrayCnt = TABLE_DEVICE_CNT;  
		break;
	}
	return iArrayCnt;
}

void CProjectXDlg::SetExcleColWith( int iType )
{
	int iArrayCnt = GetExcleCnt(iType);
	m_FlexGrid.SetColWidth(0, 400);
	switch(iType)
	{
	case Table_DeciceType:
		{
			for (int i=1;i<iArrayCnt;i++)
			{
				m_FlexGrid.SetColWidth(i, 2000);
			}
			m_FlexGrid.SetColWidth(iArrayCnt, 3500);
		}
		break;
	default:
		{
			for (int i=1;i<iArrayCnt+1;i++)
			{
				m_FlexGrid.SetColWidth(i, 2000);
			}
		}
		break;
	}
}

void CProjectXDlg::LoadInfoFormDataBase( int iType,BOOL bFresh /*= FALSE*/ )
{
	int iArrayCnt = GetExcleCnt(iType);

		//�������ݿ�                                                         
	CString strConnection;
	strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="+strDBFile;
	m_pConnection.CreateInstance(__uuidof(Connection));
	m_pConnection->CursorLocation = adUseClient;                 
	m_pRecordset.CreateInstance(__uuidof(Recordset)); 
	m_pConnection->Open((LPCTSTR)strConnection, "", "", adModeUnknown);
	
	//�����ݱ�
	CString strText;
	strText = GetSelectSentence(iType);

	m_pRecordset->PutCursorLocation(adUseClient);                 
	m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	
	//��ֹ��˸
	m_FlexGrid.SetRedraw(FALSE);
	int Row = 1;
	int iFileCNt = m_pRecordset->GetRecordCount();
	while(!m_pRecordset->EndOfFile)                                //���û�е���¼�����
	{
		CString strValue;
		_variant_t vstr;
		m_FlexGrid.SetRows(Row + 1);                               //���ü�¼��������
		
		for(int i = 0; i < iArrayCnt; i++)                                 //Ϊÿ�е�Ԫ��ֵ
		{
			int externId = 1;
			if (iType == Table_LogType || iType == Table_Log|| iType == Table_Param)
			{
				externId = 0;
			}
			vstr = m_pRecordset->GetCollect(_variant_t(long(i+externId)));
			if(vstr.vt!=VT_NULL)
				strValue = (LPCSTR)_bstr_t(vstr);
			else
				strValue = "";
			strValue.TrimLeft();
			strValue.TrimRight();
			m_FlexGrid.SetTextMatrix(Row, i+1, strValue);        //ͨ������SetTextMatrix()����Ԫ��ֵ
		}
		//�����и�
		m_FlexGrid.SetRowHeight(0, 320);                                         
		m_FlexGrid.SetRowHeight(Row,280);
		
		m_pRecordset->MoveNext();                                  //����һ��
		Row++;                                                     //�����Լ�1
	}
	
	m_pRecordset->Close();                                         //�رն���
	m_pConnection->Close();
	m_pRecordset.Release();                                        //�ͷŶ���
	m_pConnection.Release();
	
	//������ʾ��ʾ
	//�ڱ���������ʾ���м�¼����
	int counts=m_FlexGrid.GetRows();
	for(int n = 1; n < m_FlexGrid.GetRows(); n++)
	{
		m_FlexGrid.SetRow(n);
		
		for(int m = 1; m < m_FlexGrid.GetCols(); m++)
		{
            m_FlexGrid.SetCol(m);
			m_FlexGrid.SetCellAlignment(1);
		}
		ProgressShow(ProgressType_LOOK,g_strExcleTitle[iType],n*100/counts);
	}

	ProgressShow(ProgressType_LOOK,g_strExcleTitle[iType],100);
	//��ֹ��˸
	m_FlexGrid.SetRedraw(TRUE);

	CString str;
	str.Format("%s(ͳ�ƣ�%d)",g_strExcleTitle[iType],iFileCNt);	
	this->SetWindowText(str);
}

CString CProjectXDlg::GetSelectSentence( int iType )
{
	CString str;
	switch(iType)
	{
	case Table_DeciceType:
		str.Format("select * from %s order by ���",g_strExcleTitle[iType]);
		break;
	case Table_LogType:
		str.Format("select * from %s order by ����",g_strExcleTitle[iType]);
		break;
	case Table_Log:
		str.Format("select * from %s_%s order by ���� ASC,������ ASC",g_strExcleTitle[iType],m_StDeviceSelectNow.cDeviceName);
		break;
	case Table_Param:
	case Table_Cmd:
	case Table_Coeft:
	case Table_MainParam:
	case Table_MainSwith:
	case Table_WaveParam:
		str.Format("select * from %s_%s order by ���",g_strExcleTitle[iType],m_StDeviceSelectNow.cDeviceName);
		break;
	default:
		str.Format("select * from %s order by ���",g_strExcleTitle[iType]);
		break;
	}
	return str;
}
CString CProjectXDlg::GetDeleteSentence(int iType)
{
	CString str;
	switch(iType)
	{
	case Table_DeciceType:
		str.Format("delete * from %s",g_strExcleTitle[iType]);
		break;
	case Table_LogType:
		str.Format("delete * from %s",g_strExcleTitle[iType]);
		break;
	case Table_Param:
	case Table_Cmd:
	case Table_Coeft:
	case Table_Log:
	case Table_MainParam:
	case Table_MainSwith:
	case Table_WaveParam:
		str.Format("delete * from %s_%s",g_strExcleTitle[iType],m_StDeviceSelectNow.cDeviceName);
		break;
	default:
		str.Format("delete * from %s",g_strExcleTitle[iType]);
		break;
	}
	return str;
}
void CProjectXDlg::ProgressShow( int oPera,CString strName,int iVal )
{
	((CProgressCtrl *)GetDlgItem(IDC_PROGRESS_LOAD))->SetPos(iVal);
	CString strTip;
	switch(oPera)
	{
	case ProgressType_LOOK:
		strTip.Format("�鿴%s����:%d%%",strName,iVal);
		break;
	case ProgressType_IN:
		strTip.Format("����%s����:%d%%",strName,iVal);
		break;
	case ProgressType_OUT:
		strTip.Format("����%s����:%d%%",strName,iVal);
		break;
	default:
		strTip.Format("%s����:%d%%",strName,iVal);
		break;
	}
	CString strTip1;
	strTip1.Format("%s     ��ǰѡ���豸����:%s",strTip,m_StDeviceSelectNow.cDeviceName);
	GetDlgItem(IDC_STATIC_1)->SetWindowText(strTip1); 
	RECT stRect;   
    // ��ȡ�ؼ�λ��   
    GetDlgItem(IDC_STATIC_1)->GetWindowRect(&stRect);   
    // ��Ҫ�����ø����ڵ�S2C������������ת��   
    ScreenToClient(&stRect);   
    // �ػ�ؼ����������������������   
	InvalidateRect(&stRect, true);   
	UpdateWindow();
}

HBRUSH CProjectXDlg::OnCtlColor( CDC* pDC, CWnd* pWnd, UINT nCtlColor )
{
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);
	
	// TODO: Change any attributes of the DC here
	if ( pWnd->GetDlgCtrlID()==IDC_STATIC_1)//��������ؼ�ID
		
	{
		pDC->SetTextColor(RGB(255,0,255));
		pDC->SetBkMode(TRANSPARENT);
		
		pDC->SetBkColor(RGB(255,255,255));
		
		return (HBRUSH)::GetStockObject(NULL_BRUSH);//�˴�NULL_BRUSH����Ϊ͸����ˢ�����ǳ��˻�������������͸��������ĳ�WHITE_BRUSH�򻬿鱳�����Ǻڵģ����ǲ�͸��
	}
	// TODO: Return a different brush if the default is not desired
	return hbr;
}

void CProjectXDlg::OnShowParam() 
{
	// TODO: Add your command handler code here
	if(m_iExcleType!=Table_Param)
	{
		m_FlexGrid.Clear();
		m_iExcleType = Table_Param;
		InitExcleType(m_iExcleType);
	}
}
void CProjectXDlg::OnShowCoeft()
{
	// TODO: �ڴ���������������
	if(m_iExcleType!=Table_Coeft)
	{
		m_FlexGrid.Clear();
		m_iExcleType = Table_Coeft;
		InitExcleType(m_iExcleType);
	}
}

void CProjectXDlg::OnShowCmd() 
{
	// TODO: Add your command handler code here
	if(m_iExcleType!=Table_Cmd)
	{
		m_FlexGrid.Clear();
		m_iExcleType = Table_Cmd;
		InitExcleType(m_iExcleType);
	}
}

void CProjectXDlg::OnShowLogtype() 
{
	// TODO: Add your command handler code here
	if(m_iExcleType!=Table_LogType)
	{
		m_FlexGrid.Clear();
		m_iExcleType = Table_LogType;
		InitExcleType(m_iExcleType);
	}
}

void CProjectXDlg::OnShowLog() 
{
	// TODO: Add your command handler code here
	if(m_iExcleType!=Table_Log)
	{
		m_FlexGrid.Clear();
		m_iExcleType = Table_Log;
		InitExcleType(m_iExcleType);
	}
}

void CProjectXDlg::OnShowMainparam() 
{
	// TODO: Add your command handler code here
	if(m_iExcleType!=Table_MainParam)
	{
		m_FlexGrid.Clear();
		m_iExcleType = Table_MainParam;
		InitExcleType(m_iExcleType);
	}
}

void CProjectXDlg::OnShowMainswith() 
{
	// TODO: Add your command handler code here
	if(m_iExcleType!=Table_MainSwith)
	{
		m_FlexGrid.Clear();
		m_iExcleType = Table_MainSwith;
		InitExcleType(m_iExcleType);
	}
}

void CProjectXDlg::OnShowWaveparam() 
{
	// TODO: Add your command handler code here
	if(m_iExcleType!=Table_WaveParam)
	{
		m_FlexGrid.Clear();
		m_iExcleType = Table_WaveParam;
		InitExcleType(m_iExcleType);
	}
}

void CProjectXDlg::OnSelectDevice() 
{
	// TODO: Add your command handler code here
	CDlgDeviceSelect dlg;
	dlg.InitDevice(m_DeviceTypeList,m_StDeviceSelectNow.iID-1);
	if (IDOK==dlg.DoModal())
	{
		if(m_StDeviceSelectNow.iID!=dlg.m_iSelect)
		{
			memcpy(&m_StDeviceSelectNow,&m_DeviceTypeList[dlg.m_iSelect],sizeof(ST_DeviceType));
			m_FlexGrid.Clear();
			m_iExcleType = Table_Param;
			InitExcleType(m_iExcleType);
		}
	}
}

void CProjectXDlg::OnShowDevice() 
{
	// TODO: Add your command handler code here
	if(m_iExcleType!=Table_DeciceType)
	{
		m_FlexGrid.Clear();
		m_iExcleType = Table_DeciceType;
		InitExcleType(m_iExcleType);
	}
}

void CProjectXDlg::OnLoadinNow() 
{
	// TODO: Add your command handler code here
	LoadInExcle();
}

void CProjectXDlg::OnExportoutNow() 
{
	// TODO: Add your command handler code here
	ExportExcle();
}

void CProjectXDlg::OnExportAll() 
{
	// TODO: Add your command handler code here
	CString strPath = SHBrowseForFolder_DirectoryDlg();
	if (strPath == "")
		return;
	
	CString strPathDevice;
	strPathDevice.Format("%s\\%s",strPath,g_strExcleTitle[Table_DeciceType]);
	ExportDeviceExcle(strPathDevice);

	int i = 0;
	for (i=0;i<(int)m_DeviceTypeList.size();i++)
	{
		memcpy(&m_StDeviceSelectNow,&m_DeviceTypeList[i],sizeof(ST_DeviceType));
		CString strPathName;
		strPathName.Format("%s\\%s",strPath,m_StDeviceSelectNow.cDeviceName);
		BOOL bRet = ExportExcle(TRUE,strPathName);
		if (!bRet)
			break;
	}
	if (i == (int)m_DeviceTypeList.size())
	{
		MessageBox("���б�񵼳��ɹ�", "��ʾ",MB_ICONINFORMATION);
	}
}

void CProjectXDlg::OnExportDevicetype() 
{
	// TODO: Add your command handler code here

	ExportDeviceExcle();
}

void CProjectXDlg::OnLoadinDevicetype() 
{
	// TODO: Add your command handler code here
	LoadInDeviceExcle();
}

BOOL CProjectXDlg::ExportExcle( BOOL bTip /*= FALSE*/ ,CString strPath /*= ""*/)
{
	CString strExcleName;
	CString strfileName;
	
	if(strPath == "")
	{
		CString strFileMR;
		strFileMR.Format("%s.xlsx",m_StDeviceSelectNow.cDeviceName);
		CFileDialog dlg (FALSE, NULL, strFileMR/*"���ݿ�����.xlsx"*/,OFN_FILEMUSTEXIST, "Excel File(*.xlsx)|*.xlsx||", this);
		
		int structsize=0; 
		DWORD dwVersion,dwWindowsMajorVersion,dwWindowsMinorVersion; 
		dwVersion = GetVersion(); 
		dwWindowsMajorVersion = (DWORD)(LOBYTE(LOWORD(dwVersion))); 
		dwWindowsMinorVersion = (DWORD)(HIBYTE(LOWORD(dwVersion))); 
		
		if (dwVersion < 0x80000000)               
			structsize =88;                        
		else                                     
			structsize =76;                           
		
		dlg.m_ofn.lStructSize=structsize;
		dlg.m_ofn.lpstrInitialDir = strExcleFile;
		if(dlg.DoModal()==IDOK)
		{
			strExcleName = dlg.GetPathName();
			strfileName = dlg.GetFileName();
			UpdateData(FALSE);
		}
		else
		{
			return FALSE;
		}
	}
	else
	{
		strExcleName = strPath;
	}

	_Application app;                                               
    Workbooks books;                                               
    _Workbook book;                                                 
    Worksheets sheets;                                              
    _Worksheet sheet;                                               
    Range range;                                                    //��Ԫ��Χ
   // Font font;                                                      
    Range cols;
                         
    COleVariant covFalse((short)FALSE);
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

    if( !app.CreateDispatch("Excel.Application") )
    {
		 this->MessageBox("�޷�����ExcelӦ�ã�");
		 return FALSE;
    }

	app.SetVisible(FALSE);  
	
    books=app.GetWorkbooks();

    book=books.Add(covOptional);                                    //�½�������
	CString strbook = book.GetName();
    sheets=book.GetSheets();
	IDispatch* pdisp = sheets.Add(vtMissing,vtMissing , COleVariant((TABLE_TYPE_CNT -1 - 3L)), vtMissing);

	for (int m=1;m<TABLE_TYPE_CNT;m++)
	{	
		sheet=sheets.GetItem(COleVariant((short)(m/*+1*/))); 
		CString strName;
		strName.Format("%s",g_strExcleTitle[m]);
		//strName.Format("%s_%s",g_strExcleTitle[m],m_StDeviceSelectNow.cDeviceName);
		sheet.SetName(strName);
		range.AttachDispatch(sheet.GetCells(),TRUE);//�������е�Ԫ��   
		range.SetNumberFormatLocal(COleVariant("@"));
		range=sheet.GetRange(COleVariant("A1"),COleVariant("A1"));      //�����ʼλ��
		//������ͷ
		int iArrayCnt = GetExcleCnt(m);//��ȡ����
		for(int n=0;n<iArrayCnt;n++)
		{
			switch(m)
			{
			case Table_Param:
				range.SetItem(_variant_t((long)(1)),_variant_t((long)(n+1)),_variant_t(g_strTableParam[n]));
				break;
			case Table_Cmd:
				range.SetItem(_variant_t((long)(1)),_variant_t((long)(n+1)),_variant_t(g_strTableCmd[n]));
				break;
			case Table_Coeft:
				range.SetItem(_variant_t((long)(1)),_variant_t((long)(n+1)),_variant_t(g_strTableCoeft[n]));
				break;
			case Table_LogType:
				range.SetItem(_variant_t((long)(1)),_variant_t((long)(n+1)),_variant_t(g_strTableLogType[n]));
				break;
			case Table_Log:
				range.SetItem(_variant_t((long)(1)),_variant_t((long)(n+1)),_variant_t(g_strTableLog[n]));
				break;
			case Table_MainParam:
				range.SetItem(_variant_t((long)(1)),_variant_t((long)(n+1)),_variant_t(g_strTableMainParam[n]));
				break;
			case Table_MainSwith:
				range.SetItem(_variant_t((long)(1)),_variant_t((long)(n+1)),_variant_t(g_strTableMainSwith[n]));
				break;
			case Table_WaveParam:
				range.SetItem(_variant_t((long)(1)),_variant_t((long)(n+1)),_variant_t(g_strTableWaveParam[n]));
				break;
			default:
				break;
			}
		}
		//�������ݿ�                                                         
		CString strConnection;
		strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="+strDBFile;
		m_pConnection.CreateInstance(__uuidof(Connection));
		m_pConnection->CursorLocation = adUseClient;                
		m_pRecordset.CreateInstance(__uuidof(Recordset)); 
		m_pConnection->Open((LPCTSTR)strConnection, "", "", adModeUnknown);
		

		//�����ݱ�
		CString strText;
		strText = GetSelectSentence(m);
		m_pRecordset->PutCursorLocation(adUseClient);                 
		m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		
		
		int nRow = 1;
		int iFileCNt = m_pRecordset->GetRecordCount();
		while(!m_pRecordset->EndOfFile)                                //���û�е���¼�����
		{
			CString strValue;
			_variant_t vstr;	
			for(int i = 0; i < iArrayCnt; i++)                                 //Ϊÿ�е�Ԫ��ֵ
			{
				int externId = 1;
				if (m == Table_LogType || m == Table_Log|| m == Table_Param)
				{
					externId = 0;
				}
				vstr = m_pRecordset->GetCollect(_variant_t(long(i+externId)));
				if(vstr.vt!=VT_NULL)
					strValue = (LPCSTR)_bstr_t(vstr);
				else
					strValue = ""; 
				range.SetItem(_variant_t((long)(nRow+1)),_variant_t((long)(i+1)),_variant_t(strValue));  
			}
			m_pRecordset->MoveNext();                                  //����һ��
			nRow++; 
			//�����Լ�1
			ProgressShow(ProgressType_OUT,g_strExcleTitle[m],(nRow-1)*100/iFileCNt);
		}
		ProgressShow(ProgressType_OUT,g_strExcleTitle[m],100);
		m_pRecordset->Close();                                         //�رն���
		m_pConnection->Close();
		m_pRecordset.Release();                                        //�ͷŶ���
		m_pConnection.Release();
	}
    
	sheet=sheets.GetItem(COleVariant((short)(1))); 

	if(bTip)
		app.SetDisplayAlerts(FALSE);

	sheet.SaveAs(strExcleName,
		vtMissing,vtMissing,vtMissing,vtMissing,
		vtMissing,vtMissing,
		vtMissing,vtMissing,vtMissing);
	if(bTip)
		app.SetDisplayAlerts(TRUE);
	
	//�ͷŶ��� 
	range.ReleaseDispatch(); 
	// 	if (pdisp)
	// 		pdisp->Release();
	book.SetSaved(TRUE); 
	books.Close(); 
	sheet.ReleaseDispatch(); 
	sheets.ReleaseDispatch(); 
	book.ReleaseDispatch(); 
	books.ReleaseDispatch(); 
	app.ReleaseDispatch(); 	
	books.Close();
	app.Quit();
	CString strTip;
	strTip.Format("����%s�ɹ���",strfileName);
	if(!bTip)
		MessageBox(strTip, "��ʾ",MB_ICONINFORMATION);
	return TRUE;
}

int   CALLBACK   BrowserCallbackProc  
(  
 HWND   hWnd,  
 UINT   uMsg,  
 LPARAM   lParam,  
 LPARAM   lpData  
 )  
{  
	switch   (   uMsg   )  
	{  
	case   BFFM_INITIALIZED:  
		::SendMessage   (   hWnd,   BFFM_SETSELECTION,   1,   lpData   );  
		break;  
	default:  
		break;  
	}  
	return   0;  
  } 

CString CProjectXDlg::SHBrowseForFolder_DirectoryDlg()
{
	char szBuffer[MAX_PATH];
	
	CString szRet = "";
	
	//��ȡ��ǰִ�г����ַ
	char szPath[MAX_PATH];
	
// 	HMODULE hModule = GetModuleHandle(NULL);
// 	
// 	if (hModule)
// 	{
// 		GetModuleFileName(hModule,szPath,MAX_PATH);
// 		PathRemoveFileSpec(szPath);
// 		
// 	}
	strcpy(szPath,strExcleFile);
	
	
	
	BROWSEINFO bi;
	//��ʼ����ڲ���bi��ʼ
	bi.hwndOwner = AfxGetMainWnd() ->GetSafeHwnd();
	bi.pidlRoot = NULL;
	bi.pszDisplayName = szBuffer;//�˲�����ΪNULL������ʾ�Ի���
	bi.lpszTitle = "����";
	bi.ulFlags = 0;
	bi.lpfn = BrowserCallbackProc;
	bi.lParam = (LPARAM)(LPCTSTR)szPath;
	LPITEMIDLIST lpDList =SHBrowseForFolder(&bi);
	
	   if (lpDList)
	   {
		   SHGetPathFromIDList(lpDList, szBuffer);
		   szRet = szBuffer;
	   }
	   
	   LPMALLOC lpMalloc;
	   if(FAILED(SHGetMalloc(&lpMalloc))) 
		   return szRet;
	   //�ͷ��ڴ�
	   lpMalloc->Free(lpDList);
	   lpMalloc->Release();
	   
	   
	   return szRet;
}

void CProjectXDlg::ExportDeviceExcle( CString strPath /*= ""*/ )
{
	CString strExcleName;
	CString strfileName;
	
	CString strFileMR;
	if(strPath == "")
	{
		strFileMR.Format("%s.xlsx",g_strExcleTitle[Table_DeciceType]);
		CFileDialog dlg (FALSE, NULL,strFileMR/*"���ݿ�����.xlsx"*/,OFN_FILEMUSTEXIST, "Excel File(*.xlsx)|*.xlsx||", this);
		
		int structsize=0; 
		DWORD dwVersion,dwWindowsMajorVersion,dwWindowsMinorVersion; 
		dwVersion = GetVersion(); 
		dwWindowsMajorVersion = (DWORD)(LOBYTE(LOWORD(dwVersion))); 
		dwWindowsMinorVersion = (DWORD)(HIBYTE(LOWORD(dwVersion))); 
		
		if (dwVersion < 0x80000000)               
			structsize =88;                        
		else                                     
			structsize =76;                           
		
		dlg.m_ofn.lStructSize=structsize;
		dlg.m_ofn.lpstrInitialDir = strExcleFile;
		if(dlg.DoModal()==IDOK)
		{
			strExcleName = dlg.GetPathName();
			strfileName = dlg.GetFileName();
			UpdateData(FALSE);
		}
		else
		{
			return;
		}
	}
	else
	{
		strExcleName = strPath;
	}

	_Application app;                                               
    Workbooks books;                                               
    _Workbook book;                                                 
    Worksheets sheets;                                              
    _Worksheet sheet;                                               
    Range range;                                                    //��Ԫ��Χ
    //Font font;                                                      
    Range cols;
                         
    COleVariant covFalse((short)FALSE);
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

    if( !app.CreateDispatch("Excel.Application") )
    {
		this->MessageBox("�޷�����ExcelӦ�ã�");
		return;
    }

	app.SetVisible(FALSE);  
	
    books=app.GetWorkbooks();

    book=books.Add(covOptional);                                    //�½�������
	CString strbook = book.GetName();
    sheets=book.GetSheets();

	sheet=sheets.GetItem(COleVariant((short)1)); 
	sheet.SetName(g_strExcleTitle[Table_DeciceType]);
    
	range.AttachDispatch(sheet.GetCells(),TRUE);//�������е�Ԫ��   
	range.SetNumberFormatLocal(COleVariant("@"));
    range=sheet.GetRange(COleVariant("A1"),COleVariant("A1"));      //�����ʼλ��

	//������ͷ
	int iArrayCnt = GetExcleCnt(Table_DeciceType);//��ȡ����
	for(int n=0;n<iArrayCnt;n++)
	{
		range.SetItem(_variant_t((long)(1)),_variant_t((long)(n+1)),_variant_t(g_strTableDeciceType[n]));
	}
 
	//�������ݿ�                                                         
	CString strConnection;
	strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="+strDBFile;
	m_pConnection.CreateInstance(__uuidof(Connection));
	m_pConnection->CursorLocation = adUseClient;                
	m_pRecordset.CreateInstance(__uuidof(Recordset)); 
	m_pConnection->Open((LPCTSTR)strConnection, "", "", adModeUnknown);
	
	//�����ݱ�

	//�����ݱ�
	CString strText;
	strText = GetSelectSentence(Table_DeciceType);
	m_pRecordset->PutCursorLocation(adUseClient);                 
	m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);


	int nRow = 1;
	int iFileCNt = m_pRecordset->GetRecordCount();
	while(!m_pRecordset->EndOfFile)                                //���û�е���¼�����
	{
		CString strValue;
		_variant_t vstr;	
		for(int i = 0; i < iArrayCnt; i++)                                 //Ϊÿ�е�Ԫ��ֵ
		{
			int externId = 1;
			vstr = m_pRecordset->GetCollect(_variant_t(long(i+externId)));
			if(vstr.vt!=VT_NULL)
				strValue = (LPCSTR)_bstr_t(vstr);
			else
				strValue = ""; 
            range.SetItem(_variant_t((long)(nRow+1)),_variant_t((long)(i+1)),_variant_t(strValue));  
		}
		m_pRecordset->MoveNext();                                  //����һ��
		nRow++;														 //�����Լ�1 
		ProgressShow(ProgressType_OUT,g_strExcleTitle[Table_DeciceType],(nRow-1)*100/iFileCNt);
	}
	ProgressShow(ProgressType_OUT,g_strExcleTitle[Table_DeciceType],100);
	m_pRecordset->Close();                                         //�رն���
	m_pConnection->Close();
	m_pRecordset.Release();                                        //�ͷŶ���
	m_pConnection.Release();

	app.SetDisplayAlerts(FALSE);
	sheet.SaveAs(strExcleName,
		vtMissing,vtMissing,vtMissing,vtMissing,
		vtMissing,vtMissing,
            vtMissing,vtMissing,vtMissing);
	app.SetDisplayAlerts(TRUE);
 
	//�ͷŶ��� 
	range.ReleaseDispatch(); 
// 	if (pdisp)
// 		pdisp->Release();
	book.SetSaved(TRUE); 
	books.Close(); 
	sheet.ReleaseDispatch(); 
	sheets.ReleaseDispatch(); 
	book.ReleaseDispatch(); 
	books.ReleaseDispatch(); 
	app.ReleaseDispatch(); 	
	books.Close();
	app.Quit();
	CString strTip;
	strTip.Format("����%s�ɹ���",strfileName);
	if(strPath == "")
		MessageBox(strTip, "��ʾ",MB_ICONINFORMATION);
}

void CProjectXDlg::LoadInDeviceExcle( CString strPath /*= ""*/ )
{
			//��������
	_Application app;
    _Workbook book;
    _Worksheet sheet;
    Workbooks books;
    Worksheets sheets;
    Range range;
    LPDISPATCH lpDisp;
    COleVariant vResult;                              
    COleVariant covTrue((short)TRUE);
    COleVariant covFalse((short)FALSE);
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);                        
    //COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 

    //����Excel������
    if(!app.CreateDispatch("Excel.Application"))
    {
        AfxMessageBox("�޷�����Excel������!");
        return;
    }

    app.SetVisible(FALSE);         
	
	//��ȡExcel�ļ�·��
	CString filePath;
	CString fileName;
	CString strFileMR;
	if(strPath == "")
	{
		strFileMR.Format("%s.xlsx",g_strExcleTitle[Table_DeciceType]);
		CFileDialog dlg (TRUE, NULL, strFileMR/*"���ݿ�����.xlsx"*/,OFN_FILEMUSTEXIST, "Excel File(*.xlsx)|*.xlsx||", this);
		
		int structsize=0; 
		DWORD dwVersion,dwWindowsMajorVersion,dwWindowsMinorVersion; 
		dwVersion = GetVersion(); 
		dwWindowsMajorVersion = (DWORD)(LOBYTE(LOWORD(dwVersion))); 
		dwWindowsMinorVersion = (DWORD)(HIBYTE(LOWORD(dwVersion))); 
		
		if (dwVersion < 0x80000000)               
			structsize =88;                        
		else                                     
			structsize =76;                           
		
		dlg.m_ofn.lStructSize=structsize;
		dlg.m_ofn.lpstrInitialDir = strExcleFile;
		if(dlg.DoModal()==IDOK)
		{
			filePath = dlg.GetPathName();
			fileName = dlg.GetFileName();
			UpdateData(FALSE);
		}
		else
		{
			return;
		}
	}
	else
	{
		filePath = strPath;
	}

	//��Excel�ļ�
    books.AttachDispatch(app.GetWorkbooks());
    lpDisp = books.Open(filePath, covOptional, covFalse, covOptional, covOptional, covOptional, covOptional, 
         covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional);

    //�õ�Workbook
    book.AttachDispatch(lpDisp);

    //�õ�Worksheets 
    sheets.AttachDispatch(book.GetWorksheets()); 

    //�õ���ǰ��Ծsheet
    //lpDisp=book.GetActiveSheet();
    sheet.AttachDispatch(sheets.GetItem(_variant_t(g_strExcleTitle[Table_DeciceType])));

    //��ȡ�Ѿ�ʹ���������Ϣ
    Range usedRange;
    usedRange.AttachDispatch(sheet.GetUsedRange());
    range.AttachDispatch(usedRange.GetRows());
   
	long iRowNum=range.GetCount();                               //�Ѿ�ʹ�õ�����
    long iStartRow=usedRange.GetRow();                           //��ʹ���������ʼ�У��ӵ�1�п�ʼ

	//���жϣ���ֹ���ҵ��룡
	CString strValue[TABLE_DEVICE_CNT+1];  //����Excel��Ԫ������

	char TextName[TABLE_DEVICE_CNT+1][MAX_PATH]={0};  //����Excel��Ԫ���ͷ
	int iArrayCnt = GetExcleCnt(Table_DeciceType);//��ȡ����
	for (int n=0;n<iArrayCnt+1;n++)
	{
		if(n==0)
			strcpy(TextName[n],"���");
		else
			strcpy(TextName[n],g_strTableDeciceType[n-1]);

		if(n>0&&n<iArrayCnt+1)//��Ų��Ƚ�
		{
			range.AttachDispatch(usedRange.GetItem(COleVariant((long)1), COleVariant((long)(n))).pdispVal);
			vResult = range.GetText();
			strValue[n] =  vResult.bstrVal;
			if(strcmp(strValue[n],TextName[n]) !=0)
			{
				CString strError;
				strError.Format("%s��ʽ����׼��������ѡ��",g_strExcleTitle[Table_DeciceType]);
				MessageBox(strError, "��ʾ",MB_ICONEXCLAMATION);	 
				
				//�ر����е�book���˳�Excel 
				book.Close (covOptional,COleVariant(filePath),covOptional);
				books.Close();      
				app.Quit();
				
				return;
			}
		}
	}


	//�ӵ�2�п�ʼ���룬��ͷ�����롣
	//�������ݿ�                                                         
	CString strConnection;
	strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="+strDBFile;
	m_pConnection.CreateInstance(__uuidof(Connection));
	m_pConnection->CursorLocation = adUseClient;                  
	m_pRecordset.CreateInstance(__uuidof(Recordset)); 
	m_pConnection->Open((LPCTSTR)strConnection, "", "", adModeUnknown);

	//�����ݱ�
	//1.��ɾ����������
	CString strText;
	strText.Format("delete * from %s",g_strExcleTitle[Table_DeciceType]);
	m_pRecordset->PutCursorLocation(adUseClient); 
	m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);

	//2.��������
	strText = GetSelectSentence(Table_DeciceType);
	m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	int RowNumData = 1;


	for (int i=iStartRow+1; i<=iRowNum; i++)                      
	{      
	    //Access����1��
		m_pRecordset->AddNew();
		m_pRecordset->PutCollect(TextName[0],_variant_t(long(RowNumData)));//�������
		for(int j=0;j<iArrayCnt;j++)
		{
			// �õ����е�i+1����Ԫ����ַ���
			range.AttachDispatch(usedRange.GetItem(COleVariant((long)i), COleVariant((long)(j+1))).pdispVal);
			vResult = range.GetText();
			strValue[j] = vResult.bstrVal;
			m_pRecordset->PutCollect(TextName[j+1],_variant_t(strValue[j]));
		}

		RowNumData++;
		m_pRecordset->Update();
		ProgressShow(ProgressType_IN,g_strExcleTitle[Table_DeciceType],i*100/iRowNum);
	}	
	ProgressShow(ProgressType_IN,g_strExcleTitle[Table_DeciceType],100);
	//�ر�����
	m_pRecordset->Close();                                         
	m_pConnection->Close();
	m_pRecordset.Release();                                        
	m_pConnection.Release();

    //�ر����е�book���˳�Excel 
    book.Close (covOptional,COleVariant(filePath),covOptional);
    books.Close();      
    app.Quit();


	//����ˢ�º���
	if(m_iExcleType!=Table_DeciceType)
	{
		m_FlexGrid.Clear();
		m_iExcleType = Table_DeciceType;
		InitExcleType(m_iExcleType,FALSE);
	}
	else
		Refresh();
	InitDeviceTypeList();
	if(strPath == "")
	{
		CString strTip;
		strTip.Format("����%s�ɹ���",fileName);
		MessageBox(strTip, "��ʾ",MB_ICONINFORMATION);
	}
}

BOOL CProjectXDlg::LoadInExcle( CString strPath /*= ""*/ )
{
		//��������
	_Application app;
    _Workbook book;
    _Worksheet sheet;
    Workbooks books;
    Worksheets sheets;
    Range range;
    LPDISPATCH lpDisp;
    COleVariant vResult;                              
    COleVariant covTrue((short)TRUE);
    COleVariant covFalse((short)FALSE);
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);                        
    //COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 

    //����Excel������
    if(!app.CreateDispatch("Excel.Application"))
    {
        AfxMessageBox("�޷�����Excel������!");
        return FALSE;
    }

    app.SetVisible(FALSE);         
	
	//��ȡExcel�ļ�·��
	CString filePath;
	CString fileName;
	CString strFileMR;
	if(strPath=="")
	{
		strFileMR.Format("%s.xlsx",m_StDeviceSelectNow.cDeviceName);
		CFileDialog dlg (TRUE, NULL, strFileMR/*"���ݿ�����.xlsx"*/,OFN_FILEMUSTEXIST, "Excel File(*.xlsx)|*.xlsx||", this);
		
		int structsize=0; 
		DWORD dwVersion,dwWindowsMajorVersion,dwWindowsMinorVersion; 
		dwVersion = GetVersion(); 
		dwWindowsMajorVersion = (DWORD)(LOBYTE(LOWORD(dwVersion))); 
		dwWindowsMinorVersion = (DWORD)(HIBYTE(LOWORD(dwVersion))); 
		
		if (dwVersion < 0x80000000)               
			structsize =88;                        
		else                                     
			structsize =76;                           
		
		dlg.m_ofn.lStructSize=structsize;
		dlg.m_ofn.lpstrInitialDir = strExcleFile;
		if(dlg.DoModal()==IDOK)
		{
			filePath = dlg.GetPathName();
			fileName = dlg.GetFileName();
			UpdateData(FALSE);
		}
		else
		{
			return FALSE;
		}
	}
	else
	{	
		filePath = strPath;
	}
	//��Excel�ļ�
    books.AttachDispatch(app.GetWorkbooks());
    lpDisp = books.Open(filePath, covOptional, covFalse, covOptional, covOptional, covOptional, covOptional, 
         covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional);

    //�õ�Workbook
    book.AttachDispatch(lpDisp);

    //�õ�Worksheets 
    sheets.AttachDispatch(book.GetWorksheets()); 

    //�õ���ǰ��Ծsheet

    //lpDisp=book.GetActiveSheet();
	for (int m=1;m<TABLE_TYPE_CNT;m++)
	{
		sheet.AttachDispatch(sheets.GetItem(_variant_t(g_strExcleTitle[m])));
		//��ȡ�Ѿ�ʹ���������Ϣ
		Range usedRange;
		usedRange.AttachDispatch(sheet.GetUsedRange());
		range.AttachDispatch(usedRange.GetRows());
   
		long iRowNum=range.GetCount();                               //�Ѿ�ʹ�õ�����
		long iStartRow=usedRange.GetRow();                           //��ʹ���������ʼ�У��ӵ�1�п�ʼ

		//���жϣ���ֹ���ҵ��룡
		CString strValue[TABLE_PARAM_CNT];  //����Excel��Ԫ������

		char TextName[TABLE_PARAM_CNT][MAX_PATH]={0};  //����Excel��Ԫ���ͷ
		int iArrayCnt = GetExcleCnt(m);//��ȡ����
		if (m!=Table_Param&&m!=Table_LogType&&m!=Table_Log)
			iArrayCnt+=1;

		int iCheck = 0;
		for (int n=0;n<iArrayCnt;n++)
		{
			if (n==0&&m!=Table_LogType&&m!=Table_Log&&m!=Table_Param)
			{
				strcpy(TextName[n],"���");
				iCheck+=1;
				continue;
			}
			switch(m)
			{
			case Table_Param:
				strcpy(TextName[n],g_strTableParam[n-1*iCheck]);
				break;
			case Table_Coeft:
				strcpy(TextName[n],g_strTableCoeft[n-1*iCheck]);
				break;
			case Table_Cmd:
				strcpy(TextName[n],g_strTableCmd[n-1*iCheck]);
				break;
			case Table_LogType:
				strcpy(TextName[n],g_strTableLogType[n-1*iCheck]);
				break;
			case Table_Log:
				strcpy(TextName[n],g_strTableLog[n-1*iCheck]);
				break;
			case Table_MainParam:
				strcpy(TextName[n],g_strTableMainParam[n-1*iCheck]);
				break;
			case Table_MainSwith:
				strcpy(TextName[n],g_strTableMainSwith[n-1*iCheck]);
				break;
			case Table_WaveParam:
				strcpy(TextName[n],g_strTableWaveParam[n-1*iCheck]);
				break;
			default:
				strcpy(TextName[n],g_strTableParam[n-1*iCheck]);
				break;
			}
			if(n>=iCheck&&n<iArrayCnt)//��Ų��Ƚ�
			{
				range.AttachDispatch(usedRange.GetItem(COleVariant((long)1), COleVariant((long)(n+1+(-1*iCheck)))).pdispVal);
				vResult = range.GetText();
				strValue[n] =  vResult.bstrVal;
				if(strcmp(strValue[n],TextName[n]) !=0)
				{
					CString strError;
					strError.Format("%s��ʽ����׼��������ѡ��",g_strExcleTitle[m]);
					MessageBox(strError, "��ʾ",MB_ICONEXCLAMATION);	 
					
					//�ر����е�book���˳�Excel 
					book.Close (covOptional,COleVariant(filePath),covOptional);
					books.Close();      
					app.Quit();
					
					return FALSE;
				}
			}
		}


		//�ӵ�2�п�ʼ���룬��ͷ�����롣

		//�������ݿ�                                                         
		CString strConnection;
		strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="+strDBFile;
		m_pConnection.CreateInstance(__uuidof(Connection));
		m_pConnection->CursorLocation = adUseClient;                  
		m_pRecordset.CreateInstance(__uuidof(Recordset)); 
		m_pConnection->Open((LPCTSTR)strConnection, "", "", adModeUnknown);

		//�����ݱ�
		//1.��ɾ����������
		CString strText;
		strText = GetDeleteSentence(m);
		m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);

		//2.��������
		strText = GetSelectSentence(m);
		m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		int RowNumData = 1;
		if (m==Table_Param)
		{
			strcpy(TextName[0],"���");
			strcpy(TextName[6],"λ��");
		}
		for (int i=iStartRow+1; i<=iRowNum; i++)                      
		{      
			//Access����1��
			m_pRecordset->AddNew();
			for(int j=0;j<iArrayCnt;j++)
			{
				// �õ����е�i+1����Ԫ����ַ���
				if (j==0&&m!=Table_Param&&m!=Table_LogType&&m!=Table_Log)
				{
					m_pRecordset->PutCollect(TextName[0],_variant_t(long(RowNumData)));//�������
					continue;
				}
				range.AttachDispatch(usedRange.GetItem(COleVariant((long)i), COleVariant((long)(j+(iCheck==0?1:0)))).pdispVal);
				vResult = range.GetText();
				strValue[j] = vResult.bstrVal;
				m_pRecordset->PutCollect(TextName[j],_variant_t(strValue[j]));
			}
			RowNumData++;
			m_pRecordset->Update();
			ProgressShow(ProgressType_IN,g_strExcleTitle[m],i*100/iRowNum);
		}	
		ProgressShow(ProgressType_IN,g_strExcleTitle[m],100);
		//�ر�����
		m_pRecordset->Close();                                         
		m_pConnection->Close();
		m_pRecordset.Release();                                        
		m_pConnection.Release();
	}
    //�ر����е�book���˳�Excel 
    book.Close (covOptional,COleVariant(filePath),covOptional);
    books.Close();      
    app.Quit();
	

	if(strPath=="")
	{
		//����ˢ�º���
		if(m_iExcleType!=Table_Param)
		{
			m_FlexGrid.Clear();
			m_iExcleType = Table_Param;
			InitExcleType(m_iExcleType,FALSE);
		}
		else
		Refresh();

		CString strTip;
		strTip.Format("����%s�ɹ���",fileName);
		MessageBox(strTip, "��ʾ",MB_ICONINFORMATION);
	}
	return TRUE;
}

void CProjectXDlg::OnLoadinAll() 
{
	// TODO: Add your command handler code here
	CString strPath = SHBrowseForFolder_DirectoryDlg();
	if (strPath == "")
		return;
	CString strPathDevice;
	strPathDevice.Format("%s\\%s.xlsx",strPath,g_strExcleTitle[Table_DeciceType]);
	if (PathFileExists(strPathDevice))
		LoadInDeviceExcle(strPathDevice);

	int i = 0;
	for (i=0;i<(int)m_DeviceTypeList.size();i++)
	{
		memcpy(&m_StDeviceSelectNow,&m_DeviceTypeList[i],sizeof(ST_DeviceType));
		CString strPathName;
		strPathName.Format("%s\\%s.xlsx",strPath,m_StDeviceSelectNow.cDeviceName);
		if (PathFileExists(strPathName))
		{
			BOOL bRet = LoadInExcle(strPathName);
			if (!bRet)
			break;
		}
	}
	if (i == (int)m_DeviceTypeList.size())
	{
		MessageBox("���б����ɹ�", "��ʾ",MB_ICONINFORMATION);
	}
}
