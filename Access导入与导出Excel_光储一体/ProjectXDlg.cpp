

//作者：袁瑞     QQ:保密
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 1、生成默认的基本对话框工程，工程名为ProjectX。
// 2、在stdafx.h中添加导入ADO库
// 3、在ProjectXDlg.h中添加变量（智能指针变量和标记记录集数量的变量）
// 4、在ProjectXDlg.cpp中添加初始化代码等。
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 添加MSFlexGrid控件
// 1、Ctrl+W打开类向导，工程->增加到工程->Components and Contols->Registered ActiveX Controls->Microsoft FlexGrid Control ,version6.0 ->Insert
// 2、为IDC_DATAGRID1关联变量m_FlexGrid1
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 添加Excel
// 1、Ctrl+W打开类向导，新建一个类，选择从Type Library添加。如果是Office 2003，添加的是Office安装路径下的Excel.exe (在Office 2000环境下添加的应该是Excel9.OLB) 。 
//    在弹出的Confirm Classes里选择_Application，Workbooks，_Workbook，Worksheets ，_Worksheet，Range ，Font 这几个类，
//    并确定新生成的.CPP和.h文件的名称为Excel.cpp和Excel.h，然后确定。
// 2、在ProjectXDlg.cpp中添加头文件引用：#include "Excel.h"
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 重要提示：
// 1、有时候出现内存为写的错误，解决方法：组建->全部重建
// 2、导入的Excel文件必须是设置好的标准的通讯录模板
// 3、导入实际上就是插入，如果想继续添加数据，只要把已有的数据从Excel表中删除即可。
// 4、Access数据库密码为:111111
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// ProjectXDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ProjectX.h"
#include "ProjectXDlg.h"
#include "DlgDeviceSelect.h"
#include <shlwapi.h>
#pragma comment(lib,"Shlwapi.lib")

#include "excel.h"              //操作Excel 
using namespace excel9;

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//==定义======================================================B
#define TABLE_TYPE_CNT 9
#define TABLE_DEVICE_CNT 6 //数目都是不包含排序列,比如说序号
#define TABLE_PARAM_CNT 10
#define TABLE_COEFT_CNT 3
#define TABLE_CMD_CNT 5 
#define	TABLE_LOGTYPE_CNT 2
#define	TABLE_LOG_CNT 4 //先是子类型排序ASC，然后类型排序ASC(DESC降序)
#define TABLE_MAINPARAM_CNT 2
#define TABLE_MAINSWITH_CNT 2
#define TABLE_WAVEPARAM_CNT 5

#define CLIENT_DATABASE_DEF 2	//0是photovol.dll，1是photovol_测试.dll，2是photovol_客户专用.dll
enum TableType
{
	Table_DeciceType = 0,//设备类型
	Table_Param,//参数信息
	Table_Coeft,//变比系数
	Table_Cmd,//命令地址
	Table_LogType,//运行日志类型
	Table_Log,//日志
	Table_MainParam,//主界面参数信息
	Table_MainSwith,//主界面开关状态
	Table_WaveParam//波形信息
};

enum ProGressType
{
	ProgressType_LOOK=0,
	ProgressType_IN,
	ProgressType_OUT,
};

CString g_strExcleTitle[TABLE_TYPE_CNT] = {"设备类型","参数信息","变比系数","命令地址","运行日志类型","运行日志","主界面显示参数","主界面开关地址","波形界面通道参数"};//表名

CString g_strTableDeciceType[TABLE_DEVICE_CNT] = {"设备类型名称","设备型号","描述","真实值显示","当前故障显示","电池组正负对地"};//设备类型
CString g_strTableParam[TABLE_PARAM_CNT] = {"地址","通道标称","系统量","备注名","输入到FPGA值等于实际值*变比","输入到FPGA值最大限制","显示位","数据类型","组号","单位"};//参数信息第一行将序号改为地址，位号变为显示位
CString g_strTableCoeft[TABLE_COEFT_CNT] = {"系统量","备注名","地址"};//变比系数
CString g_strTableCmd[TABLE_CMD_CNT] = {"命令地址","命令名称0","命令名称1","命令位号","命令序号"};//命令地址
CString g_strTableLogType[TABLE_LOGTYPE_CNT] = {"类型","类型名称"};//运行日志类型
CString g_strTableLog[TABLE_LOG_CNT] = {"类型","子类型","日志名","寄存器位"};//日志
CString g_strTableMainParam[TABLE_MAINPARAM_CNT] = {"描述","地址"};//主界面参数信息
CString g_strTableMainSwith[TABLE_MAINSWITH_CNT] = {"地址","数值"};//主界面开关状态
CString g_strTableWaveParam[TABLE_WAVEPARAM_CNT] = {"坐标号","通道号","通道值","变比系数地址","描述"};//波形信息
//==定义======================================================E

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

//设置背景色
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
   if( nCtlColor == CTLCOLOR_STATIC)              //实现静态文本的透明显示
	{   
       pDC->SetBkMode(TRANSPARENT);                
	   return   
	   HBRUSH(GetStockObject(HOLLOW_BRUSH));   
	}

	return hbr;
}

//初始化
BOOL CAboutDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// TODO: Add extra initialization here


	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

//屏蔽Esc键
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

void CProjectXDlg::DoDataExchange(CDataExchange* pDX)  //数据交换
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

//初始化
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
	GetModuleFileName(NULL, path, MAX_PATH);        //获取到完整路径如：E:\Tools\qq.exe
	*strrchr(path,'\\') = '\0';
	
    strDBFile = path;
	strExcleFile = strDBFile ;

	if(0==CLIENT_DATABASE_DEF)
		strDBFile += "\\photovol.dll";
	else if(1==CLIENT_DATABASE_DEF)
		strDBFile += "\\photovol_测试.dll";
	else
		strDBFile += "\\photovol_客户专用.dll";

	InitDeviceTypeList();//初始化数据类型
	m_iExcleType = Table_DeciceType;
	if ((int)m_DeviceTypeList.size()!=0)
	{
		memcpy(&m_StDeviceSelectNow,&m_DeviceTypeList[0],sizeof(ST_DeviceType));
		InitExcleType(m_iExcleType);
	}
    ((CProgressCtrl *)GetDlgItem(IDC_PROGRESS_LOAD))->SetPos(0);
    CenterWindow();                 //窗体居中显示
    ShowWindow(SW_MAXIMIZE);        //最大化显示

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
		CRect rect;                                      //设置背景色                          
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

//屏蔽Esc键
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

//注销CDialog::OnOK(),焦点在FlexGrid上时，按Enter键不会退出对话框 
void CProjectXDlg::OnOK()                                        
{
	// TODO: Add extra validation here
	
	//CDialog::OnOK();                   
}

//调整空间大小
void CProjectXDlg::OnSize(UINT nType, int cx, int cy)   //对话框大小
{
	CDialog::OnSize(nType, cx, cy);
	
	// TODO: Add your message handler code here
    CWnd *pWnd; 
    
	//调整FlexGrid1位置和大小
	if(nType==1) return;                            //如果窗体最小化则返回
	pWnd = GetDlgItem(IDC_MSFLEXGRID1);             //获取控件句柄
	if(pWnd)                                        //判断是否为空，因为对话框创建时会调用此函数，而当时控件还未创建
    {
      CRect rect;                                   //变化前大小
      pWnd->GetWindowRect(&rect);
      
	  ScreenToClient(&rect);                        
      
      rect.left=rect.left*cx/m_rect.Width();        //调整横向大小
      rect.right=rect.right*cx/m_rect.Width();    
      rect.top=rect.top*cy/m_rect.Height();         //调整纵向大小
      rect.bottom=rect.bottom*cy/m_rect.Height();
      
	  rect.bottom = cy -24;
	  pWnd->MoveWindow(rect);                       //设置控件大小
	  GetDlgItem(IDC_PROGRESS_LOAD)->MoveWindow(m_rect.left+2,cy-22,600,20);
	  GetDlgItem(IDC_STATIC_1)->MoveWindow(m_rect.left+615,cy-20,500,20);
    }

	GetClientRect(&m_rect);                         //将变化后的对话框设为原始大小	
}

//响应滚轮
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


//刷新函数
void CProjectXDlg:: Refresh()
{
	LoadInfoFormDataBase(m_iExcleType,TRUE);
}   

//退出
void CProjectXDlg::OnExit() 
{
	// TODO: Add your command handler code here
	
	CDialog::OnOK(); 
}

//关于
void CProjectXDlg::OnAbout() 
{
	// TODO: Add your command handler code here
	CAboutDlg dlg;
    dlg.DoModal();
}            

void CProjectXDlg::InitDeviceTypeList()
{
	m_DeviceTypeList.clear();

	//连接数据库                                                         
	CString strConnection;
	strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="+strDBFile;
	m_pConnection.CreateInstance(__uuidof(Connection));
	m_pConnection->CursorLocation = adUseClient;                
	m_pRecordset.CreateInstance(__uuidof(Recordset)); 
	m_pConnection->Open((LPCTSTR)strConnection, "", "", adModeUnknown);
	
	//打开数据表
	
	CString strText;
	strText.Format("select * from %s order by 序号",g_strExcleTitle[Table_DeciceType]);
	m_pRecordset->PutCursorLocation(adUseClient);                 
	m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	
	while(!m_pRecordset->EndOfFile)                                //如果没有到记录集最后
	{
		CString strValue;
		_variant_t vstr;
		ST_DeviceType oParam;
		
		for(int i = 0; i < TABLE_DEVICE_CNT+1; i++)                                 //为每行单元格赋值
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
		m_pRecordset->MoveNext();                                  //下移一行                                                 //行数自加1
	}
	
	m_pRecordset->Close();                                         //关闭对象
	m_pConnection->Close();
	m_pRecordset.Release();                                        //释放对象
	m_pConnection.Release();
}

void CProjectXDlg::InitExcleType( int iType,BOOL bShowtip /*= TRUE*/ )
{
	//设置FlexGrid 
	int iArrayCnt = GetExcleCnt(iType);//获取列数
	
	m_FlexGrid.Clear();
	m_FlexGrid.SetCols(iArrayCnt+1);                                        //设置FlexGrid为9列
	
	m_FlexGrid.SetRows(2);
	
	m_FlexGrid.SetBackColorFixed(RGB(50,120, 180));               //设置固定行和列的颜色
	
	//m_FlexGrid.SetBackColor(RGB(170,230,255));                  //设置背景色
	//m_FlexGrid.SetForeColor(RGB(0,0,0));                        //设置前景色
	SetExcleColWith(iType); //设置列宽
	
	m_FlexGrid.SetAllowUserResizing(3);                           //允许通过鼠标拉动行高和列宽
	
	
	for(int k = 0; k < iArrayCnt; k++)
	{	
		m_FlexGrid.SetRow(0);                                      //设置第0行
		m_FlexGrid.SetCol(k+1);                                  //从第1列开始给单元格赋值
		m_FlexGrid.SetCellAlignment(4);                            //设置单元格居中显示
		
		switch(iType)
		{
		case Table_DeciceType :
			m_FlexGrid.SetText(g_strTableDeciceType[k]);                       //把字符串数组的值赋给单元格
			break;
		case Table_Param:
			m_FlexGrid.SetText(g_strTableParam[k]);                       //把字符串数组的值赋给单元格
			break;
		case Table_Coeft:
			m_FlexGrid.SetText(g_strTableCoeft[k]);                       //把字符串数组的值赋给单元格
			break;
		case Table_Cmd:
			m_FlexGrid.SetText(g_strTableCmd[k]);                       //把字符串数组的值赋给单元格
			break;
		case Table_LogType:
			m_FlexGrid.SetText(g_strTableLogType[k]);                       //把字符串数组的值赋给单元格
			break;
		case Table_Log:
			m_FlexGrid.SetText(g_strTableLog[k]);                       //把字符串数组的值赋给单元格
			break;
		case Table_MainParam:
			m_FlexGrid.SetText(g_strTableMainParam[k]);                       //把字符串数组的值赋给单元格
			break;
		case Table_MainSwith:
			m_FlexGrid.SetText(g_strTableMainSwith[k]);                       //把字符串数组的值赋给单元格
			break;
		case Table_WaveParam:
			m_FlexGrid.SetText(g_strTableWaveParam[k]);                       //把字符串数组的值赋给单元格
			break;
		default:
			m_FlexGrid.SetText(g_strTableDeciceType[k]);                       //把字符串数组的值赋给单元格
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

		//连接数据库                                                         
	CString strConnection;
	strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="+strDBFile;
	m_pConnection.CreateInstance(__uuidof(Connection));
	m_pConnection->CursorLocation = adUseClient;                 
	m_pRecordset.CreateInstance(__uuidof(Recordset)); 
	m_pConnection->Open((LPCTSTR)strConnection, "", "", adModeUnknown);
	
	//打开数据表
	CString strText;
	strText = GetSelectSentence(iType);

	m_pRecordset->PutCursorLocation(adUseClient);                 
	m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	
	//防止闪烁
	m_FlexGrid.SetRedraw(FALSE);
	int Row = 1;
	int iFileCNt = m_pRecordset->GetRecordCount();
	while(!m_pRecordset->EndOfFile)                                //如果没有到记录集最后
	{
		CString strValue;
		_variant_t vstr;
		m_FlexGrid.SetRows(Row + 1);                               //设置记录集的行数
		
		for(int i = 0; i < iArrayCnt; i++)                                 //为每行单元格赋值
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
			m_FlexGrid.SetTextMatrix(Row, i+1, strValue);        //通过函数SetTextMatrix()给单元格赋值
		}
		//设置行高
		m_FlexGrid.SetRowHeight(0, 320);                                         
		m_FlexGrid.SetRowHeight(Row,280);
		
		m_pRecordset->MoveNext();                                  //下移一行
		Row++;                                                     //行数自加1
	}
	
	m_pRecordset->Close();                                         //关闭对象
	m_pConnection->Close();
	m_pRecordset.Release();                                        //释放对象
	m_pConnection.Release();
	
	//居左显示显示
	//在标题栏中显示共有记录条数
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
	//防止闪烁
	m_FlexGrid.SetRedraw(TRUE);

	CString str;
	str.Format("%s(统计：%d)",g_strExcleTitle[iType],iFileCNt);	
	this->SetWindowText(str);
}

CString CProjectXDlg::GetSelectSentence( int iType )
{
	CString str;
	switch(iType)
	{
	case Table_DeciceType:
		str.Format("select * from %s order by 序号",g_strExcleTitle[iType]);
		break;
	case Table_LogType:
		str.Format("select * from %s order by 类型",g_strExcleTitle[iType]);
		break;
	case Table_Log:
		str.Format("select * from %s_%s order by 类型 ASC,子类型 ASC",g_strExcleTitle[iType],m_StDeviceSelectNow.cDeviceName);
		break;
	case Table_Param:
	case Table_Cmd:
	case Table_Coeft:
	case Table_MainParam:
	case Table_MainSwith:
	case Table_WaveParam:
		str.Format("select * from %s_%s order by 序号",g_strExcleTitle[iType],m_StDeviceSelectNow.cDeviceName);
		break;
	default:
		str.Format("select * from %s order by 序号",g_strExcleTitle[iType]);
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
		strTip.Format("查看%s进度:%d%%",strName,iVal);
		break;
	case ProgressType_IN:
		strTip.Format("导入%s进度:%d%%",strName,iVal);
		break;
	case ProgressType_OUT:
		strTip.Format("导出%s进度:%d%%",strName,iVal);
		break;
	default:
		strTip.Format("%s进度:%d%%",strName,iVal);
		break;
	}
	CString strTip1;
	strTip1.Format("%s     当前选择设备类型:%s",strTip,m_StDeviceSelectNow.cDeviceName);
	GetDlgItem(IDC_STATIC_1)->SetWindowText(strTip1); 
	RECT stRect;   
    // 获取控件位置   
    GetDlgItem(IDC_STATIC_1)->GetWindowRect(&stRect);   
    // 重要！调用父窗口的S2C函数进行坐标转换   
    ScreenToClient(&stRect);   
    // 重绘控件所在区域，在这里擦除背景   
	InvalidateRect(&stRect, true);   
	UpdateWindow();
}

HBRUSH CProjectXDlg::OnCtlColor( CDC* pDC, CWnd* pWnd, UINT nCtlColor )
{
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);
	
	// TODO: Change any attributes of the DC here
	if ( pWnd->GetDlgCtrlID()==IDC_STATIC_1)//包括滑块控件ID
		
	{
		pDC->SetTextColor(RGB(255,0,255));
		pDC->SetBkMode(TRANSPARENT);
		
		pDC->SetBkColor(RGB(255,255,255));
		
		return (HBRUSH)::GetStockObject(NULL_BRUSH);//此处NULL_BRUSH作用为透明画刷，但是除了滑块其他都可以透明；如果改成WHITE_BRUSH则滑块背景不是黑的，但是不透明
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
	// TODO: 在此添加命令处理程序代码
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
		MessageBox("所有表格导出成功", "提示",MB_ICONINFORMATION);
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
		CFileDialog dlg (FALSE, NULL, strFileMR/*"数据库配置.xlsx"*/,OFN_FILEMUSTEXIST, "Excel File(*.xlsx)|*.xlsx||", this);
		
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
    Range range;                                                    //单元格范围
   // Font font;                                                      
    Range cols;
                         
    COleVariant covFalse((short)FALSE);
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

    if( !app.CreateDispatch("Excel.Application") )
    {
		 this->MessageBox("无法创建Excel应用！");
		 return FALSE;
    }

	app.SetVisible(FALSE);  
	
    books=app.GetWorkbooks();

    book=books.Add(covOptional);                                    //新建工作簿
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
		range.AttachDispatch(sheet.GetCells(),TRUE);//加载所有单元格   
		range.SetNumberFormatLocal(COleVariant("@"));
		range=sheet.GetRange(COleVariant("A1"),COleVariant("A1"));      //表格起始位置
		//导出表头
		int iArrayCnt = GetExcleCnt(m);//获取列数
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
		//连接数据库                                                         
		CString strConnection;
		strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="+strDBFile;
		m_pConnection.CreateInstance(__uuidof(Connection));
		m_pConnection->CursorLocation = adUseClient;                
		m_pRecordset.CreateInstance(__uuidof(Recordset)); 
		m_pConnection->Open((LPCTSTR)strConnection, "", "", adModeUnknown);
		

		//打开数据表
		CString strText;
		strText = GetSelectSentence(m);
		m_pRecordset->PutCursorLocation(adUseClient);                 
		m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		
		
		int nRow = 1;
		int iFileCNt = m_pRecordset->GetRecordCount();
		while(!m_pRecordset->EndOfFile)                                //如果没有到记录集最后
		{
			CString strValue;
			_variant_t vstr;	
			for(int i = 0; i < iArrayCnt; i++)                                 //为每行单元格赋值
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
			m_pRecordset->MoveNext();                                  //下移一行
			nRow++; 
			//行数自加1
			ProgressShow(ProgressType_OUT,g_strExcleTitle[m],(nRow-1)*100/iFileCNt);
		}
		ProgressShow(ProgressType_OUT,g_strExcleTitle[m],100);
		m_pRecordset->Close();                                         //关闭对象
		m_pConnection->Close();
		m_pRecordset.Release();                                        //释放对象
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
	
	//释放对象 
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
	strTip.Format("导出%s成功！",strfileName);
	if(!bTip)
		MessageBox(strTip, "提示",MB_ICONINFORMATION);
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
	
	//获取当前执行程序地址
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
	//初始化入口参数bi开始
	bi.hwndOwner = AfxGetMainWnd() ->GetSafeHwnd();
	bi.pidlRoot = NULL;
	bi.pszDisplayName = szBuffer;//此参数如为NULL则不能显示对话框
	bi.lpszTitle = "保存";
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
	   //释放内存
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
		CFileDialog dlg (FALSE, NULL,strFileMR/*"数据库配置.xlsx"*/,OFN_FILEMUSTEXIST, "Excel File(*.xlsx)|*.xlsx||", this);
		
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
    Range range;                                                    //单元格范围
    //Font font;                                                      
    Range cols;
                         
    COleVariant covFalse((short)FALSE);
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

    if( !app.CreateDispatch("Excel.Application") )
    {
		this->MessageBox("无法创建Excel应用！");
		return;
    }

	app.SetVisible(FALSE);  
	
    books=app.GetWorkbooks();

    book=books.Add(covOptional);                                    //新建工作簿
	CString strbook = book.GetName();
    sheets=book.GetSheets();

	sheet=sheets.GetItem(COleVariant((short)1)); 
	sheet.SetName(g_strExcleTitle[Table_DeciceType]);
    
	range.AttachDispatch(sheet.GetCells(),TRUE);//加载所有单元格   
	range.SetNumberFormatLocal(COleVariant("@"));
    range=sheet.GetRange(COleVariant("A1"),COleVariant("A1"));      //表格起始位置

	//导出表头
	int iArrayCnt = GetExcleCnt(Table_DeciceType);//获取列数
	for(int n=0;n<iArrayCnt;n++)
	{
		range.SetItem(_variant_t((long)(1)),_variant_t((long)(n+1)),_variant_t(g_strTableDeciceType[n]));
	}
 
	//连接数据库                                                         
	CString strConnection;
	strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="+strDBFile;
	m_pConnection.CreateInstance(__uuidof(Connection));
	m_pConnection->CursorLocation = adUseClient;                
	m_pRecordset.CreateInstance(__uuidof(Recordset)); 
	m_pConnection->Open((LPCTSTR)strConnection, "", "", adModeUnknown);
	
	//打开数据表

	//打开数据表
	CString strText;
	strText = GetSelectSentence(Table_DeciceType);
	m_pRecordset->PutCursorLocation(adUseClient);                 
	m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);


	int nRow = 1;
	int iFileCNt = m_pRecordset->GetRecordCount();
	while(!m_pRecordset->EndOfFile)                                //如果没有到记录集最后
	{
		CString strValue;
		_variant_t vstr;	
		for(int i = 0; i < iArrayCnt; i++)                                 //为每行单元格赋值
		{
			int externId = 1;
			vstr = m_pRecordset->GetCollect(_variant_t(long(i+externId)));
			if(vstr.vt!=VT_NULL)
				strValue = (LPCSTR)_bstr_t(vstr);
			else
				strValue = ""; 
            range.SetItem(_variant_t((long)(nRow+1)),_variant_t((long)(i+1)),_variant_t(strValue));  
		}
		m_pRecordset->MoveNext();                                  //下移一行
		nRow++;														 //行数自加1 
		ProgressShow(ProgressType_OUT,g_strExcleTitle[Table_DeciceType],(nRow-1)*100/iFileCNt);
	}
	ProgressShow(ProgressType_OUT,g_strExcleTitle[Table_DeciceType],100);
	m_pRecordset->Close();                                         //关闭对象
	m_pConnection->Close();
	m_pRecordset.Release();                                        //释放对象
	m_pConnection.Release();

	app.SetDisplayAlerts(FALSE);
	sheet.SaveAs(strExcleName,
		vtMissing,vtMissing,vtMissing,vtMissing,
		vtMissing,vtMissing,
            vtMissing,vtMissing,vtMissing);
	app.SetDisplayAlerts(TRUE);
 
	//释放对象 
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
	strTip.Format("导出%s成功！",strfileName);
	if(strPath == "")
		MessageBox(strTip, "提示",MB_ICONINFORMATION);
}

void CProjectXDlg::LoadInDeviceExcle( CString strPath /*= ""*/ )
{
			//创建对象
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

    //启动Excel服务器
    if(!app.CreateDispatch("Excel.Application"))
    {
        AfxMessageBox("无法启动Excel服务器!");
        return;
    }

    app.SetVisible(FALSE);         
	
	//获取Excel文件路径
	CString filePath;
	CString fileName;
	CString strFileMR;
	if(strPath == "")
	{
		strFileMR.Format("%s.xlsx",g_strExcleTitle[Table_DeciceType]);
		CFileDialog dlg (TRUE, NULL, strFileMR/*"数据库配置.xlsx"*/,OFN_FILEMUSTEXIST, "Excel File(*.xlsx)|*.xlsx||", this);
		
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

	//打开Excel文件
    books.AttachDispatch(app.GetWorkbooks());
    lpDisp = books.Open(filePath, covOptional, covFalse, covOptional, covOptional, covOptional, covOptional, 
         covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional);

    //得到Workbook
    book.AttachDispatch(lpDisp);

    //得到Worksheets 
    sheets.AttachDispatch(book.GetWorksheets()); 

    //得到当前活跃sheet
    //lpDisp=book.GetActiveSheet();
    sheet.AttachDispatch(sheets.GetItem(_variant_t(g_strExcleTitle[Table_DeciceType])));

    //读取已经使用区域的信息
    Range usedRange;
    usedRange.AttachDispatch(sheet.GetUsedRange());
    range.AttachDispatch(usedRange.GetRows());
   
	long iRowNum=range.GetCount();                               //已经使用的行数
    long iStartRow=usedRange.GetRow();                           //已使用区域的起始行，从第1行开始

	//先判断，防止胡乱导入！
	CString strValue[TABLE_DEVICE_CNT+1];  //读入Excel单元格数据

	char TextName[TABLE_DEVICE_CNT+1][MAX_PATH]={0};  //读入Excel单元格表头
	int iArrayCnt = GetExcleCnt(Table_DeciceType);//获取列数
	for (int n=0;n<iArrayCnt+1;n++)
	{
		if(n==0)
			strcpy(TextName[n],"序号");
		else
			strcpy(TextName[n],g_strTableDeciceType[n-1]);

		if(n>0&&n<iArrayCnt+1)//序号不比较
		{
			range.AttachDispatch(usedRange.GetItem(COleVariant((long)1), COleVariant((long)(n))).pdispVal);
			vResult = range.GetText();
			strValue[n] =  vResult.bstrVal;
			if(strcmp(strValue[n],TextName[n]) !=0)
			{
				CString strError;
				strError.Format("%s格式不标准，请重新选择！",g_strExcleTitle[Table_DeciceType]);
				MessageBox(strError, "提示",MB_ICONEXCLAMATION);	 
				
				//关闭所有的book，退出Excel 
				book.Close (covOptional,COleVariant(filePath),covOptional);
				books.Close();      
				app.Quit();
				
				return;
			}
		}
	}


	//从第2行开始导入，表头不导入。
	//连接数据库                                                         
	CString strConnection;
	strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="+strDBFile;
	m_pConnection.CreateInstance(__uuidof(Connection));
	m_pConnection->CursorLocation = adUseClient;                  
	m_pRecordset.CreateInstance(__uuidof(Recordset)); 
	m_pConnection->Open((LPCTSTR)strConnection, "", "", adModeUnknown);

	//打开数据表
	//1.先删除所有数据
	CString strText;
	strText.Format("delete * from %s",g_strExcleTitle[Table_DeciceType]);
	m_pRecordset->PutCursorLocation(adUseClient); 
	m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);

	//2.插入数据
	strText = GetSelectSentence(Table_DeciceType);
	m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	int RowNumData = 1;


	for (int i=iStartRow+1; i<=iRowNum; i++)                      
	{      
	    //Access增加1行
		m_pRecordset->AddNew();
		m_pRecordset->PutCollect(TextName[0],_variant_t(long(RowNumData)));//插入序号
		for(int j=0;j<iArrayCnt;j++)
		{
			// 得到本行第i+1个单元格的字符串
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
	//关闭连接
	m_pRecordset->Close();                                         
	m_pConnection->Close();
	m_pRecordset.Release();                                        
	m_pConnection.Release();

    //关闭所有的book，退出Excel 
    book.Close (covOptional,COleVariant(filePath),covOptional);
    books.Close();      
    app.Quit();


	//调用刷新函数
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
		strTip.Format("导入%s成功！",fileName);
		MessageBox(strTip, "提示",MB_ICONINFORMATION);
	}
}

BOOL CProjectXDlg::LoadInExcle( CString strPath /*= ""*/ )
{
		//创建对象
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

    //启动Excel服务器
    if(!app.CreateDispatch("Excel.Application"))
    {
        AfxMessageBox("无法启动Excel服务器!");
        return FALSE;
    }

    app.SetVisible(FALSE);         
	
	//获取Excel文件路径
	CString filePath;
	CString fileName;
	CString strFileMR;
	if(strPath=="")
	{
		strFileMR.Format("%s.xlsx",m_StDeviceSelectNow.cDeviceName);
		CFileDialog dlg (TRUE, NULL, strFileMR/*"数据库配置.xlsx"*/,OFN_FILEMUSTEXIST, "Excel File(*.xlsx)|*.xlsx||", this);
		
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
	//打开Excel文件
    books.AttachDispatch(app.GetWorkbooks());
    lpDisp = books.Open(filePath, covOptional, covFalse, covOptional, covOptional, covOptional, covOptional, 
         covOptional, covOptional, covOptional, covOptional, covOptional,covOptional, covOptional, covOptional);

    //得到Workbook
    book.AttachDispatch(lpDisp);

    //得到Worksheets 
    sheets.AttachDispatch(book.GetWorksheets()); 

    //得到当前活跃sheet

    //lpDisp=book.GetActiveSheet();
	for (int m=1;m<TABLE_TYPE_CNT;m++)
	{
		sheet.AttachDispatch(sheets.GetItem(_variant_t(g_strExcleTitle[m])));
		//读取已经使用区域的信息
		Range usedRange;
		usedRange.AttachDispatch(sheet.GetUsedRange());
		range.AttachDispatch(usedRange.GetRows());
   
		long iRowNum=range.GetCount();                               //已经使用的行数
		long iStartRow=usedRange.GetRow();                           //已使用区域的起始行，从第1行开始

		//先判断，防止胡乱导入！
		CString strValue[TABLE_PARAM_CNT];  //读入Excel单元格数据

		char TextName[TABLE_PARAM_CNT][MAX_PATH]={0};  //读入Excel单元格表头
		int iArrayCnt = GetExcleCnt(m);//获取列数
		if (m!=Table_Param&&m!=Table_LogType&&m!=Table_Log)
			iArrayCnt+=1;

		int iCheck = 0;
		for (int n=0;n<iArrayCnt;n++)
		{
			if (n==0&&m!=Table_LogType&&m!=Table_Log&&m!=Table_Param)
			{
				strcpy(TextName[n],"序号");
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
			if(n>=iCheck&&n<iArrayCnt)//序号不比较
			{
				range.AttachDispatch(usedRange.GetItem(COleVariant((long)1), COleVariant((long)(n+1+(-1*iCheck)))).pdispVal);
				vResult = range.GetText();
				strValue[n] =  vResult.bstrVal;
				if(strcmp(strValue[n],TextName[n]) !=0)
				{
					CString strError;
					strError.Format("%s格式不标准，请重新选择！",g_strExcleTitle[m]);
					MessageBox(strError, "提示",MB_ICONEXCLAMATION);	 
					
					//关闭所有的book，退出Excel 
					book.Close (covOptional,COleVariant(filePath),covOptional);
					books.Close();      
					app.Quit();
					
					return FALSE;
				}
			}
		}


		//从第2行开始导入，表头不导入。

		//连接数据库                                                         
		CString strConnection;
		strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="+strDBFile;
		m_pConnection.CreateInstance(__uuidof(Connection));
		m_pConnection->CursorLocation = adUseClient;                  
		m_pRecordset.CreateInstance(__uuidof(Recordset)); 
		m_pConnection->Open((LPCTSTR)strConnection, "", "", adModeUnknown);

		//打开数据表
		//1.先删除所有数据
		CString strText;
		strText = GetDeleteSentence(m);
		m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);

		//2.插入数据
		strText = GetSelectSentence(m);
		m_pRecordset->Open(_variant_t(strText), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		int RowNumData = 1;
		if (m==Table_Param)
		{
			strcpy(TextName[0],"序号");
			strcpy(TextName[6],"位号");
		}
		for (int i=iStartRow+1; i<=iRowNum; i++)                      
		{      
			//Access增加1行
			m_pRecordset->AddNew();
			for(int j=0;j<iArrayCnt;j++)
			{
				// 得到本行第i+1个单元格的字符串
				if (j==0&&m!=Table_Param&&m!=Table_LogType&&m!=Table_Log)
				{
					m_pRecordset->PutCollect(TextName[0],_variant_t(long(RowNumData)));//插入序号
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
		//关闭连接
		m_pRecordset->Close();                                         
		m_pConnection->Close();
		m_pRecordset.Release();                                        
		m_pConnection.Release();
	}
    //关闭所有的book，退出Excel 
    book.Close (covOptional,COleVariant(filePath),covOptional);
    books.Close();      
    app.Quit();
	

	if(strPath=="")
	{
		//调用刷新函数
		if(m_iExcleType!=Table_Param)
		{
			m_FlexGrid.Clear();
			m_iExcleType = Table_Param;
			InitExcleType(m_iExcleType,FALSE);
		}
		else
		Refresh();

		CString strTip;
		strTip.Format("导入%s成功！",fileName);
		MessageBox(strTip, "提示",MB_ICONINFORMATION);
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
		MessageBox("所有表格导入成功", "提示",MB_ICONINFORMATION);
	}
}
