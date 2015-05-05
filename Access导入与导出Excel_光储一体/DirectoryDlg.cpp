// DirectoryDlg.cpp: implementation of the CDirectoryDlg class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DirectoryDlg.h"

#include <shlwapi.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

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

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CDirectoryDlg::CDirectoryDlg()
{

}

CDirectoryDlg::~CDirectoryDlg()
{

}

CString CDirectoryDlg::SHBrowseForFolder_DirectoryDlg()
{
	 	char szBuffer[MAX_PATH];

		CString szRet = "";

		ITEMIDLIST IDlist;
		LPCITEMIDLIST lpIDList;
		//��ȡ��ǰִ�г����ַ
		char szPath[MAX_PATH];
		
		HMODULE hModule = GetModuleHandle(NULL);
		
		if (hModule)
		{
			GetModuleFileName(hModule,szPath,MAX_PATH);
			PathRemoveFileSpec(szPath);
			
		}
	
		
	 
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
