// DirectoryDlg.h: interface for the CDirectoryDlg class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DIRECTORYDLG_H__AB8FAEAB_D3FE_4BA4_9900_E6CA0449C7B8__INCLUDED_)
#define AFX_DIRECTORYDLG_H__AB8FAEAB_D3FE_4BA4_9900_E6CA0449C7B8__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CDirectoryDlg  
{
public:
	CDirectoryDlg();
	virtual ~CDirectoryDlg();

	//弹出选择文件夹的，保存文件列表
	static CString SHBrowseForFolder_DirectoryDlg();



};

#endif // !defined(AFX_DIRECTORYDLG_H__AB8FAEAB_D3FE_4BA4_9900_E6CA0449C7B8__INCLUDED_)
