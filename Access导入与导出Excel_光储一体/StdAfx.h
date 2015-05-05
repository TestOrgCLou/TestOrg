// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__B4E97C29_DC6D_4866_BCE1_695E93EECA65__INCLUDED_)
#define AFX_STDAFX_H__B4E97C29_DC6D_4866_BCE1_695E93EECA65__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define VC_EXTRALEAN		// Exclude rarely-used stuff from Windows headers

#include <afxcview.h>       //添加

#include <afxwin.h>         // MFC core and standard components
#include <afxext.h>         // MFC extensions
#include <afxdisp.h>        // MFC Automation classes
#include <afxdtctl.h>		// MFC support for Internet Explorer 4 Common Controls
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			// MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT

//添加（打印）
typedef void(*DRAWFUN)(CDC* pDC,CPrintInfo* pInfo,void* pVoid);

//导入Ado库
#import "D:\ado\msado15.dll" no_namespace \
rename("EOF","EndOfFile") \
rename("LockTypeEnum","newLockTypeEnum")\
rename("DataTypeEnum","newDataTypeEnum")\
rename("FieldAttributeEnum","newFieldAttributeEnum")\
rename("EditModeEnum","newEditModeEnum")\
rename("RecordStatusEnum","newRecordStatusEnum")\
rename("ParameterDirectionEnum","newParameterDirectionEnum")

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__B4E97C29_DC6D_4866_BCE1_695E93EECA65__INCLUDED_)
