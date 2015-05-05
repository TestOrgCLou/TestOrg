// ProjectX.h : main header file for the PROJECTX application
//

#if !defined(AFX_PROJECTX_H__9821987A_9DDD_4A78_8A00_5B7BFBDC5771__INCLUDED_)
#define AFX_PROJECTX_H__9821987A_9DDD_4A78_8A00_5B7BFBDC5771__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CProjectXApp:
// See ProjectX.cpp for the implementation of this class
//

class CProjectXApp : public CWinApp
{
public:
	CProjectXApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CProjectXApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CProjectXApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PROJECTX_H__9821987A_9DDD_4A78_8A00_5B7BFBDC5771__INCLUDED_)
