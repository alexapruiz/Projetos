// Eduardo2Sem.h : main header file for the EDUARDO2SEM application
//

#if !defined(AFX_EDUARDO2SEM_H__B73E9385_6B5A_42C0_81CB_FE0AF756D5BE__INCLUDED_)
#define AFX_EDUARDO2SEM_H__B73E9385_6B5A_42C0_81CB_FE0AF756D5BE__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CEduardo2SemApp:
// See Eduardo2Sem.cpp for the implementation of this class
//

class CEduardo2SemApp : public CWinApp
{
public:
	CEduardo2SemApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CEduardo2SemApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation
	//{{AFX_MSG(CEduardo2SemApp)
	afx_msg void OnAppAbout();
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EDUARDO2SEM_H__B73E9385_6B5A_42C0_81CB_FE0AF756D5BE__INCLUDED_)
