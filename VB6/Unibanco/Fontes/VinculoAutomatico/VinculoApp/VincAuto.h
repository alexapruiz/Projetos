// VincAuto.h : main header file for the VINCAUTO application
//

#if !defined(AFX_VINCAUTO_H__980227BD_4690_11D4_AF4D_000629E201DC__INCLUDED_)
#define AFX_VINCAUTO_H__980227BD_4690_11D4_AF4D_000629E201DC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CVincAutoApp:
// See VincAuto.cpp for the implementation of this class
//

class CVincAutoApp : public CWinApp
{
public:
	CVincAutoApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CVincAutoApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CVincAutoApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_VINCAUTO_H__980227BD_4690_11D4_AF4D_000629E201DC__INCLUDED_)
