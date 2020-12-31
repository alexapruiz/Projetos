// VincAutoDlg.h : header file
//

#if !defined(AFX_VINCAUTODLG_H__980227BF_4690_11D4_AF4D_000629E201DC__INCLUDED_)
#define AFX_VINCAUTODLG_H__980227BF_4690_11D4_AF4D_000629E201DC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Vinculador.h"

/////////////////////////////////////////////////////////////////////////////
// CVincAutoDlg dialog

class CVincAutoDlg : public CDialog
{
// Construction
public:
	CVincAutoDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CVincAutoDlg)
	enum { IDD = IDD_VINCAUTO_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CVincAutoDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;
	BOOL  m_Inicializado;
	CVinculador m_Vinculador;

	// Generated message map functions
	//{{AFX_MSG(CVincAutoDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnButtonInit();
	afx_msg void OnButtonDone();
	afx_msg void OnButtonExec();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_VINCAUTODLG_H__980227BF_4690_11D4_AF4D_000629E201DC__INCLUDED_)
