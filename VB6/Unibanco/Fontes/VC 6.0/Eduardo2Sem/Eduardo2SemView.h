// Eduardo2SemView.h : interface of the CEduardo2SemView class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_EDUARDO2SEMVIEW_H__63F81989_C81C_4E41_BA2F_3F12FC0EC2F5__INCLUDED_)
#define AFX_EDUARDO2SEMVIEW_H__63F81989_C81C_4E41_BA2F_3F12FC0EC2F5__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CEduardo2SemSet;

class CEduardo2SemView : public COleDBRecordView
{
protected: // create from serialization only
	CEduardo2SemView();
	DECLARE_DYNCREATE(CEduardo2SemView)

public:
	//{{AFX_DATA(CEduardo2SemView)
	enum{ IDD = IDD_EDUARDO2SEM_FORM };
	CEduardo2SemSet* m_pSet;
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

// Attributes
public:
	CEduardo2SemDoc* GetDocument();

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CEduardo2SemView)
	public:
	virtual CRowset* OnGetRowset();
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual void OnInitialUpdate(); // called first time after construct
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CEduardo2SemView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	//{{AFX_MSG(CEduardo2SemView)
	afx_msg void OnButton1();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

#ifndef _DEBUG  // debug version in Eduardo2SemView.cpp
inline CEduardo2SemDoc* CEduardo2SemView::GetDocument()
   { return (CEduardo2SemDoc*)m_pDocument; }
#endif

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EDUARDO2SEMVIEW_H__63F81989_C81C_4E41_BA2F_3F12FC0EC2F5__INCLUDED_)
