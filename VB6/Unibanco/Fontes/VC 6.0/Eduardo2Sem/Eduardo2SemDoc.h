// Eduardo2SemDoc.h : interface of the CEduardo2SemDoc class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_EDUARDO2SEMDOC_H__4C6B4336_31B4_4272_8430_EF9D72A4DE70__INCLUDED_)
#define AFX_EDUARDO2SEMDOC_H__4C6B4336_31B4_4272_8430_EF9D72A4DE70__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "Eduardo2SemSet.h"


class CEduardo2SemDoc : public CDocument
{
protected: // create from serialization only
	CEduardo2SemDoc();
	DECLARE_DYNCREATE(CEduardo2SemDoc)

// Attributes
public:
	CEduardo2SemSet m_eduardo2SemSet;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CEduardo2SemDoc)
	public:
	virtual BOOL OnNewDocument();
	virtual void Serialize(CArchive& ar);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CEduardo2SemDoc();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	//{{AFX_MSG(CEduardo2SemDoc)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EDUARDO2SEMDOC_H__4C6B4336_31B4_4272_8430_EF9D72A4DE70__INCLUDED_)
