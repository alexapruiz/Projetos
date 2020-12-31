#if !defined(AFX_GETCAPAVINCULAR_H__5C544C43_460B_11D4_AF4D_000629E201DC__INCLUDED_)
#define AFX_GETCAPAVINCULAR_H__5C544C43_460B_11D4_AF4D_000629E201DC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GetCapaVincular.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGetCapaVincular recordset

class CGetCapaVincular : public CRecordset
{
public:
	CGetCapaVincular(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CGetCapaVincular)

// Field/Param Data
	//{{AFX_FIELD(CGetCapaVincular, CRecordset)
	long	m_IdCapa;
	CString	m_IdEnv_Mal;
	CString	m_NumMalote;
	CString m_Status;
	//}}AFX_FIELD
	LONG    m_DataProc;


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGetCapaVincular)
	public:
	virtual CString GetDefaultConnect();    // Default connection string
	virtual CString GetDefaultSQL();    // Default SQL for Recordset
	virtual void DoFieldExchange(CFieldExchange* pFX);  // RFX support
	//}}AFX_VIRTUAL

// Implementation
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GETCAPAVINCULAR_H__5C544C43_460B_11D4_AF4D_000629E201DC__INCLUDED_)
