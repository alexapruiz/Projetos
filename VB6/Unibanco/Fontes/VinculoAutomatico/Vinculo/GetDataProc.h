#if !defined(AFX_GETDATAPROC_H__D5F5FCEF_45CC_11D4_AF4D_000629E201DC__INCLUDED_)
#define AFX_GETDATAPROC_H__D5F5FCEF_45CC_11D4_AF4D_000629E201DC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GetDataProc.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGetDataProc recordset

class CGetDataProc : public CRecordset
{
public:
	CGetDataProc(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CGetDataProc)

// Field/Param Data
	//{{AFX_FIELD(CGetDataProc, CRecordset)
	long	m_DataProc;
	int     m_Sleep;
	//}}AFX_FIELD


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGetDataProc)
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

#endif // !defined(AFX_GETDATAPROC_H__D5F5FCEF_45CC_11D4_AF4D_000629E201DC__INCLUDED_)
