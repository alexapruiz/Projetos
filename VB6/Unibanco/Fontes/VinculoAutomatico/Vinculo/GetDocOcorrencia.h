#if !defined(AFX_GETDOCOCORRENCIA_H__A5815D81_6D35_11D4_A770_0080C8E52FD2__INCLUDED_)
#define AFX_GETDOCOCORRENCIA_H__A5815D81_6D35_11D4_A770_0080C8E52FD2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GetDocOcorrencia.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGetDocOcorrencia recordset

class CGetDocOcorrencia : public CRecordset
{
public:
	CGetDocOcorrencia(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CGetDocOcorrencia)

// Field/Param Data
	//{{AFX_FIELD(CGetDocOcorrencia, CRecordset)
	long	m_Qtde;
	//}}AFX_FIELD
	long    m_DataProc;
	long    m_IdCapa;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGetDocOcorrencia)
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

#endif // !defined(AFX_GETDOCOCORRENCIA_H__A5815D81_6D35_11D4_A770_0080C8E52FD2__INCLUDED_)
