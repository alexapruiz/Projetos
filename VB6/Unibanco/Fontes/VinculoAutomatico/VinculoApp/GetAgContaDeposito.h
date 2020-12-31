#if !defined(AFX_GETAGCONTADEPOSITO_H__FB177120_482E_11D4_AF4D_000629E201DC__INCLUDED_)
#define AFX_GETAGCONTADEPOSITO_H__FB177120_482E_11D4_AF4D_000629E201DC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GetAgContaDeposito.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGetAgContaDeposito recordset

class CGetAgContaDeposito : public CRecordset
{
public:
	CGetAgContaDeposito(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CGetAgContaDeposito)

// Field/Param Data
	//{{AFX_FIELD(CGetAgContaDeposito, CRecordset)
	int		m_Agencia;
	CString	m_Conta;
	//}}AFX_FIELD
	long    m_DataProc;
	long	m_IdDocto;     
	long    m_TipoDocto;


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGetAgContaDeposito)
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

#endif // !defined(AFX_GETAGCONTADEPOSITO_H__FB177120_482E_11D4_AF4D_000629E201DC__INCLUDED_)
