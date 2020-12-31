#if !defined(AFX_GETDOCUMENTOS_H__3EE628C1_46D0_11D4_AF4D_000629E201DC__INCLUDED_)
#define AFX_GETDOCUMENTOS_H__3EE628C1_46D0_11D4_AF4D_000629E201DC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GetDocumentos.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGetDocumentos recordset

class CGetDocumentos : public CRecordset
{
public:
	CGetDocumentos(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CGetDocumentos)

// Field/Param Data
	//{{AFX_FIELD(CGetDocumentos, CRecordset)
	long	m_IdDocto;
	int		m_TipoDocto;
	CString	m_Valor;
	CString m_Leitura;
	CString	m_Status;
	long    m_Vinculo;
	//}}AFX_FIELD
	long    m_DataProc;
	long    m_IdCapa;


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGetDocumentos)
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

#endif // !defined(AFX_GETDOCUMENTOS_H__3EE628C1_46D0_11D4_AF4D_000629E201DC__INCLUDED_)
