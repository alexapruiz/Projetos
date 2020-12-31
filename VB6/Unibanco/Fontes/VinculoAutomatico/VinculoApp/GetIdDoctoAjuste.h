#if !defined(AFX_GETIDDOCTOAJUSTE_H__AC974C82_48EE_11D4_AF4D_000629E201DC__INCLUDED_)
#define AFX_GETIDDOCTOAJUSTE_H__AC974C82_48EE_11D4_AF4D_000629E201DC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GetIdDoctoAjuste.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGetIdDoctoAjuste recordset

class CGetIdDoctoAjuste : public CRecordset
{
public:
	CGetIdDoctoAjuste(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CGetIdDoctoAjuste)

// Field/Param Data
	//{{AFX_FIELD(CGetIdDoctoAjuste, CRecordset)
	long	m_IdDocto;
	//}}AFX_FIELD
	long    m_DataProc;
	long    m_IdCapa;
	int     m_TipoDocto;
	double  m_Valor;


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGetIdDoctoAjuste)
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

#endif // !defined(AFX_GETIDDOCTOAJUSTE_H__AC974C82_48EE_11D4_AF4D_000629E201DC__INCLUDED_)
