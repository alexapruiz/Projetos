#if !defined(AFX_GETVALORESPARAMETRO_H__5C544C40_460B_11D4_AF4D_000629E201DC__INCLUDED_)
#define AFX_GETVALORESPARAMETRO_H__5C544C40_460B_11D4_AF4D_000629E201DC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GetValoresParametro.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGetValoresParametro recordset

class CGetValoresParametro : public CRecordset
{
public:
	CGetValoresParametro(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CGetValoresParametro)

// Field/Param Data
	//{{AFX_FIELD(CGetValoresParametro, CRecordset)
	CString	m_ValorAjusteAuto_Env;
	CString	m_ValorAjusteAuto_Mal;
	CString	m_ValorAlcada_Env;
	CString	m_ValorAlcada_Mal;
	CString	m_ValorAlcadaDep_Env;
	CString	m_ValorAlcadaDep_Mal;
	CString	m_ValorAlcadaOutros_Env;
	CString	m_ValorAlcadaOutros_Mal;
	CString m_ValorAjusteContabil;
	long    m_DataFinalRegraAntiga;
	CString m_LimiteMaxDifLancto_Mal;
	//}}AFX_FIELD
	LONG    m_DataProc;



// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGetValoresParametro)
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

#endif // !defined(AFX_GETVALORESPARAMETRO_H__5C544C40_460B_11D4_AF4D_000629E201DC__INCLUDED_)
