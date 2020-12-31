#if !defined(AFX_GETCONTROLECAPA_H__D185DE3F_2337_4B70_A095_62861986595A__INCLUDED_)
#define AFX_GETCONTROLECAPA_H__D185DE3F_2337_4B70_A095_62861986595A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GetControleCapa.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGetControleCapa recordset

class CGetControleCapa : public CRecordset
{
public:
	CGetControleCapa(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CGetControleCapa)

// Field/Param Data
	//{{AFX_FIELD(CGetControleCapa, CRecordset)
	long m_IdModulo;
	CString m_Comentario;
	long i;
	//}}AFX_FIELD
	long m_DataProc;
	long m_IdCapa;


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGetControleCapa)
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
protected:
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GETCONTROLECAPA_H__D185DE3F_2337_4B70_A095_62861986595A__INCLUDED_)
