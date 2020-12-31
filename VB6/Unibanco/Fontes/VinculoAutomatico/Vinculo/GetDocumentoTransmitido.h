#if !defined(AFX_GETDOCUMENTOTRANSMITIDO_H__E94B32C0_83F6_11D4_A1FA_0080C8E45072__INCLUDED_)
#define AFX_GETDOCUMENTOTRANSMITIDO_H__E94B32C0_83F6_11D4_A1FA_0080C8E45072__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GetDocumentoTransmitido.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGetDocumentoTransmitido recordset

class CGetDocumentoTransmitido : public CRecordset
{
public:
	CGetDocumentoTransmitido(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CGetDocumentoTransmitido)

// Field/Param Data
	//{{AFX_FIELD(CGetDocumentoTransmitido, CRecordset)
	long	m_Qtde;
	long	m_QtdeDoctosDaCapa;
	//}}AFX_FIELD
	long    m_DataProc;
	long    m_IdCapa;


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGetDocumentoTransmitido)
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

#endif // !defined(AFX_GETDOCUMENTOTRANSMITIDO_H__E94B32C0_83F6_11D4_A1FA_0080C8E45072__INCLUDED_)
