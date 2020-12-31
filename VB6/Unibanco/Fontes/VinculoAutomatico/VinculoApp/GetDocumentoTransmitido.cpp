// GetDocumentoTransmitido.cpp : implementation file
//

#include "stdafx.h"
#include "GetDocumentoTransmitido.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGetDocumentoTransmitido

IMPLEMENT_DYNAMIC(CGetDocumentoTransmitido, CRecordset)

CGetDocumentoTransmitido::CGetDocumentoTransmitido(CDatabase* pdb)
	: CRecordset(pdb)
{
	//{{AFX_FIELD_INIT(CGetDocumentoTransmitido)
	m_Qtde = 0;
	m_QtdeDoctosDaCapa = 0;
	m_nFields = 2;
	//}}AFX_FIELD_INIT
	m_DataProc  = 0;
	m_IdCapa    = 0;
	m_nParams   = 2;
	m_nDefaultType = snapshot;
}


CString CGetDocumentoTransmitido::GetDefaultConnect()
{
	return _T("ODBC;DSN=MDI_Ubb");
}

CString CGetDocumentoTransmitido::GetDefaultSQL()
{
	return _T("{ Call VA_GetDocumentosTransmitidos( ?, ? ) }");
}

void CGetDocumentoTransmitido::DoFieldExchange(CFieldExchange* pFX)
{
	//{{AFX_FIELD_MAP(CGetDocumentoTransmitido)
	pFX->SetFieldType(CFieldExchange::outputColumn);
	RFX_Long(pFX, _T("Qtde"), m_Qtde);
	RFX_Long(pFX, _T("TotalDoctosDaCapa"), m_QtdeDoctosDaCapa);
	//}}AFX_FIELD_MAP
    pFX->SetFieldType(CFieldExchange::param);
    RFX_Long(pFX, "DataProc", m_DataProc);
    RFX_Long(pFX, "IdCapa", m_IdCapa);
}

/////////////////////////////////////////////////////////////////////////////
// CGetDocumentoTransmitido diagnostics

#ifdef _DEBUG
void CGetDocumentoTransmitido::AssertValid() const
{
	CRecordset::AssertValid();
}

void CGetDocumentoTransmitido::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG
