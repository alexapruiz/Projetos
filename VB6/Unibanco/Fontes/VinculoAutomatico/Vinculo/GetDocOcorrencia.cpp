// GetDocOcorrencia.cpp : implementation file
//

#include "stdafx.h"
//#include "VincAuto.h"
#include "GetDocOcorrencia.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGetDocOcorrencia

IMPLEMENT_DYNAMIC(CGetDocOcorrencia, CRecordset)

CGetDocOcorrencia::CGetDocOcorrencia(CDatabase* pdb)
	: CRecordset(pdb)
{
	//{{AFX_FIELD_INIT(CGetDocOcorrencia)
	m_Qtde = 0;
	m_nFields = 1;
	//}}AFX_FIELD_INIT
	m_DataProc  = 0;
	m_IdCapa    = 0;
	m_nParams   = 2;
	m_nDefaultType = snapshot;
}


CString CGetDocOcorrencia::GetDefaultConnect()
{
	return _T("ODBC;DSN=MDI_Ubb");
}

CString CGetDocOcorrencia::GetDefaultSQL()
{
	return _T("{ Call VA_GetDocumentosOcorrencia( ?, ? ) }");
}

void CGetDocOcorrencia::DoFieldExchange(CFieldExchange* pFX)
{
	//{{AFX_FIELD_MAP(CGetDocOcorrencia)
	pFX->SetFieldType(CFieldExchange::outputColumn);
	RFX_Long(pFX, _T("Qtde"), m_Qtde);
	//}}AFX_FIELD_MAP
    pFX->SetFieldType(CFieldExchange::param);
    RFX_Long(pFX, "DataProc", m_DataProc);
    RFX_Long(pFX, "IdCapa", m_IdCapa);
}

/////////////////////////////////////////////////////////////////////////////
// CGetDocOcorrencia diagnostics

#ifdef _DEBUG
void CGetDocOcorrencia::AssertValid() const
{
	CRecordset::AssertValid();
}

void CGetDocOcorrencia::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG
