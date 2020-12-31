// GetDocumentos.cpp : implementation file
//

#include "stdafx.h"
#include "GetDocumentos.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGetDocumentos

IMPLEMENT_DYNAMIC(CGetDocumentos, CRecordset)

CGetDocumentos::CGetDocumentos(CDatabase* pdb)
	: CRecordset(pdb)
{
	//{{AFX_FIELD_INIT(CGetDocumentos)
	m_IdDocto = 0;
	m_TipoDocto = 0;
	m_Valor = _T("");
	m_Leitura = _T("");
	m_Status =  _T("");
	m_Vinculo = 0;
	m_nFields = 6;
	//}}AFX_FIELD_INIT
	m_DataProc = 0;
	m_IdCapa = 0;
	m_nParams = 2;
	m_nDefaultType = snapshot;
}


CString CGetDocumentos::GetDefaultConnect()
{
	return _T("ODBC;DSN=MDI_Ubb");
}

CString CGetDocumentos::GetDefaultSQL()
{
	return _T("{ Call VA_GetDocumentos( ?, ? ) }");
}

void CGetDocumentos::DoFieldExchange(CFieldExchange* pFX)
{
	//{{AFX_FIELD_MAP(CGetDocumentos)
	pFX->SetFieldType(CFieldExchange::outputColumn);
	RFX_Long(pFX, _T("IdDocto"), m_IdDocto);
	RFX_Int(pFX, _T("TipoDocto"), m_TipoDocto);
	RFX_Text(pFX, _T("Valor"), m_Valor);
	RFX_Text(pFX, _T("Leitura"), m_Leitura);
	RFX_Text(pFX, _T("Status"), m_Status);
	RFX_Long(pFX, _T("Vinculo"), m_Vinculo);
	//}}AFX_FIELD_MAP
    pFX->SetFieldType(CFieldExchange::param);
    RFX_Long(pFX, "DataProc", m_DataProc);
    RFX_Long(pFX, "IdCapa", m_IdCapa);
}

/////////////////////////////////////////////////////////////////////////////
// CGetDocumentos diagnostics

#ifdef _DEBUG
void CGetDocumentos::AssertValid() const
{
	CRecordset::AssertValid();
}

void CGetDocumentos::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG
