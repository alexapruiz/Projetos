// GetCapaVincular.cpp : implementation file
//

#include "stdafx.h"
#include "GetCapaVincular.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGetCapaVincular

IMPLEMENT_DYNAMIC(CGetCapaVincular, CRecordset)

CGetCapaVincular::CGetCapaVincular(CDatabase* pdb)
	: CRecordset(pdb)
{
	//{{AFX_FIELD_INIT(CGetCapaVincular)
	m_IdCapa = 0;
	m_IdEnv_Mal = _T("");
	m_NumMalote = _T("");
	m_Status    = _T("");
	m_nFields = 4;
	//}}AFX_FIELD_INIT
	m_DataProc = 0;
	m_nParams = 1;
	m_nDefaultType = snapshot;
}


CString CGetCapaVincular::GetDefaultConnect()
{
	return _T("ODBC;DSN=MDI_Ubb");
}

CString CGetCapaVincular::GetDefaultSQL()
{
	return _T("{ Call VA_GetCapaPrioridadeVincular( ? ) }");
}

void CGetCapaVincular::DoFieldExchange(CFieldExchange* pFX)
{
	//{{AFX_FIELD_MAP(CGetCapaVincular)
	pFX->SetFieldType(CFieldExchange::outputColumn);
	RFX_Long(pFX, _T("IdCapa"), m_IdCapa);
	RFX_Text(pFX, _T("idEnv_Mal"), m_IdEnv_Mal);
	RFX_Text(pFX, _T("Num_Malote"), m_NumMalote);
	RFX_Text(pFX, _T("Status"), m_Status);
	//}}AFX_FIELD_MAP
    pFX->SetFieldType(CFieldExchange::param);
    RFX_Long(pFX, "DataProc", m_DataProc);
}

/////////////////////////////////////////////////////////////////////////////
// CGetCapaVincular diagnostics

#ifdef _DEBUG
void CGetCapaVincular::AssertValid() const
{
	CRecordset::AssertValid();
}

void CGetCapaVincular::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG
