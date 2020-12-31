// GetIdDoctoAjuste.cpp : implementation file
//

#include "stdafx.h"
#include "GetIdDoctoAjuste.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGetIdDoctoAjuste

IMPLEMENT_DYNAMIC(CGetIdDoctoAjuste, CRecordset)

CGetIdDoctoAjuste::CGetIdDoctoAjuste(CDatabase* pdb)
	: CRecordset(pdb)
{
	//{{AFX_FIELD_INIT(CGetIdDoctoAjuste)
	m_IdDocto = 0;
	m_nFields = 1;
	//}}AFX_FIELD_INIT
	m_DataProc  = 0;
	m_IdCapa    = 0;
	m_TipoDocto = 0;
	m_Valor     = 0.0;
	m_nParams   = 4;
	m_nDefaultType = snapshot;
}


CString CGetIdDoctoAjuste::GetDefaultConnect()
{
	return _T("ODBC;DSN=MDI_Ubb");
}

CString CGetIdDoctoAjuste::GetDefaultSQL()
{
	return _T("{ Call VA_GetIdDoctoAjuste( ?, ?, ?, ? ) }");
}

void CGetIdDoctoAjuste::DoFieldExchange(CFieldExchange* pFX)
{
	//{{AFX_FIELD_MAP(CGetIdDoctoAjuste)
	pFX->SetFieldType(CFieldExchange::outputColumn);
	RFX_Long(pFX, _T("IdDocto"), m_IdDocto);
	//}}AFX_FIELD_MAP
	pFX->SetFieldType(CFieldExchange::param);
	RFX_Long(pFX, _T("DataProc"), m_DataProc);
	RFX_Long(pFX, _T("IdCapa"), m_IdCapa);
	RFX_Int(pFX, _T("TipoDocto"), m_TipoDocto);
	RFX_Double(pFX, _T("Valor"), m_Valor);
}

/////////////////////////////////////////////////////////////////////////////
// CGetIdDoctoAjuste diagnostics

#ifdef _DEBUG
void CGetIdDoctoAjuste::AssertValid() const
{
	CRecordset::AssertValid();
}

void CGetIdDoctoAjuste::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG
