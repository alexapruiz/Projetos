// GetAgContaDeposito.cpp : implementation file
//

#include "stdafx.h"
#include "GetAgContaDeposito.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGetAgContaDeposito

IMPLEMENT_DYNAMIC(CGetAgContaDeposito, CRecordset)

CGetAgContaDeposito::CGetAgContaDeposito(CDatabase* pdb)
	: CRecordset(pdb)
{
	//{{AFX_FIELD_INIT(CGetAgContaDeposito)
	m_Agencia = 0;
	m_Conta = _T("");
	m_nFields = 2;
	//}}AFX_FIELD_INIT
	m_DataProc  = 0;
	m_IdDocto   = 0;
	m_TipoDocto = 0;
	m_nParams   = 3;
	m_nDefaultType = snapshot;
}


CString CGetAgContaDeposito::GetDefaultConnect()
{
	return _T("ODBC;DSN=MDI_Ubb");
}

CString CGetAgContaDeposito::GetDefaultSQL()
{
	return _T("{ Call VA_GetAgContaDeposito( ?, ?, ? ) }");
}

void CGetAgContaDeposito::DoFieldExchange(CFieldExchange* pFX)
{
	//{{AFX_FIELD_MAP(CGetAgContaDeposito)
	pFX->SetFieldType(CFieldExchange::outputColumn);
	RFX_Int(pFX, _T("[Agencia]"), m_Agencia);
	RFX_Text(pFX, _T("[Conta]"), m_Conta);
	//}}AFX_FIELD_MAP
    pFX->SetFieldType(CFieldExchange::param);
    RFX_Long(pFX, "DataProc", m_DataProc);
    RFX_Long(pFX, "IdDocto", m_IdDocto);
    RFX_Long(pFX, "TipoDocto", m_TipoDocto);
}

/////////////////////////////////////////////////////////////////////////////
// CGetAgContaDeposito diagnostics

#ifdef _DEBUG
void CGetAgContaDeposito::AssertValid() const
{
	CRecordset::AssertValid();
}

void CGetAgContaDeposito::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG
