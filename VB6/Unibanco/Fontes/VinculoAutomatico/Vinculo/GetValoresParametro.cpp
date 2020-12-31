// GetValoresParametro.cpp : implementation file
//

#include "stdafx.h"
#include "GetValoresParametro.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGetValoresParametro

IMPLEMENT_DYNAMIC(CGetValoresParametro, CRecordset)

CGetValoresParametro::CGetValoresParametro(CDatabase* pdb)
	: CRecordset(pdb)
{
	//{{AFX_FIELD_INIT(CGetValoresParametro)
	m_ValorAjusteAuto_Env = _T("");
	m_ValorAjusteAuto_Mal = _T("");
	m_ValorAlcada_Env = _T("");
	m_ValorAlcada_Mal = _T("");
	m_ValorAlcadaDep_Env = _T("");
	m_ValorAlcadaDep_Mal = _T("");
	m_ValorAlcadaOutros_Env = _T("");
	m_ValorAlcadaOutros_Mal = _T("");
	m_ValorAjusteContabil = _T("");
	m_DataFinalRegraAntiga = 0;
	m_LimiteMaxDifLancto_Mal = _T("");
	m_nFields = 11;
	//}}AFX_FIELD_INIT
	m_DataProc = 0;
	m_nParams = 1;
	m_nDefaultType = snapshot;
}


CString CGetValoresParametro::GetDefaultConnect()
{
	return _T("ODBC;DSN=MDI_Ubb");
}

CString CGetValoresParametro::GetDefaultSQL()
{
	return _T("{ Call VA_GetValoresParametro( ? ) }");
}

void CGetValoresParametro::DoFieldExchange(CFieldExchange* pFX)
{
	//{{AFX_FIELD_MAP(CGetValoresParametro)
	pFX->SetFieldType(CFieldExchange::outputColumn);
	RFX_Text(pFX, _T("[ValorAlcada_Env]"), m_ValorAlcada_Env);
	RFX_Text(pFX, _T("[ValorAlcada_Mal]"), m_ValorAlcada_Mal);
	RFX_Text(pFX, _T("[ValorAlcadaDep_Env]"), m_ValorAlcadaDep_Env);
	RFX_Text(pFX, _T("[ValorAlcadaDep_Mal]"), m_ValorAlcadaDep_Mal);
	RFX_Text(pFX, _T("[ValorAlcadaOutros_Env]"), m_ValorAlcadaOutros_Env);
	RFX_Text(pFX, _T("[ValorAlcadaOutros_Mal]"), m_ValorAlcadaOutros_Mal);
	RFX_Text(pFX, _T("[ValorAjusteAuto_Env]"), m_ValorAjusteAuto_Env);
	RFX_Text(pFX, _T("[ValorAjusteAuto_Mal]"), m_ValorAjusteAuto_Mal);
	RFX_Text(pFX, _T("[ValorAjusteContabil]"), m_ValorAjusteContabil);
    RFX_Long(pFX, "DataFinalRegraAntiga_Mal", m_DataFinalRegraAntiga);
	RFX_Text(pFX, "LimiteMaxDifLancto_Mal", m_LimiteMaxDifLancto_Mal);
	//}}AFX_FIELD_MAP
    pFX->SetFieldType(CFieldExchange::param);
    RFX_Long(pFX, "DataProc", m_DataProc);

}


/////////////////////////////////////////////////////////////////////////////
// CGetValoresParametro diagnostics

#ifdef _DEBUG
void CGetValoresParametro::AssertValid() const
{
	CRecordset::AssertValid();
}

void CGetValoresParametro::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG
