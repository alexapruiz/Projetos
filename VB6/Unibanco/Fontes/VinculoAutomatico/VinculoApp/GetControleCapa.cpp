// GetControleCapa.cpp : implementation file
//

#include "stdafx.h"
#include "GetControleCapa.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGetControleCapa

IMPLEMENT_DYNAMIC(CGetControleCapa, CRecordset)

CGetControleCapa::CGetControleCapa(CDatabase* pdb)
	: CRecordset(pdb)
{
	//{{AFX_FIELD_INIT(CGetControleCapa)
	m_IdModulo   = 0;
	m_Comentario = "";
	m_nFields    = 2;
	//}}AFX_FIELD_INIT
	m_DataProc   = 0;
	m_IdCapa     = 0;
	m_nParams   = 2;
	m_nDefaultType = snapshot;

}


CString CGetControleCapa::GetDefaultConnect()
{
	return _T("ODBC;DSN=MDI_UBB");
}

CString CGetControleCapa::GetDefaultSQL()
{
	return _T("{ Call GetControleCapa( ? , ? ) }");
}

void CGetControleCapa::DoFieldExchange(CFieldExchange* pFX)
{
	//{{AFX_FIELD_MAP(CGetControleCapa)
	pFX->SetFieldType(CFieldExchange::outputColumn);
	RFX_Text(pFX, _T("Comentario"), m_Comentario);
	RFX_Long(pFX, _T("IdModulo"), m_IdModulo);
	//}}AFX_FIELD_MAP
	pFX->SetFieldType(CFieldExchange::param);
	RFX_Long(pFX, "DataProc", m_DataProc);
	RFX_Long(pFX, "IdCapa", m_IdCapa);

}

/////////////////////////////////////////////////////////////////////////////
// CGetControleCapa diagnostics

#ifdef _DEBUG
void CGetControleCapa::AssertValid() const
{
	CRecordset::AssertValid();
}

void CGetControleCapa::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


