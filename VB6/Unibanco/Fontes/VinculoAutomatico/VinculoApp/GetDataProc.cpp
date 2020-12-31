// GetDataProc.cpp : implementation file
//

#include "stdafx.h"
#include "GetDataProc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGetDataProc

IMPLEMENT_DYNAMIC(CGetDataProc, CRecordset)

CGetDataProc::CGetDataProc(CDatabase* pdb)
	: CRecordset(pdb)
{
	//{{AFX_FIELD_INIT(CGetDataProc)
	m_DataProc = 0;
	m_Sleep    = 0;
	m_nFields = 2;
	//}}AFX_FIELD_INIT
	m_nDefaultType = snapshot;
}


CString CGetDataProc::GetDefaultConnect()
{
	return _T("ODBC;DSN=MDI_Ubb");
}

CString CGetDataProc::GetDefaultSQL()
{
	return _T("{ Call VA_GetDataProc }");
}

void CGetDataProc::DoFieldExchange(CFieldExchange* pFX)
{
	//{{AFX_FIELD_MAP(CGetDataProc)
	pFX->SetFieldType(CFieldExchange::outputColumn);
	RFX_Long(pFX, _T("DataProcessamento"), m_DataProc);
	RFX_Int(pFX, _T("Sleep"), m_Sleep);
	//}}AFX_FIELD_MAP
}

/////////////////////////////////////////////////////////////////////////////
// CGetDataProc diagnostics

#ifdef _DEBUG
void CGetDataProc::AssertValid() const
{
	CRecordset::AssertValid();
}

void CGetDataProc::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG
