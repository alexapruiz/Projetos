// Eduardo2SemSet.h : interface of the CEduardo2SemSet class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_EDUARDO2SEMSET_H__139DCD44_6C14_4AA0_8F27_5B7898A1CFFC__INCLUDED_)
#define AFX_EDUARDO2SEMSET_H__139DCD44_6C14_4AA0_8F27_5B7898A1CFFC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CCLIENTES
{
public:
	CCLIENTES()
	{
		memset( (void*)this, 0, sizeof(*this) );
	};

	int m_COD_CLI;
	char m_NOM_CLI[31];
	char m_RG_CLI[101];
	char m_END_CLI[51];
	char m_FONE1_CLI[11];
	char m_FONE2_CLI[11];
	char m_CEL_CLI[11];


BEGIN_COLUMN_MAP(CCLIENTES)
		COLUMN_ENTRY_TYPE(1, DBTYPE_I4, m_COD_CLI)
		COLUMN_ENTRY_TYPE(2, DBTYPE_STR, m_NOM_CLI)
		COLUMN_ENTRY_TYPE(3, DBTYPE_STR, m_RG_CLI)
		COLUMN_ENTRY_TYPE(4, DBTYPE_STR, m_END_CLI)
		COLUMN_ENTRY_TYPE(5, DBTYPE_STR, m_FONE1_CLI)
		COLUMN_ENTRY_TYPE(6, DBTYPE_STR, m_FONE2_CLI)
		COLUMN_ENTRY_TYPE(7, DBTYPE_STR, m_CEL_CLI)
END_COLUMN_MAP()

};

class CEduardo2SemSet : public CCommand<CAccessor<CCLIENTES> >
{
public:

	HRESULT Open()
	{
		CDataSource db;
		CSession	session;
		HRESULT		hr;

		CDBPropSet	dbinit(DBPROPSET_DBINIT);
		dbinit.AddProperty(DBPROP_AUTH_INTEGRATED, "SSPI");
		dbinit.AddProperty(DBPROP_AUTH_PERSIST_SENSITIVE_AUTHINFO, false);
		dbinit.AddProperty(DBPROP_INIT_CATALOG,  "BUFFET");
		dbinit.AddProperty(DBPROP_INIT_DATASOURCE, "(local)");
		dbinit.AddProperty(DBPROP_INIT_LCID, (long)1046);
		dbinit.AddProperty(DBPROP_INIT_PROMPT, (short)4);

		hr = db.OpenWithServiceComponents("SQLOLEDB.1", &dbinit);
		if (FAILED(hr))
			return hr;

		hr = session.Open(db);
		if (FAILED(hr))
			return hr;

		CDBPropSet	propset(DBPROPSET_ROWSET);
		propset.AddProperty(DBPROP_CANFETCHBACKWARDS, true);
		propset.AddProperty(DBPROP_IRowsetScroll, true);
		propset.AddProperty(DBPROP_IRowsetChange, true);
		propset.AddProperty(DBPROP_UPDATABILITY, DBPROPVAL_UP_CHANGE | DBPROPVAL_UP_INSERT | DBPROPVAL_UP_DELETE );

		hr = CCommand<CAccessor<CCLIENTES> >::Open(session, "SELECT * FROM dbo.CLIENTES", &propset);
		if (FAILED(hr))
			return hr;

		return MoveNext();
	}

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EDUARDO2SEMSET_H__139DCD44_6C14_4AA0_8F27_5B7898A1CFFC__INCLUDED_)

