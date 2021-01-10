// Eduardo2SemDoc.cpp : implementation of the CEduardo2SemDoc class
//

#include "stdafx.h"
#include "Eduardo2Sem.h"

#include "Eduardo2SemSet.h"
#include "Eduardo2SemDoc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CEduardo2SemDoc

IMPLEMENT_DYNCREATE(CEduardo2SemDoc, CDocument)

BEGIN_MESSAGE_MAP(CEduardo2SemDoc, CDocument)
	//{{AFX_MSG_MAP(CEduardo2SemDoc)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CEduardo2SemDoc construction/destruction

CEduardo2SemDoc::CEduardo2SemDoc()
{
	// TODO: add one-time construction code here

}

CEduardo2SemDoc::~CEduardo2SemDoc()
{
}

BOOL CEduardo2SemDoc::OnNewDocument()
{
	if (!CDocument::OnNewDocument())
		return FALSE;

	// TODO: add reinitialization code here
	// (SDI documents will reuse this document)

	return TRUE;
}



/////////////////////////////////////////////////////////////////////////////
// CEduardo2SemDoc serialization

void CEduardo2SemDoc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}

/////////////////////////////////////////////////////////////////////////////
// CEduardo2SemDoc diagnostics

#ifdef _DEBUG
void CEduardo2SemDoc::AssertValid() const
{
	CDocument::AssertValid();
}

void CEduardo2SemDoc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CEduardo2SemDoc commands
