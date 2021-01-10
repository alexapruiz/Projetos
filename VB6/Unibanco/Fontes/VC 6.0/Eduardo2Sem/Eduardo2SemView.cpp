// Eduardo2SemView.cpp : implementation of the CEduardo2SemView class
//

#include "stdafx.h"
#include "Eduardo2Sem.h"

#include "Eduardo2SemSet.h"
#include "Eduardo2SemDoc.h"
#include "Eduardo2SemView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CEduardo2SemView

IMPLEMENT_DYNCREATE(CEduardo2SemView, COleDBRecordView)

BEGIN_MESSAGE_MAP(CEduardo2SemView, COleDBRecordView)
	//{{AFX_MSG_MAP(CEduardo2SemView)
	ON_BN_CLICKED(IDC_BUTTON1, OnButton1)
	//}}AFX_MSG_MAP
	// Standard printing commands
	ON_COMMAND(ID_FILE_PRINT, COleDBRecordView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, COleDBRecordView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, COleDBRecordView::OnFilePrintPreview)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CEduardo2SemView construction/destruction

CEduardo2SemView::CEduardo2SemView()
	: COleDBRecordView(CEduardo2SemView::IDD)
{
	//{{AFX_DATA_INIT(CEduardo2SemView)
		// NOTE: the ClassWizard will add member initialization here
	m_pSet = NULL;
	//}}AFX_DATA_INIT
	// TODO: add construction code here

}

CEduardo2SemView::~CEduardo2SemView()
{
}

void CEduardo2SemView::DoDataExchange(CDataExchange* pDX)
{
	COleDBRecordView::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CEduardo2SemView)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}

BOOL CEduardo2SemView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return COleDBRecordView::PreCreateWindow(cs);
}

void CEduardo2SemView::OnInitialUpdate()
{
	m_pSet = &GetDocument()->m_eduardo2SemSet;
	{
		CWaitCursor wait;
		HRESULT hr = m_pSet->Open();
		if (hr != S_OK)
		{
			AfxMessageBox(_T("Record set failed to open."), MB_OK);
			// Disable the Next and Previous record commands,
			// since attempting to change the current record without an
			// open RecordSet will cause a crash.
			m_bOnFirstRecord = TRUE;
			m_bOnLastRecord = TRUE;
		}				
	}
	COleDBRecordView::OnInitialUpdate();

}

/////////////////////////////////////////////////////////////////////////////
// CEduardo2SemView printing

BOOL CEduardo2SemView::OnPreparePrinting(CPrintInfo* pInfo)
{
	// default preparation
	return DoPreparePrinting(pInfo);
}

void CEduardo2SemView::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add extra initialization before printing
}

void CEduardo2SemView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add cleanup after printing
}

/////////////////////////////////////////////////////////////////////////////
// CEduardo2SemView diagnostics

#ifdef _DEBUG
void CEduardo2SemView::AssertValid() const
{
	COleDBRecordView::AssertValid();
}

void CEduardo2SemView::Dump(CDumpContext& dc) const
{
	COleDBRecordView::Dump(dc);
}

CEduardo2SemDoc* CEduardo2SemView::GetDocument() // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CEduardo2SemDoc)));
	return (CEduardo2SemDoc*)m_pDocument;
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CEduardo2SemView database support
CRowset* CEduardo2SemView::OnGetRowset()
{
	return m_pSet;
}


/////////////////////////////////////////////////////////////////////////////
// CEduardo2SemView message handlers

void CEduardo2SemView::OnButton1() 
{
	// TODO: Add your control notification handler code here
	AfxMessageBox(IDC_EDIT1 ,MB_OK);
}
