// VincAutoDlg.cpp : implementation file
//

#include "stdafx.h"
#include "VincAuto.h"
#include "VincAutoDlg.h"
#include "Vinculador.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CVincAutoDlg dialog

CVincAutoDlg::CVincAutoDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CVincAutoDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CVincAutoDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CVincAutoDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CVincAutoDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CVincAutoDlg, CDialog)
	//{{AFX_MSG_MAP(CVincAutoDlg)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_INIT, OnButtonInit)
	ON_BN_CLICKED(IDC_BUTTON_DONE, OnButtonDone)
	ON_BN_CLICKED(IDC_BUTTON_EXEC, OnButtonExec)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CVincAutoDlg message handlers

BOOL CVincAutoDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	m_Inicializado = FALSE;

	SetWindowText("Vínculo Automático 2.16");
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CVincAutoDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CVincAutoDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CVincAutoDlg::OnButtonInit() 
{
	char Msg[256];
	int iCodErr;

	if( !m_Vinculador.Init() )
	{
		m_Vinculador.GetLastErrorInfo( iCodErr, Msg );
		MessageBox( Msg, "Vinculo Automatico", MB_OK | MB_ICONSTOP );
	}
	else
	{
		m_Inicializado = TRUE;
	}
}

void CVincAutoDlg::OnButtonDone() 
{
	if( m_Inicializado )
	{
		m_Vinculador.Done();
	}
}

void CVincAutoDlg::OnButtonExec() 
{
	char Msg[256];
	int iCodErr;
	int iCodRet;

	iCodRet = m_Vinculador.ProcessaVinculo();
	if( iCodRet < 0 )
	{
		m_Vinculador.GetLastErrorInfo( iCodErr, Msg );
		MessageBox( Msg, "Vinculo Automatico", MB_OK | MB_ICONSTOP );
	}
	else if( iCodRet == 0 )
	{
		MessageBox( "Não há capas para vincular", "Vinculo Automatico", MB_OK | MB_ICONEXCLAMATION );
	}
	else
	{
		MessageBox( "Vinculação concluída com sucesso!", "Vinculo Automatico", MB_OK | MB_ICONINFORMATION );
	}
}
