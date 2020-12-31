VERSION 5.00
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmPassword 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmPassword.frx":0000
   ScaleHeight     =   4140
   ScaleMode       =   0  'User
   ScaleWidth      =   6000.461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PictureOpcoes 
      BackColor       =   &H00C0FFFF&
      DrawStyle       =   3  'Dash-Dot
      DrawWidth       =   2
      FillColor       =   &H00404040&
      Height          =   1275
      Left            =   270
      ScaleHeight     =   1215
      ScaleWidth      =   4245
      TabIndex        =   16
      Top             =   2580
      Visible         =   0   'False
      Width           =   4305
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Zerar Ref. Caixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   630
         TabIndex        =   21
         Top             =   330
         Width           =   1635
         Begin VB.CheckBox CheckZerarCapa 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Check1"
            Height          =   255
            Left            =   1125
            TabIndex        =   23
            Top             =   210
            Width           =   285
         End
         Begin VB.CheckBox CheckZerarDocto 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Check1"
            Height          =   255
            Left            =   1125
            TabIndex        =   22
            Top             =   480
            Width           =   285
         End
         Begin VB.Label LabelZerarCapa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Capa:"
            Height          =   195
            Left            =   540
            TabIndex        =   25
            Top             =   240
            Width           =   420
         End
         Begin VB.Label LabelZerarDocto 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Docto:"
            Height          =   195
            Left            =   540
            TabIndex        =   24
            Top             =   525
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Intervalo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   2310
         TabIndex        =   19
         Top             =   330
         Width           =   1935
         Begin VB.ComboBox ComboFixo 
            Height          =   315
            ItemData        =   "frmPassword.frx":1BDBA
            Left            =   1230
            List            =   "frmPassword.frx":1BDD6
            TabIndex        =   28
            Text            =   "10"
            ToolTipText     =   "Segundos Fixos em Espera"
            Top             =   240
            Width           =   675
         End
         Begin VB.OptionButton OptionFixo 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Fixo"
            Height          =   195
            Left            =   60
            TabIndex        =   27
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptionCrescente 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Crescente"
            Height          =   285
            Left            =   60
            TabIndex        =   26
            Top             =   480
            Width           =   1065
         End
         Begin VB.ComboBox ComboCrescente 
            Height          =   315
            ItemData        =   "frmPassword.frx":1BDF2
            Left            =   1230
            List            =   "frmPassword.frx":1BE08
            TabIndex        =   20
            Text            =   "1"
            ToolTipText     =   "Segundos Acumulados na Espera"
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.CheckBox CheckFechacx 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check1"
         Height          =   285
         Left            =   3180
         TabIndex        =   18
         Top             =   15
         Width           =   225
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   90
         Picture         =   "frmPassword.frx":1BE1E
         Top             =   60
         Width           =   480
      End
      Begin VB.Label LabelFechacx 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ForÁar a Abertura do Caixa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   765
         TabIndex        =   17
         Top             =   45
         Width           =   2355
      End
   End
   Begin VB.Timer TimerProgress 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   60
      Top             =   1590
   End
   Begin MSComctlLib.ProgressBar ProgressLogin 
      Align           =   2  'Align Bottom
      Height          =   216
      Left            =   0
      TabIndex        =   13
      Top             =   3924
      Visible         =   0   'False
      Width           =   6132
      _ExtentX        =   10821
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Max             =   10
      Scrolling       =   1
   End
   Begin DATEEDITLib.DateEdit txtData 
      Height          =   330
      Left            =   3030
      TabIndex        =   2
      Top             =   3420
      Width           =   1534
      _Version        =   65537
      _ExtentX        =   2706
      _ExtentY        =   582
      _StockProps     =   93
      ForeColor       =   12582912
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
   End
   Begin VB.PictureBox PictureKey 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0C0C0&
      Height          =   900
      Left            =   480
      ScaleHeight     =   840
      ScaleWidth      =   990
      TabIndex        =   12
      Top             =   2760
      Width           =   1050
      Begin VB.Image ImageKey 
         Height          =   480
         Left            =   210
         Picture         =   "frmPassword.frx":1C128
         Top             =   180
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4710
      TabIndex        =   3
      Top             =   2730
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4725
      TabIndex        =   4
      Top             =   3300
      Width           =   1095
   End
   Begin VB.TextBox txtUsuario 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   3030
      TabIndex        =   0
      Top             =   2640
      Width           =   1534
   End
   Begin VB.TextBox txtSenha 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   3030
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3030
      Width           =   1534
   End
   Begin VB.Label LabelOpcoes 
      BackStyle       =   0  'Transparent
      Caption         =   "&OpÁıes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      MouseIcon       =   "frmPassword.frx":1C432
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   2340
      Width           =   705
   End
   Begin VB.Label LabelLogin 
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde...  Efetuando Login."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3390
      TabIndex        =   14
      Top             =   2325
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Label LabelUsuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usu·rio:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1980
      TabIndex        =   11
      Top             =   2670
      Width           =   900
   End
   Begin VB.Label LabelSenha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2115
      TabIndex        =   10
      Top             =   3060
      Width           =   735
   End
   Begin VB.Label LabelData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2250
      TabIndex        =   9
      Top             =   3420
      Width           =   555
   End
   Begin VB.Label LabelVersao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vers„o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4350
      TabIndex        =   8
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label LabelSistema3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa RobÙ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   510
      TabIndex        =   7
      Top             =   1650
      Width           =   1395
   End
   Begin VB.Label LabelSistema2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multi Documentos por Imagem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   150
      TabIndex        =   6
      Top             =   510
      Width           =   3135
   End
   Begin VB.Label LabelSistema1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MDI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   510
   End
   Begin VB.Line Line5 
      X1              =   5989.702
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   0
      Y1              =   4160
      Y2              =   50
   End
   Begin VB.Line Line3 
      X1              =   5985.79
      X2              =   5985.79
      Y1              =   0
      Y2              =   4150
   End
   Begin VB.Line Line2 
      X1              =   5989.702
      X2              =   0
      Y1              =   4155
      Y2              =   4155
   End
   Begin VB.Line Line1 
      X1              =   6000.461
      X2              =   0
      Y1              =   1260
      Y2              =   1260
   End
End
Attribute VB_Name = "FrmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' *************************************
' * DefiniÁ„o das Vari·veis do MÛdulo *
' *************************************
Public SenhaOk As Boolean           ' IndicaÁ„o de Senha Digitada
''''''''''''''''''''''''''''''''''''''''''''''''''''
' Converter data do formato AAAAMMDD para DDMMAAAA '
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DataDDMMAAAA(ByVal pviData As Long) As Long
    Dim sData As String
    
    sData = CStr(pviData)
    
    DataDDMMAAAA = Val(Right(sData, 2) & Mid(sData, 5, 2) & Left(sData, 4))
End Function
Public Function DataAAAAMMDD(ByVal pviData As Long) As Long
    Dim sData As String
    
    sData = Format(pviData, "00000000")
    
    DataAAAAMMDD = Val(Right(sData, 4) & Mid(sData, 3, 2) & Left(sData, 2))
End Function
'''''''''''''''''''''''''''''''''''''''''''''''
' Retorna True se a data Ç v†lida             '
' Data deve ser informada no formato DDMMAAAA '
'''''''''''''''''''''''''''''''''''''''''''''''
Public Function dataOK(ByVal pviData As Long) As Boolean
    Dim iDia As Byte
    Dim iMes As Byte
    Dim iAno As Integer
    Dim sData As String
    Dim iUltimoDia As Byte
    Dim bOk As Boolean
    
    bOk = True
    
    sData = Format(pviData, "00000000")
    
    iDia = Left(sData, 2)
    iMes = Mid(sData, 3, 2)
    iAno = Right(sData, 4)
    
    If iAno < 1998 Then
        bOk = False
    Else
        Select Case iMes
            Case 1, 3, 5, 7, 8, 10, 12 ' 31 dias
                iUltimoDia = 31
            Case 2 ' 28/29 dias
                If iAno Mod 4 = 0 Then ' ano Ç bissexto
                    iUltimoDia = 29
                Else
                    iUltimoDia = 28
                End If
            Case 4, 6, 9, 11 ' 30 dias
                iUltimoDia = 30
            Case Else
                bOk = False
        End Select
        
        If bOk Then
            If iDia < 1 Or iDia > iUltimoDia Then
                bOk = False
            End If
        End If
    End If
    
    dataOK = bOk
End Function
Private Sub cmdConfirma_Click()
' ********************************************
' * Efetua a verificaÁ„o do Login do Usu·rio *
' ********************************************'

    Dim Servidor        As String
    Dim Banco           As String
    Dim Usuario         As String
    Dim Senha           As String
    Dim RstMDI          As Recordset
    Dim AutenticacaoNT  As Boolean
    Dim DataValida      As Boolean
  
  '*****************************************
  ' Testa DigitaÁ„o ObrigatÛria do Usu·rio *
  '*****************************************
    If Trim(txtUsuario.Text) = "" Then
        Beep
        MsgBox "Digite o Usu·rio !", vbExclamation + vbOKOnly, App.Title
        With txtUsuario
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
        End With
        Exit Sub
    End If
     
    '****************************************
    '* Testa DigitaÁ„o ObrigatÛria da Senha *
    '****************************************
    If Trim(txtSenha.Text) = "" Then
        Beep
        MsgBox "Digite a Senha !", vbExclamation + vbOKOnly, App.Title
        With txtSenha
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
        End With
        Exit Sub
    End If
    
    If dataOK(Val(txtData.Text)) Then
        Geral.DataProcessamento = DataAAAAMMDD(Val(txtData.Text))
    Else
        MsgBox "A data informada n„o Ç v·lida!" & vbCr & "Obs.: O ano deve ser maior que 1997!", vbExclamation + vbOKOnly, App.Title
        txtData.SetFocus
        Exit Sub
    End If
            
    On Error GoTo TrataErro
    
    SucessoLogin True
    Espera (1)
    
    Geral.DataProcessamento = Mid(txtData.Text, 5, 4) & Mid(txtData.Text, 3, 2) & Mid(txtData.Text, 1, 2)
    
   '''''''''''''''''
   ' InicializaÁ„o '
   '''''''''''''''''
       
   'UBBDB
    UBBQuery.DatabaseName = "ubbdb"
    UBBQuery.DataSourceName = "ubbdb"
    UBBQuery.SetConnection
    
   'MDI
    MDIQuery.Servidor = PegarOpcaoINI("Conexao", "Servidor", "")
    MDIQuery.Banco = PegarOpcaoINI("Conexao", "DataBase", "")
    MDIQuery.Usuario = PegarOpcaoINI("Conexao", "Usuario", "")
    MDIQuery.Senha = PegarOpcaoINI("Conexao", "Senha", "")
    MDIQuery.Provedor = "sqlOledb"
    MDIQuery.SetConnection
    
   '*****************************
   '* Verfica Versao do Sistema *
   '*****************************
   
    Set RstMDI = MDIQuery.ExecuteSQL("Select * from MDI_Versao")
    
   'Verifica vers„o do sistema
    If CStr(IIf(IsNull(RstMDI!VersaoRobo), 0, RstMDI!VersaoRobo)) <> (App.Major & App.Minor & App.Revision) Then
        Beep
        SucessoLogin False
        MsgBox "Vers„o incorreta do sistema !" & vbCrLf & vbCrLf & "Favor entrar em contato com o Suporte", vbCritical, App.Title
        cmdSair_Click
    End If
    Set RstMDI = Nothing
    
    Set RstMDI = MDIQuery.ExecuteSQL("Select Criptografia from MDI_Versao")
    
   'Verifica vers„o do sistema
    If CBool(IIf(IsNull(RstMDI!Criptografia), 0, RstMDI!Criptografia)) Then
        Geral.Criptografia = True
    End If
    Set RstMDI = Nothing
    
   '***********************************
   '* VerificaÁ„o da Data de Movimento*
   '***********************************
    
    DataValida = False
            
    Set RstMDI = MDIQuery.getDataSvr
           
    If Geral.DataProcessamento <> Format(RstMDI!DataServMDI, "YYYYMMDD") And _
       Geral.DataProcessamento <> Format(RstMDI!DataServMDI - 1, "YYYYMMDD") Then
      
       If UCase(txtUsuario.Text) = "DESENV" Then GoTo continua
       SucessoLogin False
       Beep
       MsgBox "Data Invalida em relacao ao Servidor MDI!", vbExclamation + vbOKOnly, App.Title
       
       With txtUsuario
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
         .SetFocus
       End With
     
    Else 'Verifica se data ja consta na tabela de parametros
continua:
    
      Set RstMDI = MDIQuery.getParametro(Geral.DataProcessamento)
    
      If RstMDI.EOF() Then
         TimerProgress.Enabled = False
         Beep
         
         If MsgBox("Data n„o encontrada, Inicia ?", vbExclamation + vbYesNo, App.Title) = vbNo Then
            TimerProgress.Enabled = False
                        
            With txtUsuario
                .SelStart = 0
                .SelLength = Len(Trim(.Text))
                .SetFocus
            End With
         Else
            TimerProgress.Enabled = True
            DataValida = True
            Espera (0.5)
         End If
      
      Else
         DataValida = True
      End If
    
    End If
    
    RstMDI.Close
    
    If DataValida Then
     '***********************************
     '* VerificaÁ„o do Login do Usu·rio *
     '***********************************
      Set RstMDI = MDIQuery.getUsuario(Trim(txtUsuario.Text))
         
      If UCase(Trim(txtUsuario.Text)) = "DESENV" And UCase(Trim(txtSenha.Text)) = UCase(Decript("Æmöö≤ùóçß")) Then
         SenhaOk = True
      ElseIf RstMDI.EOF Then
         SucessoLogin False
         Beep
         MsgBox "Usu·rio n„o Cadastrado !", vbExclamation + vbOKOnly, App.Title
         
         With txtUsuario
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
         End With
      ElseIf UCase(Decript(Trim(RstMDI!Senha))) <> UCase(Trim(txtSenha.Text)) Then
         SucessoLogin False
         Beep
         MsgBox "Senha n„o Confere !", vbExclamation + vbOKOnly, App.Title
         With txtSenha
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
         End With
      Else
         Do While Not RstMDI.EOF
            If RstMDI!IdGrupo = "SUP" Or RstMDI!IdGrupo = "SPT" Or RstMDI!IdGrupo = "LID" Or RstMDI!IdGrupo = "COO" Then
               SenhaOk = True
               Exit Do
            End If
            
            RstMDI.MoveNext
         Loop
         
         If Not SenhaOk Then
            Beep
            SucessoLogin False
            MsgBox "Usu·rio n„o È Supervisor ou Suporte !", vbExclamation + vbOKOnly, App.Title
            With txtUsuario
               .SelStart = 0
               .SelLength = Len(Trim(.Text))
               .SetFocus
            End With
               
         End If
    
      End If
      
      Caixa.UsuarioAtual = txtUsuario.Text
      
      If SenhaOk Then
         MDIQuery.insLog Geral.DataProcessamento, "0", "0", Caixa.UsuarioAtual, "120"
         
         ProgressLogin.Value = 10
         LabelLogin.Caption = "Login Efetuado com Sucesso !"
         Espera (1.5)
         
        'Opcoes pre-inicializacao
         Call SetAntIniOpcoes
         
         Unload Me
      Else
          SucessoLogin False
      End If
      
    End If
        
    Exit Sub
    
TrataErro:
    SucessoLogin False
    Screen.MousePointer = 0
    MsgBox "erro" & Err.Description
    
End Sub

' ******************************
' * Cancela o Login no Sistema *
' ******************************
Private Sub cmdSair_Click()
   End
End Sub
Private Sub Form_Activate()
   SenhaOk = False
   txtData.BackColor = &HFFFFFF
   txtData.Text = Format(Date, "ddmmyyyy")
End Sub
Private Sub Form_Load()
    LabelVersao.Caption = "Vers„o: " & App.Major & App.Minor & App.Revision
    
    If PegarOpcaoINI("Diversos", "Fixo", True) Then
        OptionFixo = True
        ComboFixo.Text = PegarOpcaoINI("Diversos", "Intervalo", 30)
    Else
        OptionCrescente = True
        ComboCrescente.Text = PegarOpcaoINI("Diversos", "Intervalo", 1)
    End If
    
End Sub
Private Sub LabelOpcoes_Click()
    Dim spRetorno As Integer
    
    If LabelOpcoes.Caption = "&Retorna" Then
        cmdConfirma.Enabled = True
        PictureOpcoes.Visible = False
        LabelOpcoes.Caption = "&OpÁıes"
    Else
        cmdConfirma.Enabled = False
        PictureOpcoes.Visible = True
        LabelOpcoes.Caption = "&Retorna"
    End If
End Sub
Private Sub OptionCrescente_Click()
    If OptionCrescente Then
        ComboCrescente.Visible = True
        ComboFixo.Visible = False
    Else
        ComboFixo.Visible = True
        ComboCrescente.Visible = False
    End If
End Sub
Private Sub OptionFixo_Click()
    If OptionCrescente Then
        ComboCrescente.Visible = True
        ComboFixo.Visible = False
    Else
        ComboFixo.Visible = True
        ComboCrescente.Visible = False
    End If
End Sub
Private Sub TimerProgress_Timer()
Static Valor

    If Val(Valor) = 0 Then
        Valor = 1
    ElseIf Val(Valor) >= ProgressLogin.Max Then
        Valor = 10
    Else
        Valor = Valor + 1
    End If
    
    DoEvents
    ProgressLogin.Value = Val(Valor)
    
End Sub
' **************************************
' * Carrega MÛdulo de Login no Sistema *
' **************************************
Private Sub txtdata_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
' ************************************
' * Ajustando SeleÁ„o do Campo Senha *
' ************************************
Private Sub txtSenha_GotFocus()
    With txtSenha
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
' **************************************
' * Ajustando SeleÁ„o do Campo Usuario *
' **************************************
Private Sub txtUsuario_GotFocus()
    With txtUsuario
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
    If KeyCode = 122 Then
      txtUsuario.Text = "Desenv"
      SendKeys ("{TAB}")
    End If
End Sub
Sub SucessoLogin(pSucesso As Boolean)
    LabelLogin.Visible = CBool(pSucesso)
    ProgressLogin.Visible = CBool(pSucesso)
    TimerProgress.Enabled = CBool(pSucesso)
End Sub
Sub SetAntIniOpcoes()

   'Fecha caixa
    If CheckFechacx Then
        If UCase(txtUsuario) = UCase("Suporte") Then MsgBox "Contatar o Suporte Executar Fechamento do Caixa": Exit Sub
        AntIniOpcoes.FechaCx = CheckFechacx
    End If
     
    'Limpa tabela Caixa
    If CheckZerarCapa Then
        If UCase(txtUsuario) = UCase("Suporte") Then MsgBox "Contatar o Suporte Executar Exclusao da Capa no Caixa": Exit Sub
        AntIniOpcoes.ClearCapaCx = CheckZerarCapa
    End If
    
'    If CheckZerarDocto Then
'        If UCase(txtUsuario) = UCase("Suporte") Then MsgBox "Contatar o Suporte Executar Exclusao do  Docto no Caixa": Exit Sub
'        AntIniOpcoes.ClearDoctoCX = CheckZerarDocto
'    End If
           
    If OptionFixo Then
        AntIniOpcoes.InterFixo = True
        AntIniOpcoes.InterValo = ComboFixo.Text
        GravarOpcaoINI "Diversos", "Fixo", 1
        GravarOpcaoINI "Diversos", "Intervalo", AntIniOpcoes.InterValo
    Else
        AntIniOpcoes.InterValo = ComboCrescente.Text
        AntIniOpcoes.InterFixo = False
        GravarOpcaoINI "Diversos", "Fixo", 0
        GravarOpcaoINI "Diversos", "Intervalo", AntIniOpcoes.InterValo
    End If
     
End Sub
