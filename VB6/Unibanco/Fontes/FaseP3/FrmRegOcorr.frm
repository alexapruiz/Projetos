VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRegOcorr 
   Caption         =   "Registro de Ocorrência"
   ClientHeight    =   6876
   ClientLeft      =   204
   ClientTop       =   1848
   ClientWidth     =   11688
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6876
   ScaleWidth      =   11688
   Begin VB.Frame Frame3 
      Height          =   1260
      Left            =   156
      TabIndex        =   15
      Top             =   0
      Width           =   9456
      Begin VB.TextBox TxtNumMalote 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   7116
         MaxLength       =   12
         TabIndex        =   1
         Top             =   228
         Width           =   2196
      End
      Begin VB.ComboBox CmbAgencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   396
         ItemData        =   "FrmRegOcorr.frx":0000
         Left            =   2280
         List            =   "FrmRegOcorr.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   732
         Width           =   2604
      End
      Begin VB.PictureBox Picture6 
         Height          =   396
         Left            =   108
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   732
         Width           =   2100
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Agência"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   -48
            TabIndex        =   21
            Top             =   12
            Width           =   984
         End
      End
      Begin VB.ComboBox cmbCapa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   396
         Left            =   2280
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2604
      End
      Begin VB.PictureBox Picture4 
         Height          =   396
         Left            =   108
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   2100
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Capa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   12
            TabIndex        =   19
            Top             =   12
            Width           =   1992
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   396
         Left            =   4944
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   2100
         Begin VB.Label Label3 
            Caption         =   "Número do Malote"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   252
            Left            =   36
            TabIndex        =   17
            Top             =   36
            Width           =   1956
         End
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   396
         Left            =   4944
         TabIndex        =   23
         Top             =   732
         Width           =   2100
      End
      Begin VB.Label lblLote 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   7128
         TabIndex        =   22
         Top             =   732
         Width           =   2196
      End
   End
   Begin TabDlg.SSTab TabTipOcorr 
      Height          =   5124
      Left            =   120
      TabIndex        =   8
      Top             =   1632
      Width           =   11436
      _ExtentX        =   20172
      _ExtentY        =   9038
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Envelope/Malote"
      TabPicture(0)   =   "FrmRegOcorr.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListOcorrencias1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Depósito"
      TabPicture(1)   =   "FrmRegOcorr.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListOcorrencias2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Pagamento"
      TabPicture(2)   =   "FrmRegOcorr.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListOcorrencias3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Di&versos"
      TabPicture(3)   =   "FrmRegOcorr.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ListOcorrencias4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&Aut. Débito"
      TabPicture(4)   =   "FrmRegOcorr.frx":0074
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ListOcorrencias5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Transf. Valor"
      TabPicture(5)   =   "FrmRegOcorr.frx":0090
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ListOcorrencias6"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&Fininvest"
      TabPicture(6)   =   "FrmRegOcorr.frx":00AC
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ListOcorrencias7"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "&Operacional"
      TabPicture(7)   =   "FrmRegOcorr.frx":00C8
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "ListOcorrencias8"
      Tab(7).ControlCount=   1
      Begin VB.ListBox ListOcorrencias8 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4536
         ItemData        =   "FrmRegOcorr.frx":00E4
         Left            =   -74640
         List            =   "FrmRegOcorr.frx":00EB
         TabIndex        =   25
         Top             =   456
         Width           =   10728
      End
      Begin VB.ListBox ListOcorrencias7 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4536
         ItemData        =   "FrmRegOcorr.frx":0101
         Left            =   -74640
         List            =   "FrmRegOcorr.frx":0108
         TabIndex        =   24
         Top             =   456
         Width           =   10728
      End
      Begin VB.ListBox ListOcorrencias6 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4536
         ItemData        =   "FrmRegOcorr.frx":011E
         Left            =   -74640
         List            =   "FrmRegOcorr.frx":0125
         TabIndex        =   14
         Top             =   456
         Width           =   10728
      End
      Begin VB.ListBox ListOcorrencias5 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4536
         ItemData        =   "FrmRegOcorr.frx":013B
         Left            =   -74640
         List            =   "FrmRegOcorr.frx":0142
         TabIndex        =   13
         Top             =   456
         Width           =   10728
      End
      Begin VB.ListBox ListOcorrencias4 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4536
         ItemData        =   "FrmRegOcorr.frx":0158
         Left            =   -74640
         List            =   "FrmRegOcorr.frx":015F
         TabIndex        =   12
         Top             =   456
         Width           =   10728
      End
      Begin VB.ListBox ListOcorrencias3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4536
         ItemData        =   "FrmRegOcorr.frx":0175
         Left            =   -74640
         List            =   "FrmRegOcorr.frx":017C
         TabIndex        =   11
         Top             =   456
         Width           =   10728
      End
      Begin VB.ListBox ListOcorrencias2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4536
         ItemData        =   "FrmRegOcorr.frx":0192
         Left            =   -74640
         List            =   "FrmRegOcorr.frx":0199
         TabIndex        =   10
         Top             =   456
         Width           =   10728
      End
      Begin VB.ListBox ListOcorrencias1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4536
         ItemData        =   "FrmRegOcorr.frx":01AF
         Left            =   360
         List            =   "FrmRegOcorr.frx":01B6
         TabIndex        =   9
         Top             =   456
         Width           =   10728
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1260
      Left            =   9816
      TabIndex        =   7
      Top             =   0
      Width           =   1752
      Begin VB.CommandButton CmdExec 
         Caption         =   "&Confirma"
         Height          =   324
         Left            =   144
         TabIndex        =   4
         Top             =   528
         Width           =   1464
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Height          =   324
         Left            =   144
         TabIndex        =   3
         Top             =   180
         Width           =   1464
      End
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   324
         Left            =   144
         TabIndex        =   5
         Top             =   876
         Width           =   1464
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Ocorrências"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   336
      Left            =   -48
      TabIndex        =   6
      Top             =   1320
      Width           =   2664
   End
End
Attribute VB_Name = "FrmRegOcorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsAux               As rdoResultset
Dim RsStatus            As rdoResultset
Dim TEnvMal             As String
Dim OcorrIndex          As Integer
Dim lIdCapa             As Long
Dim Status              As String
Dim Situacao            As Integer
Dim ErrorOcorr          As Integer
Dim m_bIsEvent          As Boolean

Private Type MdregOcorrencia
    qryGetLstOcorrencias As rdoQuery
    qryGetMaloteRegOcorr As rdoQuery
    qryGetMudaStatus     As rdoQuery
    qryGetStatus         As rdoQuery
End Type

Private Type MdSupExclusao
    qryGetCapaRegOcorr   As rdoQuery
    qryGetAgCapaRegOcorr As rdoQuery
    qryGetCapaRegOcorr1  As rdoQuery
End Type

Private MdregOcorrencia  As MdregOcorrencia
Private MdSupExclusao    As MdSupExclusao
Private Function ImprimeHeaderOcorrencia() As Boolean
    
    Dim ret_imp, ret_aut As Integer
    Dim Buff_st As String * 3
    Dim buff_aut As String * 45
    Dim buff_linha As String * 2
    Dim r1, r2, r3 As Integer
     
    ImprimeHeaderOcorrencia = False
    
    buff_linha = Chr(13)
    
    buff_aut = Space(16) & "OCORRENCIAS"
    
    If Geral.autenticadora = 2 Then
        ret_imp = Autentica.Status(Buff_st)
        r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
        r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
        r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
        If (r1) <> 0 Then
            ret_aut = 1
        Else
            ret_aut = Autentica.Imprimir(buff_aut, False)
        End If
    Else
        ret_aut = Autentica.Imprimir(buff_aut, False)
    End If
    
    If (ret_aut <> 0) Then
        ret_imp = Autentica.Status(Buff_st)
    
        If (ret_imp = 0) Then
           
           r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
           r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
           r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
           
           If (r1 + r2 + r3) <> 0 Then
              'teste do 1 byte
              Select Case r1
                 Case 1
                    MsgBox "Autenticadora está off-line.", vbInformation + vbOKOnly, App.Title
                 Case 2
                    MsgBox "Autenticadora está desligada.", vbInformation + vbOKOnly, App.Title
                 Case 3
                    MsgBox "Autenticadora está com buffer cheio.", vbInformation + vbOKOnly, App.Title
                 Case 4
                    MsgBox "Autenticadora está inoperante.", vbInformation + vbOKOnly, App.Title
              End Select
           
              'teste do 3 byte
              If (r3 <> 0) Then
                 MsgBox "Verifique a bobina da Autenticadora.", vbInformation + vbOKOnly, App.Title
              End If
              Exit Function
           End If
        Else
           MsgBox "Verifique a Autenticadora.", vbInformation + vbOKOnly, App.Title
           Exit Function
        End If
    Else
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
        If TEnvMal = "E" Then
           buff_aut = Space(12) & "Envelope - " & Format(cmbCapa, "00000000")
           ret_aut = Autentica.Imprimir(buff_aut, False)
           ret_aut = Autentica.Imprimir(buff_linha, False)
           
           buff_aut = String(45, "=")
           ret_aut = Autentica.Imprimir(buff_aut, False)
           ret_aut = Autentica.Imprimir(buff_linha, False)
        End If
        
        If TEnvMal = "F" Then
           buff_aut = Space(6) & "Envelope Fininvest - " & Format(cmbCapa, "0000000000")
           ret_aut = Autentica.Imprimir(buff_aut, False)
           ret_aut = Autentica.Imprimir(buff_linha, False)
           
           buff_aut = String(45, "=")
           ret_aut = Autentica.Imprimir(buff_aut, False)
           ret_aut = Autentica.Imprimir(buff_linha, False)
        End If
        
        
        If TEnvMal = "M" Then
            buff_aut = Space(4) & "Capa Malote Empresa - " & Format(cmbCapa, "00000000000000")
            buff_aut = Space(4) & "Nº   Malote Empresa - " & FormataMalote(TxtNumMalote)
            ret_aut = Autentica.Imprimir(buff_aut, False)
            ret_aut = Autentica.Imprimir(buff_linha, False)
            
            buff_aut = String(45, "=")
            ret_aut = Autentica.Imprimir(buff_aut, False)
            ret_aut = Autentica.Imprimir(buff_linha, False)
        End If
        
        buff_aut = Space(7) & "Data do Movimento: " + Mid(Geral.DataProcessamento, 7, 2) + "/" + Mid(Geral.DataProcessamento, 5, 2) + "/" + Mid(Geral.DataProcessamento, 1, 4)
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = Space(4) & "Data de Emissão: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "  Ag. Coleta: " + Format(CmbAgencia, "0000") & " - Ag. Processadora: " + Geral.AgenciaCentral
        ret_aut = Autentica.Imprimir(buff_aut, False)
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
        buff_aut = String(45, "-")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
    End If
    
    Call ImprimeOcorrenciaCapa
    
    ImprimeHeaderOcorrencia = True
    
End Function
Private Sub ImprimeOcorrenciaCapa()
    
    Dim ret_aut As Integer
    Dim StrMotivo As String
    Dim buff_aut As String * 45
    Dim buff_linha As String * 2
    Dim Pos As Integer
    
    buff_linha = Chr(13)
    
    If Status = "P" Then
        If ListOcorrencias1.Text <> "" Then StrMotivo = Mid(ListOcorrencias1.Text, 7, 200)
        If ListOcorrencias2.Text <> "" Then StrMotivo = Mid(ListOcorrencias2.Text, 7, 200)
        If ListOcorrencias3.Text <> "" Then StrMotivo = Mid(ListOcorrencias3.Text, 7, 200)
        If ListOcorrencias4.Text <> "" Then StrMotivo = Mid(ListOcorrencias4.Text, 7, 200)
        If ListOcorrencias5.Text <> "" Then StrMotivo = Mid(ListOcorrencias5.Text, 7, 200)
        If ListOcorrencias6.Text <> "" Then StrMotivo = Mid(ListOcorrencias6.Text, 7, 200)
        If ListOcorrencias7.Text <> "" Then StrMotivo = Mid(ListOcorrencias7.Text, 7, 200)
    End If
    
    If TEnvMal = "E" Then
        buff_aut = "Envelope Devolvido"
    ElseIf TEnvMal = "M" Then
        buff_aut = "Malote Devolvido"
    Else
        buff_aut = "Envelope Fininvest Devolvido"     'Envelope Fininvest
    End If
    
    ret_aut = Autentica.Imprimir(buff_aut, False)
    ret_aut = Autentica.Imprimir(buff_linha, False)
    
    buff_aut = "Motivo devolucao: "
    ret_aut = Autentica.Imprimir(buff_aut, False)
    
    If Len(StrMotivo) < 45 Then
        buff_aut = StrMotivo
        ret_aut = Autentica.Imprimir(buff_aut, False)
    Else
        Pos = 45
        While Pos < Len(StrMotivo)
            buff_aut = QuebraBuffer(StrMotivo, Pos)
            ret_aut = Autentica.Imprimir(buff_aut, False)
            StrMotivo = Right(StrMotivo, Len(StrMotivo) - Pos)
        Wend
        If Len(StrMotivo) > 0 Then
            buff_aut = StrMotivo
            ret_aut = Autentica.Imprimir(buff_aut, False)
        End If
    End If
    ret_aut = Autentica.Imprimir(buff_linha, False)
    
    ImprimeTrailler (False)
    
End Sub
Private Function QuebraBuffer(ByVal Buf As String, ByRef Pos As Integer) As String
    Dim Tam As Integer
    
    Tam = Pos
    Do While Tam > 0
        If Mid(Buf, Tam, 1) = " " Then
            Exit Do
        End If
        Tam = Tam - 1
    Loop
    If Tam > 0 Then
        Pos = Tam
    End If
    QuebraBuffer = Mid(Buf, 1, Pos)
End Function
Private Function ImprimeTrailler(ByVal ShowMsg As Boolean) As Boolean
    Dim ret_imp, ret_aut As Integer
    Dim Buff_st As String * 3
    Dim buff_aut As String * 45
    Dim buff_linha As String * 2
    Dim r1, r2, r3, i As Integer
     
    ImprimeTrailler = False
    
    buff_linha = Chr(13)
    
    buff_aut = String(45, "-")
    
    ret_aut = Autentica.Imprimir(buff_aut, False)
    
    If (ret_aut <> 0) Then
        ret_imp = Autentica.Status()
    
        If (ret_imp = 0) Then
           
           r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
           r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
           r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
           
           If (r1 + r2 + r3) <> 0 Then
              'teste do 1 byte
              Select Case r1
                 Case 1
                    MsgBox "Autenticadora está off-line.", vbInformation + vbOKOnly, App.Title
                 Case 2
                    MsgBox "Autenticadora está desligada.", vbInformation + vbOKOnly, App.Title
                 Case 3
                    MsgBox "Autenticadora está com buffer cheio.", vbInformation + vbOKOnly, App.Title
                 Case 4
                    MsgBox "Autenticadora está inoperante.", vbInformation + vbOKOnly, App.Title
              End Select
           
              'teste do 3 byte
              If (r3 <> 0) Then
                 MsgBox "Verifique a bobina da Autenticadora.", vbInformation + vbOKOnly, App.Title
              End If
              Exit Function
           End If
        Else
           MsgBox "Verifique a Autenticadora.", vbInformation + vbOKOnly, App.Title
           Exit Function
        End If
    Else
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
        If ShowMsg Then
            buff_aut = Space(9) & "Ticket de Caixa Unibanco."
            ret_aut = Autentica.Imprimir(buff_aut, False)
            ret_aut = Autentica.Imprimir(buff_linha, False)
            
            buff_aut = Space(3) & "Feito para facilitar o seu dia-a-dia."
            ret_aut = Autentica.Imprimir(buff_aut, False)
            ret_aut = Autentica.Imprimir(buff_linha, False)
        End If
        
        'Imprime 10 linhas no final da impressão do ticket
        For i = 1 To 10
            ret_aut = Autentica.Imprimir(buff_linha, False)
        Next i
        
    End If
    
    ImprimeTrailler = True
    
End Function
Private Sub limpa_Header()
    lblLote.Caption = ""
    cmbCapa.Clear
    TxtNumMalote.Text = ""
    CmbAgencia.Clear
End Sub
Private Sub cmdLimpaCampos_Click()
    LimpaTela Me
    txtNumEnvMal.SetFocus
End Sub
Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub cbmCapa_Change()

End Sub

Private Sub cmbAgencia_Change()
    If CmbAgencia.ListCount = 1 Then
        CmbAgencia.Text = CmbAgencia.List(0)
    End If
End Sub
Private Sub cmbAgencia_Click()
'* Muda o Número de Lote - Quando for Selecionado outra Agência *'

    If CmbAgencia.Text = "" Or m_bIsEvent Then Exit Sub

    With MdSupExclusao.qryGetAgCapaRegOcorr
        .rdoParameters(0).Value = CDbl(cmbCapa)
        .rdoParameters(1).Value = Geral.DataProcessamento
        .rdoParameters(2).Value = CInt(CmbAgencia)
        .rdoParameters(3).Value = CmbAgencia.ItemData(CmbAgencia.ListIndex)
        Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If Not RsAux.EOF Then
        lblLote = Format(RsAux!IdLote, "0000-00000")
        If TEnvMal = "M" Then
            TxtNumMalote = FormataMalote(RsAux!Num_Malote)
            Call CmdExec_Click
        End If
    End If
    
End Sub
Private Sub cmbAgencia_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 13) Then
       If Len(CmbAgencia) > 0 Then
            Call CmdExec_Click
       End If

    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click
    End If

End Sub
Private Sub cmbCapa_Change()
    If Len(TxtNumMalote) <> 0 Then
        If Len(cmbCapa) <> 0 And cmbCapa.ListCount = 1 Then
            Call Pesquisa_Dados
        End If
    ElseIf Not IsNumeric(cmbCapa.Text) Then
        cmbCapa.Text = ""
    End If
End Sub
Private Sub cmbCapa_Click()
     Call Pesquisa_Dados
End Sub
Private Sub cmbCapa_KeyPress(KeyAscii As Integer)

    If KeyAscii <> vbKeyReturn Then
        SoNumero KeyAscii
    End If

    If (KeyAscii = vbKeyReturn) Then
        Call LimpaSelecaoOcorrencia
       If Len(cmbCapa) > 0 Then
          Call Pesquisa_Dados
       End If
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
       KeyAscii = 0
    Else
       'Não permitir a digitação de mais de 18 caracteres
       If Len(cmbCapa.Text) >= 18 And cmbCapa.SelLength = 0 And (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
          KeyAscii = 0
       End If
    End If
   
End Sub
Private Sub CmdExec_Click()
  
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Dim strDescricao  As String

Msg = "Registro possui Ocorrência, Deseja Alterá-la!"   ' Define Mensagem
Style = vbYesNo + vbInformation + vbDefaultButton2      ' Define buttons.
Title = App.Title                                       ' Define title.
Ctxt = 1000                                             ' Define topic

    'Valida Capa
    If Trim(cmbCapa.Text) = "" Then
        MsgBox "Campo Obrigatório não preenchido!", vbInformation, App.Title
        cmbCapa.SetFocus
        Exit Sub
    End If

    'Valida Agência
    If Trim(CmbAgencia.Text) = "" Then
        MsgBox "Campo Obrigatório não preenchido!", vbInformation, App.Title
        CmbAgencia.SetFocus
        Exit Sub
    End If
        
    'Valida ocorrência para Envelope Fininvest
    If TEnvMal = "F" Then
        If ListOcorrencias1.Text <> "" Or ListOcorrencias2.Text <> "" Or _
           ListOcorrencias3.Text <> "" Or ListOcorrencias4.Text <> "" Or _
           ListOcorrencias5.Text <> "" Or ListOcorrencias6.Text <> "" Or _
           ListOcorrencias8.Text <> "" Then
            MsgBox "Escolha uma das opções da pasta de ocorrências para Envelope Fininvest.", vbInformation, App.Title
            TabTipOcorr.Tab = 6
            TabTipOcorr.SetFocus
            Exit Sub
        End If
    Else
        If ListOcorrencias7.Text <> "" Then
            MsgBox "Escolha uma das opções das pastas de ocorrências válidas para " & IIf(tenv_mal = "M", "Malote", "Envelope") & ".", vbInformation, App.Title
            TabTipOcorr.Tab = 0
            TabTipOcorr.SetFocus
            Exit Sub
        End If
    
    End If
        
    'Parametros para exeção da query GetTodasCapas
     With MdSupExclusao.qryGetCapaRegOcorr1
         .rdoParameters(0).Value = Geral.DataProcessamento
         .rdoParameters(1).Value = CmbAgencia.ItemData(CmbAgencia.ListIndex)
         .rdoParameters(2).Value = TEnvMal
         Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
     End With
        
    Select Case Situacao
    
        Case 1
            'Se o resultado da pesquisa for maior que zero
            'Executa ações a seguir
            If Not RsAux.EOF Then
                With MdregOcorrencia.qryGetStatus
                    .rdoParameters(0).Value = Geral.DataProcessamento
                    .rdoParameters(1).Value = Null
                    .rdoParameters(2).Value = Null
                    .rdoParameters(3).Value = CDbl(cmbCapa)
                    Set RsStatus = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
                End With
        
                If RsAux.RowCount = 0 Then Exit Sub
                Dim a As String
                    If RsAux!Status = "0" Then
                        lIdCapa = RsAux!IdCapa
                        
'''                        frmComplRegOcorr.m_Descricao = ""
                        
                        Call MudaStatus
                        If ErrorOcorr = 2 Then
                            Situacao = 1
                            ErrorOcorr = 0
                        Exit Sub
                        End If

                        MsgBox "Ocorrência Efetuada com Sucesso!", vbInformation, App.Title
                    
                    ElseIf RsAux!Status = "P" Then
                        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
                        If Response = vbYes Then
                            OcorrIndex = RsAux!Ocorrencia
                            Call Posiciona_Ocorrencia
                            Situacao = 2
                        Else
                            OcorrIndex = RsAux!Ocorrencia
                            Call Posiciona_Ocorrencia
                            CmdExec.Enabled = False
                            Situacao = 3
                        End If
                    ElseIf RsAux!Status = "X" Or RsAux!Status = "F" Or RsAux!Status = "D" Then
                        MsgBox "Registro Excluido - " & RsStatus!Descricao, vbInformation, App.Title
                        Exit Sub
                    Else
                    
                    MsgBox "Este " & IIf(TEnvMal = "M", "Malote", "Envelope") & " não está disponível. Pode estar sendo tratado por outra estação ou já foi tratado.", vbInformation
                    CmbAgencia.Clear
                    
                    Exit Sub
                    RsStatus.Close
                    End If
            Else
            'Finaliza operação
            Exit Sub
            End If

        Case 2
        'Case 2 AlteraStatus
        lIdCapa = RsAux!IdCapa
        
        'Obtem Complemento da Ocorrência
        strDescricao = ""
'''        Call GravaComplementoOcorrencia(lIdCapa, "C", strDescricao)
'''        frmComplRegOcorr.m_Descricao = strDescricao
        
        Call MudaStatus

        MsgBox "Alteração Efetuada com Sucesso!", vbInformation, App.Title
    
        Case 3
    
    End Select
    
    
End Sub
Private Sub CmdFechar_Click()
    Unload Me
End Sub
Private Sub cmdLimpar_Click()
    
    CmbAgencia.Clear
    cmbCapa.Clear
    LimpaTela Me
    lblLote = ""
    Call Form_Load
    CmdExec.Enabled = True
    cmbCapa.SetFocus
    TabTipOcorr.Tab = 0
    
End Sub

Private Sub Form_Activate()

    'Posiciona o form no centro da tela
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
   
    'Inclusão de chamada a rotina AtualizaAtividade
     Call AtualizaAtividade(4)
     
End Sub

Private Sub Form_Load()

Dim Contocorr As Integer

    'Limpa Listas de Ocorrênicias
    ListOcorrencias1.Clear
    ListOcorrencias2.Clear
    ListOcorrencias3.Clear
    ListOcorrencias4.Clear
    ListOcorrencias5.Clear
    ListOcorrencias6.Clear
    ListOcorrencias7.Clear
    ListOcorrencias8.Clear
    
    'Tab Default = 0
    TabTipOcorr.Tab = 0
    
    'Traz Todas as Capas (Malote ou Envolpe) do Período Corrente
    Set MdSupExclusao.qryGetCapaRegOcorr = Geral.Banco.CreateQuery("", "{Call GetTodasCapas_OC(?,?,?,?)}")
    
    Set MdSupExclusao.qryGetCapaRegOcorr1 = Geral.Banco.CreateQuery("", "{Call GetTodasCapas(?,?,?)}")
    
    'Traz Todas as Agencias da Capa (Malote ou Envolpe) do Período Corrente
    Set MdregOcorrencia.qryGetMaloteRegOcorr = Geral.Banco.CreateQuery("", "{Call GetMaloteExpedicao_oc(?,?)}")
    
    'Traz Todas as Agencias da Capa (Malote ou Envolpe) do Período Corrente
    Set MdSupExclusao.qryGetAgCapaRegOcorr = Geral.Banco.CreateQuery("", "{Call GetAgenciasCapa(?,?,?,?)}")
    
    'Traz Todas as Capas (Malote ou Envolpe) do Período Corrente
    Set MdregOcorrencia.qryGetMudaStatus = Geral.Banco.CreateQuery("", "{Call GetMudaStatus(?,?,?,?)}")
    
    'Traz o Status Atual do Malote
    Set MdregOcorrencia.qryGetStatus = Geral.Banco.CreateQuery("", "{Call GetrecuperaStatus (?,?,?,?)}")
    
    'Traz Todas as Capas (Malote ou Envolpe) do Período Corrente
    Set MdregOcorrencia.qryGetLstOcorrencias = Geral.Banco.CreateQuery("", "{Call GetTodasOcorrencia}")

    'Rotina que Popula List Ocorrência
    With MdregOcorrencia.qryGetLstOcorrencias
        Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If Not RsAux.EOF Then
        
        For CountOcorr = 0 To RsAux.RowCount - 1
            ' Lista Ocorrências de 1 a 100
            If RsAux!Ocorrencia >= 1 And RsAux!Ocorrencia <= 100 Then
                ListOcorrencias1.AddItem Format(RsAux!Ocorrencia, "000") & " - " & RsAux!Descricao
            End If
            ' Lista Ocorrências de 101 a 200
            If RsAux!Ocorrencia >= 101 And RsAux!Ocorrencia <= 200 Then
                ListOcorrencias2.AddItem Format(RsAux!Ocorrencia, "000") & " - " & RsAux!Descricao
            End If
            ' Lista Ocorrências de 201 a 300
            If RsAux!Ocorrencia >= 201 And RsAux!Ocorrencia <= 300 Then
                ListOcorrencias3.AddItem Format(RsAux!Ocorrencia, "000") & " - " & RsAux!Descricao
            End If
            ' Lista Ocorrências de 301 a 400
            If RsAux!Ocorrencia >= 301 And RsAux!Ocorrencia <= 400 Then
                ListOcorrencias4.AddItem Format(RsAux!Ocorrencia, "000") & " - " & RsAux!Descricao
            End If
            ' Lista Ocorrências de 401 a 500
            If RsAux!Ocorrencia >= 401 And RsAux!Ocorrencia <= 500 Then
                ListOcorrencias5.AddItem Format(RsAux!Ocorrencia, "000") & " - " & RsAux!Descricao
            End If
            
            ' Lista Ocorrências de 501 a 599
            If RsAux!Ocorrencia >= 501 And RsAux!Ocorrencia <= 599 Then
                ListOcorrencias6.AddItem Format(RsAux!Ocorrencia, "000") & " - " & RsAux!Descricao
            End If
            
            ' Lista Ocorrências de 600 a 699
            If RsAux!Ocorrencia >= 600 And RsAux!Ocorrencia <= 699 Then
                ListOcorrencias7.AddItem Format(RsAux!Ocorrencia, "000") & " - " & RsAux!Descricao
            End If
            
            'Move Para o próximo registro
            RsAux.MoveNext
        Next
            ListOcorrencias8.AddItem "999 - Erro Operacional"
    End If
    ' Situação Default de Ocorrência
    Situacao = 1

    
End Sub
Private Sub OptCapaEnvelope_Click()
    'Posiciona o foco no Text - txtnumEnvmal
     txtNumEnvMal.SetFocus
End Sub
Private Sub OptCapaMalote_Click()
    'Posiciona o foco no Text - txtnumEnvmal
     txtNumEnvMal.SetFocus
End Sub
Private Sub TxtNumEnvMal_LostFocus()

    If OptCapaEnvelope.Value Then
        txtNumEnvMal = Format(txtNumEnvMal, "00000000")
    Else
        txtNumEnvMal = Format(txtNumEnvMal, "00000000000000")
    End If

End Sub
Private Sub LstOcorrencias_DblClick()
    Call CmdExec_Click
End Sub

Private Sub ListOcorrencias2_DblClick()
    Call CmdExec_Click
End Sub
Private Sub ListOcorrencias3_DblClick()
    Call CmdExec_Click
End Sub
Private Sub ListOcorrencias4_DblClick()
    Call CmdExec_Click
End Sub
Private Sub ListOcorrencias5_DblClick()
    Call CmdExec_Click
End Sub
Private Sub ListOcorrencias6_DblClick()
    Call CmdExec_Click
End Sub
Private Sub ListOcorrencias1_DblClick()
    Call CmdExec_Click
End Sub

Private Sub ListOcorrencias7_DblClick()
    Call CmdExec_Click
End Sub

Private Sub ListOcorrencias8_DblClick()
    Call CmdExec_Click
End Sub

Private Sub TabTipOcorr_Click(PreviousTab As Integer)

    If ListOcorrencias1.ListCount = 0 Then Exit Sub

    If Situacao = 1 Or Situacao = 2 Then
        'Tab Envelope/Malote
        If TabTipOcorr.Tab = 0 Then
            ListOcorrencias1.Selected(0) = True
            ListOcorrencias2 = ListOcorrencias2.ListIndex - 1
            ListOcorrencias3 = ListOcorrencias3.ListIndex - 1
            ListOcorrencias4 = ListOcorrencias4.ListIndex - 1
            ListOcorrencias5 = ListOcorrencias5.ListIndex - 1
            ListOcorrencias6 = ListOcorrencias6.ListIndex - 1
            ListOcorrencias7 = ListOcorrencias7.ListIndex - 1
            ListOcorrencias8 = ListOcorrencias8.ListIndex - 1
        'Tab Depósito
        ElseIf TabTipOcorr.Tab = 1 Then
            ListOcorrencias2.Selected(0) = True
            ListOcorrencias1 = ListOcorrencias1.ListIndex - 1
            ListOcorrencias3 = ListOcorrencias3.ListIndex - 1
            ListOcorrencias4 = ListOcorrencias4.ListIndex - 1
            ListOcorrencias5 = ListOcorrencias5.ListIndex - 1
            ListOcorrencias6 = ListOcorrencias6.ListIndex - 1
            ListOcorrencias7 = ListOcorrencias7.ListIndex - 1
            ListOcorrencias8 = ListOcorrencias8.ListIndex - 1
        'Tab Pagamento
        ElseIf TabTipOcorr.Tab = 2 Then
            ListOcorrencias3.Selected(0) = True
            ListOcorrencias1 = ListOcorrencias1.ListIndex - 1
            ListOcorrencias2 = ListOcorrencias2.ListIndex - 1
            ListOcorrencias4 = ListOcorrencias4.ListIndex - 1
            ListOcorrencias5 = ListOcorrencias5.ListIndex - 1
            ListOcorrencias6 = ListOcorrencias6.ListIndex - 1
            ListOcorrencias7 = ListOcorrencias7.ListIndex - 1
            ListOcorrencias8 = ListOcorrencias8.ListIndex - 1
        'Tab Diversos
        ElseIf TabTipOcorr.Tab = 3 Then
            ListOcorrencias4.Selected(0) = True
            ListOcorrencias1 = ListOcorrencias1.ListIndex - 1
            ListOcorrencias2 = ListOcorrencias2.ListIndex - 1
            ListOcorrencias3 = ListOcorrencias3.ListIndex - 1
            ListOcorrencias5 = ListOcorrencias5.ListIndex - 1
            ListOcorrencias6 = ListOcorrencias6.ListIndex - 1
            ListOcorrencias7 = ListOcorrencias7.ListIndex - 1
            ListOcorrencias8 = ListOcorrencias8.ListIndex - 1
        'Tab Aut. Débito
        ElseIf TabTipOcorr.Tab = 4 Then
            ListOcorrencias5.Selected(0) = True
            ListOcorrencias1 = ListOcorrencias1.ListIndex - 1
            ListOcorrencias2 = ListOcorrencias2.ListIndex - 1
            ListOcorrencias3 = ListOcorrencias3.ListIndex - 1
            ListOcorrencias4 = ListOcorrencias4.ListIndex - 1
            ListOcorrencias6 = ListOcorrencias6.ListIndex - 1
            ListOcorrencias7 = ListOcorrencias7.ListIndex - 1
            ListOcorrencias8 = ListOcorrencias8.ListIndex - 1
        'Tab Transf. Valor
        ElseIf TabTipOcorr.Tab = 5 Then
            ListOcorrencias6.Selected(0) = True
            ListOcorrencias1 = ListOcorrencias1.ListIndex - 1
            ListOcorrencias2 = ListOcorrencias2.ListIndex - 1
            ListOcorrencias3 = ListOcorrencias3.ListIndex - 1
            ListOcorrencias4 = ListOcorrencias4.ListIndex - 1
            ListOcorrencias5 = ListOcorrencias5.ListIndex - 1
            ListOcorrencias7 = ListOcorrencias7.ListIndex - 1
            ListOcorrencias8 = ListOcorrencias8.ListIndex - 1
        'Tab Fininvest
        ElseIf TabTipOcorr.Tab = 6 Then
            If ListOcorrencias7.ListCount > 0 Then ListOcorrencias7.Selected(0) = True
            ListOcorrencias1 = ListOcorrencias1.ListIndex - 1
            ListOcorrencias2 = ListOcorrencias2.ListIndex - 1
            ListOcorrencias3 = ListOcorrencias3.ListIndex - 1
            ListOcorrencias4 = ListOcorrencias4.ListIndex - 1
            ListOcorrencias5 = ListOcorrencias5.ListIndex - 1
            ListOcorrencias6 = ListOcorrencias6.ListIndex - 1
            ListOcorrencias8 = ListOcorrencias8.ListIndex - 1
        'Tab Operacional
        ElseIf TabTipOcorr.Tab = 7 Then
            ListOcorrencias8.Selected(0) = True
            ListOcorrencias1 = ListOcorrencias1.ListIndex - 1
            ListOcorrencias2 = ListOcorrencias2.ListIndex - 1
            ListOcorrencias3 = ListOcorrencias3.ListIndex - 1
            ListOcorrencias4 = ListOcorrencias4.ListIndex - 1
            ListOcorrencias5 = ListOcorrencias5.ListIndex - 1
            ListOcorrencias6 = ListOcorrencias6.ListIndex - 1
            ListOcorrencias7 = ListOcorrencias7.ListIndex - 1
        End If
    End If
    
End Sub
Private Sub txtNumMalote_Change()

    If Not IsNumeric(TxtNumMalote.Text) Then
        TxtNumMalote.Text = ""
    End If
End Sub

Private Sub txtNumMalote_KeyPress(KeyAscii As Integer)

    If KeyAscii <> vbKeyReturn Then
        SoNumero KeyAscii
    End If

   If (KeyAscii = vbKeyReturn) Then
      If Len(TxtNumMalote) > 0 Then
         If VerificaMalote(TxtNumMalote) = False Then
            MsgBox "Número de Malote inválido.", vbInformation, App.Title
            TxtNumMalote.SetFocus
            Exit Sub
         End If
         Call ProcCapa
      End If
   ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   End If
   
End Sub
Public Sub Pesquisa_Dados()

    Dim CountAg As Integer
    Dim CountRegExcluido As Integer

    CmbAgencia.Clear

    If cmbCapa.Text = "" Then Exit Sub

    'Se Capa Não For diferente de Zero verifica seu dados
    With MdSupExclusao.qryGetAgCapaRegOcorr
        .rdoParameters(0).Value = CDbl(cmbCapa)
        .rdoParameters(1).Value = Geral.DataProcessamento
        .rdoParameters(2).Value = Null
        .rdoParameters(3).Value = Null
        Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If Not RsAux.EOF Then
        TEnvMal = RsAux!IdEnv_Mal
        '''''''''''''''''''''''''''''''
        'Loop Insere Dados de Agências'
        '''''''''''''''''''''''''''''''
        For CountAg = 0 To RsAux.RowCount - 1
            CmbAgencia.AddItem RsAux!AgOrig
            CmbAgencia.ItemData(CmbAgencia.NewIndex) = RsAux!IdCapa
            RsAux.MoveNext
        Next
        ''''''''''''''''''''''''''''''''''''''''''''
        'Abre Combobox Se ListCount For Maior que 1'
        ''''''''''''''''''''''''''''''''''''''''''''
        If CmbAgencia.ListCount = 1 Then
            m_bIsEvent = True
            CmbAgencia.Text = CmbAgencia.List(0)
            m_bIsEvent = False
            CmdExec_Click
        ElseIf CmbAgencia.ListCount > 1 Then
            CmbAgencia.SetFocus
            SendKeys "{F4}"
        End If
    Else
        MsgBox "Registro não Encontrado!", vbInformation, App.Title
        Exit Sub
    End If

End Sub
Public Sub MudaStatus()

    Status = "P" 'Status do titulo após receber ocorrência

    With MdregOcorrencia.qryGetMudaStatus
        .rdoParameters(0).Value = "P"
        
        If ListOcorrencias1.Text = "" And ListOcorrencias2.Text = "" And _
           ListOcorrencias3.Text = "" And ListOcorrencias4.Text = "" And _
           ListOcorrencias5.Text = "" And ListOcorrencias6.Text = "" And _
           ListOcorrencias7.Text = "" And ListOcorrencias8.Text = "" Then
            
            MsgBox "Selecione uma Ocorrência !", vbInformation, App.Title
            TabTipOcorr.Tab = 0
            ListOcorrencias1.Selected(0) = True
            ListOcorrencias1.SetFocus
            ErrorOcorr = 2
            Exit Sub
        End If
        
         If ListOcorrencias1.Text <> "" Then .rdoParameters(1).Value = Mid(ListOcorrencias1.Text, 1, 4)
         If ListOcorrencias2.Text <> "" Then .rdoParameters(1).Value = Mid(ListOcorrencias2.Text, 1, 4)
         If ListOcorrencias3.Text <> "" Then .rdoParameters(1).Value = Mid(ListOcorrencias3.Text, 1, 4)
         If ListOcorrencias4.Text <> "" Then .rdoParameters(1).Value = Mid(ListOcorrencias4.Text, 1, 4)
         If ListOcorrencias5.Text <> "" Then .rdoParameters(1).Value = Mid(ListOcorrencias5.Text, 1, 4)
         If ListOcorrencias6.Text <> "" Then .rdoParameters(1).Value = Mid(ListOcorrencias6.Text, 1, 4)
         If ListOcorrencias7.Text <> "" Then .rdoParameters(1).Value = Mid(ListOcorrencias7.Text, 1, 4)
         If ListOcorrencias8.Text <> "" Then .rdoParameters(1).Value = Mid(ListOcorrencias8.Text, 1, 4)
    
        .rdoParameters(2).Value = Geral.DataProcessamento
        .rdoParameters(3).Value = lIdCapa
        
        Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        
'''        frmComplRegOcorr.Show vbModal, Me
        'Grava/Altera ou Exclui Complemento da Ocorrência
'''        Call GravaComplementoOcorrencia(lIdCapa, IIf(frmComplRegOcorr.m_Descricao = "", "E", "G"), frmComplRegOcorr.m_Descricao)
        
        Call ImprimeHeaderOcorrencia 'Imprime ticket informando ocorrência
        Call cmdLimpar_Click         'Limpa Tela
                                     'Posiciona usuário quando Ok!
        Call GravaLog(lIdCapa, 0, 21)
                
    End With

End Sub
Public Sub Posiciona_Ocorrencia()

    '  Lista Ocorrências de 1 a 100
    '* Situação List 1
    If OcorrIndex >= 1 And OcorrIndex <= 100 Then
        For i = 0 To (ListOcorrencias1.ListCount) - 1
            If Val(Mid(ListOcorrencias1.List(i), 1, 3)) = Format(OcorrIndex, "000") Then
                TabTipOcorr.Tab = 0
                TabTipOcorr.TabEnabled(0) = True
                ListOcorrencias1.ListIndex = i
                ListOcorrencias1.SetFocus
                Exit For
                Exit Sub
             Else
                ListOcorrencias1.Selected(i) = False
             End If
        Next
    
    '  Lista Ocorrências de 101 a 200
    '* Situação List 2
    ElseIf OcorrIndex >= 101 And OcorrIndex <= 200 Then
        
        For i = 0 To (ListOcorrencias2.ListCount) - 1
            If Val(Mid(ListOcorrencias2.List(i), 1, 3)) = Format(OcorrIndex, "000") Then
                TabTipOcorr.Tab = 1
                TabTipOcorr.TabEnabled(1) = True
                ListOcorrencias2.ListIndex = i
                ListOcorrencias2.SetFocus
                Exit For
             End If
        Next
                 
    '  Lista Ocorrências de 201 a 300
    '* Situação List 3
    ElseIf OcorrIndex >= 201 And OcorrIndex <= 300 Then
         
        For i = 0 To ListOcorrencias3.ListCount - 1
            If Val(Mid(ListOcorrencias3.List(i), 1, 3)) = Format(OcorrIndex, "000") Then
                TabTipOcorr.Tab = 2
                TabTipOcorr.TabEnabled(2) = True
                ListOcorrencias3.ListIndex = i
                ListOcorrencias3.SetFocus
                Exit For
             End If
        Next
        
    '  Lista Ocorrências de 301 a 400
    '* Situação List 4
    ElseIf OcorrIndex >= 301 And OcorrIndex <= 400 Then
            
        For i = 0 To ListOcorrencias4.ListCount - 1
            If Val(Mid(ListOcorrencias4.List(i), 1, 3)) = Format(OcorrIndex, "000") Then
                TabTipOcorr.Tab = 3
                TabTipOcorr.TabEnabled(3) = True
                ListOcorrencias4.ListIndex = i
                ListOcorrencias4.SetFocus
                Exit For
                Exit Sub
             End If
        Next
        
    '  Lista Ocorrências de 401 a 500
    '* Situação List 5
    ElseIf OcorrIndex >= 401 And OcorrIndex <= 500 Then
                 
        For i = 0 To ListOcorrencias5.ListCount - 1
            If Val(Mid(ListOcorrencias5.List(i), 1, 3)) = Format(OcorrIndex, "000") Then
                TabTipOcorr.Tab = 4
                TabTipOcorr.TabEnabled(4) = True
                ListOcorrencias5.ListIndex = i
                ListOcorrencias5.SetFocus
                Exit For
                Exit Sub
             End If
        Next
        
    '  Lista Ocorrências de 501 a 599
    '* Situação List 6
    ElseIf OcorrIndex >= 501 And OcorrIndex <= 599 Then
        
        For i = 0 To ListOcorrencias6.ListCount - 1
            If Val(Mid(ListOcorrencias6.List(i), 1, 3)) = Format(OcorrIndex, "000") Then
                TabTipOcorr.Tab = 5
                TabTipOcorr.TabEnabled(5) = True
                ListOcorrencias6.ListIndex = i
                ListOcorrencias6.SetFocus
                Exit For
                Exit Sub
             End If
        Next
    
    '  Lista Ocorrências de 600 a 699
    '* Situação List 7
    ElseIf OcorrIndex >= 600 And OcorrIndex <= 699 Then
        
        For i = 0 To ListOcorrencias7.ListCount - 1
            If Val(Mid(ListOcorrencias7.List(i), 1, 3)) = Format(OcorrIndex, "000") Then
                TabTipOcorr.Tab = 6
                TabTipOcorr.TabEnabled(6) = True
                ListOcorrencias7.ListIndex = i
                ListOcorrencias7.SetFocus
                Exit For
                Exit Sub
             End If
        Next
    
    '  Lista Ocorrências de 700 a 1000
    '* Situação List 8
    ElseIf OcorrIndex >= 700 And OcorrIndex <= 1000 Then
        
        For i = 0 To ListOcorrencias8.ListCount - 1
            If Val(Mid(ListOcorrencias8.List(i), 1, 3)) = Format(OcorrIndex, "000") Then
                TabTipOcorr.Tab = 7
                TabTipOcorr.TabEnabled(7) = True
                ListOcorrencias8.ListIndex = i
                ListOcorrencias8.SetFocus
                Exit For
                Exit Sub
             End If
        Next
        
    End If

End Sub
Function VerificaMalote(ValNumMalote As String) As Boolean
'* Verifica se número de malote é Válido *'

    If Len(ValNumMalote) = 12 And CStr(Mid(ValNumMalote, 1, 2)) <> "09" Then
       VerificaMalote = False
    Else
       VerificaMalote = True
    End If
        
End Function
Public Sub ProcCapa()

On Error Resume Next

Dim CountCapaMalote As Integer
Dim CountRegExcluido As Integer

'Valida Texto - Digitado nº Malote
If Len(TxtNumMalote) >= 9 Then
   CmbAgencia.Clear
   cmbCapa.Clear
   RsAux.Close
      'Valida nº da Capa
      With MdregOcorrencia.qryGetMaloteRegOcorr
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = Val(FormataMalote(TxtNumMalote))
        If Err = 13 Then
            MsgBox "Valor Invalido, Reentre!", vbExclamation
            TxtNumMalote.Text = ""
            TxtNumMalote.SetFocus
            Exit Sub
        End If
        Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
      End With
      If RsAux.RowCount = 0 Then
          MsgBox "Registro não encontrado!", vbInformation, App.Title
          TxtNumMalote.SelStart = 0
          TxtNumMalote.SelLength = Len(TxtNumMalote)
          Exit Sub
      End If
      'Verifica Tipo de Status da Capa
      'Se For diferente de Zero devolve Mensagem
      'com a descrição do Status...
       With MdregOcorrencia.qryGetStatus
            .rdoParameters(0).Value = Geral.DataProcessamento
            .rdoParameters(1).Value = Val(FormataMalote(TxtNumMalote))
            .rdoParameters(2).Value = IIf(IsNull(RsAux!IdCapa), 0, RsAux!IdCapa)
            .rdoParameters(3).Value = Null
            Set RsStatus = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
       End With
       
       If RsStatus.RowCount <> 0 Then
          If RsStatus!Status <> "0" And RsStatus!Status <> "P" Then
             MsgBox "Registro possui Status :" & vbCr & RsStatus!Descricao, vbInformation, App.Title
             TxtNumMalote.Text = ""
             TxtNumMalote.SetFocus
             Exit Sub
        End If
        End If
       For CountCapaMalote = 0 To RsAux.RowCount - 1
           If RsAux!Status = "P" Or RsAux!Status = "E" Then
              CountRegExcluido = CountRegExcluido + 1
              If RsAux!Status = "P" Then
                cmbCapa.AddItem RsAux!Capa
                Call RetiraDuplicidade(RsAux!Capa)
              End If
           Else
              cmbCapa.AddItem RsAux!Capa
              Call RetiraDuplicidade(RsAux!Capa)
           End If
           RsAux.MoveNext
       Next
              
      If cmbCapa.ListCount = 1 Then
         cmbCapa.Text = cmbCapa.List(0)
      ElseIf cmbCapa.ListCount > 1 Then
         cmbCapa.SetFocus
         SendKeys "{F4}"
      End If
      
Else
    MsgBox "Dados inválidos, verifique!", vbInformation, App.TaskVisible
    TxtNumMalote.SelStart = 0
    TxtNumMalote.SelLength = Len(TxtNumMalote)
    TxtNumMalote.SetFocus
End If

End Sub
Function RetiraDuplicidade(NumCapaMalote As Double)
'* Elimina da Lista de Capas sua duplicidades *'

Dim CountLoop   As Integer  'Conta o Loop de acordo com a quantidade de Capas no Combo
Dim CountCapa   As Integer  'Traz  a quantidade de registros duplidados
Dim GuardaItem  As Integer

    For CountLoop = 0 To cmbCapa.ListCount - 1
        If NumCapaMalote = cmbCapa.List(CountLoop) Then
            CountCapa = CountCapa + 1
            GuardaItem = CountLoop
        End If
        
        If CountCapa >= 2 Then
            cmbCapa.RemoveItem (CountLoop)
        End If
    Next
      
End Function
Private Sub LimpaSelecaoOcorrencia()
    
    'Limpa Listas de Ocorrênicias
    ListOcorrencias1.Selected(0) = False
    ListOcorrencias2.Selected(0) = False
    ListOcorrencias3.Selected(0) = False
    ListOcorrencias4.Selected(0) = False
    ListOcorrencias5.Selected(0) = False
    ListOcorrencias6.Selected(0) = False
    If ListOcorrencias7.ListCount > 0 Then ListOcorrencias7.Selected(0) = False
    ListOcorrencias8.Selected(0) = False

End Sub
