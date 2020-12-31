VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form AguardarRobo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Parâmetros"
   ClientHeight    =   1956
   ClientLeft      =   2844
   ClientTop       =   3372
   ClientWidth     =   3912
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1956
   ScaleWidth      =   3912
   Begin VB.Frame Frame1 
      Height          =   1476
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   3792
      Begin ComctlLib.ProgressBar ProgressBar 
         Height          =   372
         Left            =   120
         TabIndex        =   3
         Top             =   984
         Width           =   3552
         _ExtentX        =   6265
         _ExtentY        =   656
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblData 
         Caption         =   "10/10/2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   1164
         TabIndex        =   4
         Top             =   528
         Width           =   1416
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Aguardando Carga de Parâmetros..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   3960
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Default         =   -1  'True
      Height          =   372
      Left            =   1140
      TabIndex        =   0
      Top             =   1548
      Width           =   1572
   End
End
Attribute VB_Name = "AguardarRobo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cancelou As Boolean

Private Sub CmdSair_Click()
    Cancelou = True
End Sub

Private Sub Form_Activate()

    Dim qryCargaTabelas As rdoQuery, qryUpdateCargaTabelas As rdoQuery
    Dim tb As rdoResultset
    Dim bPrimeiro As Boolean

    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    lblData.Caption = Mid(Geral.DataProcessamento, 7, 2) + "/" + Mid(Geral.DataProcessamento, 5, 2) + "/" + Mid(Geral.DataProcessamento, 1, 4)
    On Error GoTo ErroActive
    
    Set qryCargaTabelas = Geral.Banco.CreateQuery("", "{call GetCargaTabelas (?)}")
    qryCargaTabelas.rdoParameters(0) = Geral.DataProcessamento
    bPrimeiro = True
    
    Do While True And Not Cancelou
        Set tb = qryCargaTabelas.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        If Not tb.EOF And Not IsNull(tb!CargaTabelas) Then
            If tb!CargaTabelas = "S" Then
                tb.Close
                Exit Do
            ElseIf bPrimeiro Then
                If GrupoUsuario(Geral.Usuario, eG_SUPORTE) _
                    Or GrupoUsuario(Geral.Usuario, eG_SUPERVISOR) _
                    Or GrupoUsuario(Geral.Usuario, eG_LIDER) Then
                'If UsuarioSuporte(Geral.Usuario) Then
                    If MsgBox("A inicialização do link ainda não foi completada." & vbCr & vbCr & "Deseja liberar sistema?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                        If MsgBox(Space(45) + "A T E N Ç Ã O" + Space(45) + vbCrLf + vbCrLf + _
                                    "Ao confirmar será criada nova data de movimento e alterada a data do vinculador !" + vbCrLf + vbCrLf + vbCrLf + _
                                    "Confirma liberação do sistema para esta nova data  (" + DataDD_MM_AAAA(Geral.DataProcessamento) + ") ?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then

                            Set qryUpdateCargaTabelas = Geral.Banco.CreateQuery("", "{? = call AtualizaCargaTabelas (?,?)}")
                            qryUpdateCargaTabelas.rdoParameters(1) = Geral.DataProcessamento
                            qryUpdateCargaTabelas.rdoParameters(2) = "S"
                            qryUpdateCargaTabelas.Execute
                            If (qryUpdateCargaTabelas.rdoParameters(0) <> 0) Then
                                MsgBox "Não foi possível atualizar parâmetros para esta data de movimento.", vbCritical + vbOKOnly, App.Title
                            Else
                                GravaLog 0, 0, 51
                            End If
                            qryUpdateCargaTabelas.Close
                        End If
                    End If
                End If
            End If
            bPrimeiro = False
        Else
            If InStr(1, Geral.StringConexao, "Backup", vbTextCompare) <> 0 Then
                MsgBox "Não existe esta Data de Movimento na base Backup.", _
                    vbExclamation + vbOKOnly, App.Title
                Cancelou = True
            Else
                If GrupoUsuario(Geral.Usuario, eG_SUPORTE) _
                    Or GrupoUsuario(Geral.Usuario, eG_SUPERVISOR) _
                    Or GrupoUsuario(Geral.Usuario, eG_LIDER) Then
                
                'If UsuarioSuporte(Geral.Usuario) Then
                    If bPrimeiro Then
                        If MsgBox("A inicialização do link ainda não foi completada." & vbCr & vbCr & "Deseja liberar sistema?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                        
                            If MsgBox(Space(45) + "A T E N Ç Ã O" + Space(45) + vbCrLf + vbCrLf + _
                                        "Ao confirmar será criada nova data de movimento e alterada a data do vinculador !" + vbCrLf + vbCrLf + vbCrLf + _
                                        "Confirma liberação do sistema para esta nova data  (" + DataDD_MM_AAAA(Geral.DataProcessamento) + ") ?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                                Geral.qryCriarParametro.rdoParameters(1) = Geral.DataProcessamento
                                Geral.qryCriarParametro.Execute
                                If (Geral.qryCriarParametro.rdoParameters(0) <> 0) Then
                                    MsgBox "Não foi possível criar parâmetros para esta data de movimento.", vbCritical + vbOKOnly, App.Title
                                Else
                                    GravaLog 0, 0, 50
                                End If
                                Geral.qryCriarParametro.Close
    
                                Set qryUpdateCargaTabelas = Geral.Banco.CreateQuery("", "{? = call AtualizaCargaTabelas (?,?)}")
                                qryUpdateCargaTabelas.rdoParameters(1) = Geral.DataProcessamento
                                qryUpdateCargaTabelas.rdoParameters(2) = "S"
                                qryUpdateCargaTabelas.Execute
                                If (qryUpdateCargaTabelas.rdoParameters(0) <> 0) Then
                                    MsgBox "Não foi possível atualizar parâmetros para esta data de movimento.", vbCritical + vbOKOnly, App.Title
                                Else
                                    GravaLog 0, 0, 51
                                End If
                                qryUpdateCargaTabelas.Close
                            End If
                        End If
                    End If
                    bPrimeiro = False
                End If
            End If
        End If
        tb.Close
        Movimenta
        DoEvents
    Loop
    
    Me.Hide
    
    Exit Sub
    
ErroActive:

    Select Case TratamentoErro("Erro na obtenção da carga das Tabelas (Robô)", Err, rdoErrors)
        Case vbCancel
            End
        Case vbRetry
            Resume
    End Select


End Sub

Private Sub Form_Load()
    Cancelou = False
    ProgressBar.Value = 0
End Sub

Private Sub Movimenta()
    Static i As Integer
    Dim Count As Integer

    If i >= 100 Then
        i = 0
    End If
    
    ProgressBar.Value = i
    
    i = i + 2
    
    For Count = 0 To 10000
        DoEvents
    Next
    
End Sub

Private Function UsuarioSuporte(ByVal User As String) As Boolean
    
    UsuarioSuporte = GrupoUsuario(User, eG_SUPORTE)
                            
End Function
