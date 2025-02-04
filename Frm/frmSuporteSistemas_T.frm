VERSION 5.00
Begin VB.Form frmSuporteSistemas 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Suporte Simac - Sugestões de Melhorias"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSuporteSistemas_T.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPropostaResultado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4080
      Width           =   7335
   End
   Begin VB.ComboBox cboLocal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.ComboBox cboAtendente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "frmSuporteSistemas_T.frx":12CA
      Left            =   2520
      List            =   "frmSuporteSistemas_T.frx":12CC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1035
      Width           =   4695
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdCadastrar 
      Appearance      =   0  'Flat
      Caption         =   "&Cadastrar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox txtObservacao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1920
      Width           =   7335
   End
   Begin VB.ComboBox cboElemento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
      Width           =   4695
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label vCaracteres2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6720
      TabIndex        =   15
      Top             =   3840
      Width           =   480
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Propostas de Melhorias e Resultado Esperado:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   4515
   End
   Begin VB.Label Local 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   510
   End
   Begin VB.Label lblAtendente 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Atendente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2520
      TabIndex        =   12
      Top             =   795
      Width           =   1005
   End
   Begin VB.Label vCaracteres 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6720
      TabIndex        =   11
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label lblDescricaoServico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Descreva Objetivamente o Processo Atual:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   4170
   End
   Begin VB.Label Elemento 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Elemento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblDivisao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmSuporteSistemas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private strRelatorio As String

Private Sub suListarTipo()
Dim rs As New ADODB.Recordset
Call ConectarBD
    strSQL = "SELECT * FROM tb_Tipo"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboTipo.AddItem rs!Tipo
        rs.MoveNext
    Loop
    Set rs = Nothing
  
End Sub

Private Sub suListarElemento()
    strSQL = "SELECT * FROM tb_Elemento"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboElemento.AddItem rs!Elemento
        rs.MoveNext
    Loop
    Set rs = Nothing
    
End Sub

Private Sub suListarLocal()
    strSQL = "SELECT * FROM tb_Local"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboLocal.AddItem rs!Local
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Private Sub suListarAtendimento()
    strSQL = "SELECT * FROM tb_Atendente"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboAtendente.AddItem rs!Atendente
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Private Sub cmdCadastrar_Click()
    If Len(Trim(Me.cboTipo)) = 0 Then
        MsgBox "Selecione o tipo!", vbExclamation, "Suporte Simac"
        Me.cboTipo.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.cboLocal)) = 0 Then
        MsgBox " Selecione o Local!", vbExclamation, "Suporte Simac"
        Me.cboLocal.SetFocus
        Exit Sub
        
    End If
    If Len(Trim(Me.cboAtendente)) = 0 Then
        MsgBox "Selecione o atendendte!", vbExclamation, "Suporte Simac"
        Me.cboAtendente.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Me.txtObservacao)) = 0 Then
        MsgBox "Por favor, descreva a situação atual.", vbExclamation, "Descrição da Situação Atual"
        Me.txtObservacao.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Me.txtPropostaResultado)) = 0 Then
        MsgBox "Por favor, descreva sua proposta de melhoria e o resultado esperado.", vbExclamation, "Proposta de Melhoria e Resultado Esperado"
        Me.txtPropostaResultado.SetFocus
        Exit Sub
    End If
    
    If fnCadastrarOS(Me.cboTipo, Me.cboLocal, Me.cboAtendente, Me.txtObservacao, gstrNome, gstrDepto, gstrEMail, gintUsuarioID, Me.txtPropostaResultado, sFullName, strDescription) = True Then
        MsgBox "Atendimento N° " & Format(gintOSID, "0000") & " gerada com sucesso!", vbOKOnly + vbInformation, "Suporte Simac"
        Call LimparDados
        Call Unload(Me)
    End If

End Sub
Private Function fnCadastrarOS(ByVal vTipo As String, ByVal vLocal As String, ByVal vAtendente As String, ByVal vObservacao As String, vUsuario As String, ByVal vDepartamento As String, ByVal vEMail As String, ByVal vUsuarioID, ByVal vDescricao2, ByVal vUsuarioAD As String, ByVal vDepartamentoAD As String) As Boolean
On Error GoTo Erro
    fnCadastrarOS = False
    gintOSID = 0
    
    strSQL = "SELECT * FROM tb_Atendimento WHERE 1=2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = True Then
        rs.AddNew
    
        rs!Tipo = vTipo
        rs!Local = vLocal
        rs!Atendente = vAtendente
        rs!DescricaoMelhorias = vObservacao
        If Len(vUsuario) > 0 Then
            rs!Nome = vUsuario
        Else
            rs!Nome = vUsuarioAD
        End If
        If Len(vDepartamento) > 0 Then
            rs!Departamento = vDepartamento
        Else
            rs!Departamento = vDepartamentoAD
        End If
        rs!Email = vEMail
        rs!usuarioID = vUsuarioID
        rs!PropostaResultado = vDescricao2
        rs!DataCadastro = Format(Now, "dd/mm/yyyy HH:mm")
        rs!Status = 0
 
        rs.Update
 
        gintOSID = rs!ATID
        fnCadastrarOS = True
        
        

        
        If fnEnviarEmail(gintOSID, gstrNome, sFullName) = False Then
            MsgBox "Erro ao enviar o E-mail", vbOKOnly + vbCritical, "Suporte Simac"
            Exit Function
        End If
    End If
    
    Set rs = Nothing
    Exit Function
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Simac"
    Set rs = Nothing
End Function

Private Sub cmdFechar_Click()
Call Unload(Me)
End Sub

Private Sub Form_Load()


Me.Elemento.Enabled = False
Me.cboElemento.Enabled = False

Call suListarTipo
Call suListarElemento
Call suListarLocal
Call suListarAtendimento
Me.cboAtendente.ListIndex = 0
Me.cboTipo.ListIndex = 0

End Sub

Private Sub Text1_Change()

End Sub
Private Sub LimparDados()
Me.cboTipo.ListIndex = -1
Me.cboLocal.ListIndex = -1
Me.cboAtendente.ListIndex = -1
Me.txtObservacao = ""
Me.txtPropostaResultado = ""
End Sub


Private Sub vCaracteres2_Change()
If Me.vCaracteres2.Caption = 0 Then
    Me.vCaracteres2.ForeColor = vbRed
Else
    Me.vCaracteres2.ForeColor = vbBlue
End If
End Sub

Private Sub vCaracteres_Change()
If Me.vCaracteres.Caption = 0 Then
    Me.vCaracteres.ForeColor = vbRed
Else
    Me.vCaracteres.ForeColor = vbBlue
End If
End Sub
Private Sub txtPropostaResultado_Change()
Me.vCaracteres2.Caption = 1000 - Len(Me.txtPropostaResultado.Text)
End Sub
Private Sub txtObservacao_Change()
    Me.vCaracteres.Caption = 1000 - Len(Me.txtObservacao.Text)
End Sub



Private Function fnEnviarEmail(ByVal vOSID As Long, ByVal vUsuario As String, vUsuarioAD) As Boolean
On Error GoTo Erro
Dim poSendMail As vbSendMail.clsSendMail
    fnEnviarEmail = False
    Me.Enabled = False
    Screen.MousePointer = 11

    DoEvents
    
    Set poSendMail = New vbSendMail.clsSendMail
    poSendMail.SMTPHost = "email-ssl.com.br"
    poSendMail.SMTPPort = "587"
    poSendMail.UseAuthentication = True
    poSendMail.Username = fnContaSMTP
    poSendMail.password = fnSenhaSMTP
    poSendMail.From = fnContaSMTP
    If Len(vUsuario) > 0 Then
        poSendMail.FromDisplayName = vUsuario
    Else
        poSendMail.FromDisplayName = vUsuarioAD
    End If
    poSendMail.Recipient = "simac_telecom@cablena.com.br"
    poSendMail.RecipientDisplayName = "ADM Suporte Simac"
    'poSendMail.CcRecipient = "emperes@cablena.com.br"
    poSendMail.Subject = "Suporte Simac Telecom - Novo Atendimento " & Format(vOSID, "0000")
    poSendMail.Priority = HIGH_PRIORITY
    
    Call suRelatorio(vOSID)
    poSendMail.Message = strRelatorio
    poSendMail.Send
    
    Set poSendMail = Nothing
    fnEnviarEmail = True
    Screen.MousePointer = 0
    Me.Enabled = True
    Exit Function
    
Erro:
    Screen.MousePointer = 0
    Me.Enabled = True
    MsgBox "Erro : " & Err.Description, vbOKOnly = vbCritical, "Suporte Simac"
    Set poSendMail = Nothing

End Function

Private Sub suRelatorio(ByVal vOSID As Long)
'On Error GoTo erro
    strSQL = "SELECT * FROM vw_Chamados WHERE ATID = " & vOSID
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        strRelatorio = ""
        strRelatorio = String(120, "=") & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Nº do Atendimento: " & Format(rs!ATID, "0000") & vbCrLf & vbCrLf
        If Len(gstrNome) > 0 Then
            strRelatorio = strRelatorio & "Usúario: " & gstrNome & vbCrLf & vbCrLf
        Else
            strRelatorio = strRelatorio & "Usúario: " & sFullName & vbCrLf & vbCrLf
        End If
        strRelatorio = strRelatorio & "Tipo: " & rs!Tipo & String(10, " ") & "Local: " & rs!Local & String(10, " ") & "Elemento: " & rs!Elemento & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Descrição da Situação Atual: " & rs!DescricaoMelhorias & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Proposta de Melhoria e Resultado Esperado: " & rs!PropostaResultado & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Atendente: " & rs!Atendente & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & "Data Cadastro: " & Format(rs!DataCadastro, "dd/MM/yyyy HH:mm") & vbCrLf & vbCrLf
        strRelatorio = strRelatorio & String(120, "=")
    End If
    
    rs.Close
    Set rs = Nothing
'erro:
'    MsgBox "erro " & Err.Description, vbCritical, "Suporte Simac"
End Sub

Private Function fnContaSMTP() As String
    fnContaSMTP = ""
    strSQL = "SELECT * FROM tb_Email WHERE ID = 1"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnContaSMTP = Trim(rs!valor)
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function fnSenhaSMTP() As String
    fnSenhaSMTP = ""
    strSQL = "SELECT * FROM tb_Email WHERE ID = 2"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        fnSenhaSMTP = Trim(rs!valor)
    End If
    
    rs.Close
    Set rs = Nothing
End Function


