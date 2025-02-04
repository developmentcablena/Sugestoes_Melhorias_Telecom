VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMensagem 
   Caption         =   "Reporte Tecnico"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   Icon            =   "frmMensagem_T.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboElemento2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1080
      Width           =   4455
   End
   Begin MSMask.MaskEdBox mkdHorario 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      BackColor       =   -2147483624
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtOS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   3
      Text            =   "         "
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4440
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5640
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtDecricao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1800
      Width           =   6975
   End
   Begin VB.Label Label4 
      Caption         =   "Elemento:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Data/ Hora."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Reporte Técnico:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Atendimento N° :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private rs As ADODB.Recordset
Private strSQL As String
Private strRelatorio As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Tempo()
Sleep 5000
End Sub
Private Sub cmdEnviar_Click()
Dim frm As New frmChamados
    If Len(Trim(Me.cboElemento2.Text)) = 0 Then
        MsgBox "Selecione um Elemento!", vbInformation, "Suporte Simac"
        Me.cboElemento2.SetFocus
      
        Exit Sub
    End If
    
    If Len(Trim(Me.txtDecricao)) = 0 Then
        MsgBox "Digite o Reporte Tecnico!", vbExclamation, "Suporte Simac"
        Me.txtDecricao.SetFocus
        Exit Sub
    End If

    If IsDate(Me.mkdHorario.Text) = False Then
        MsgBox "Data/ Hora de finalização de serviço!", vbOKOnly + vbExclamation, "Suporte Simac"
        Me.mkdHorario.SetFocus
        Exit Sub
    End If
    
    
    If fnReporteTecnico(Me.txtDecricao, Me.cboElemento2, CDate(Me.mkdHorario), gintOSID) = True Then
        MsgBox "Enviado com sucesso!", vbInformation, "Suporte Simac"
        Call LimparDados
        gintOSID = 0
     
        Exit Sub
    End If

End Sub
Private Sub LimparDados()
Me.cboElemento2.ListIndex = -1
Me.txtDecricao.Text = ""
Me.txtOS.Text = ""


End Sub

Private Sub Command1_Click()
Call Unload(Me)
End Sub

Private Sub Form_Load()

With Me
    .txtOS.Text = Format(gintOSID, "0000")
    .txtOS.Enabled = False
    .mkdHorario.Text = Format(Now, "dd/MM/yyyy HH:mm")
    
End With

Call suListarElemento
Me.mkdHorario.Mask = "##/##/#### ##:##"
End Sub

Private Sub suListarElemento()
    Call ConectarBD
    strSQL = "SELECT * FROM tb_Elemento"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboElemento2.AddItem rs!Elemento
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Private Function fnReporteTecnico(ByVal vDescricao As String, ByVal VElemento As String, ByVal vData As Date, ByVal vAtendimento As Integer) As Boolean
Call ConectarBD
On Error GoTo Erro

    fnReporteTenico = False
    Me.Enabled = False
    'Screen.MousePointer = 11
    'Call Tempo
    
    strSQL = "UPDATE tb_Atendimento SET " & _
    "ReporteTecnico = '" & vDescricao & "'," & _
    "Elemento = '" & VElemento & "', " & _
    "DataInicio = '" & Format(vData, "yyyy-MM-dd HH:mm") & "' " & _
    "WHERE ATID = " & vAtendimento & " "
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    fnReporteTecnico = True
    Me.Enabled = True
    'Screen.MousePointer = 0
    
    If fnEnviarEmail(VElemento, vAtendimento) = False Then
        MsgBox "Erro ao enviar o E-mail", vbOKOnly + vbCritical, "Suporte Simac"
        Exit Function
    End If

    Set rs = Nothing
    Exit Function
Erro:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Simac"
End Function

Private Function fnEnviarEmail(ByVal VElemento As String, ByVal vAtendimento As String) As Boolean
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
    poSendMail.FromDisplayName = "Kleber Alves da Silva"
    poSendMail.RecipientDisplayName = "ADM Suporte Simac Telecom"
    poSendMail.Subject = "Sugestões de Melhorias - Atendimento " & Format(vAtendimento, "0000")
    poSendMail.Priority = HIGH_PRIORITY
    
    If Trim(VElemento) = Trim("6.1 - Responsabilidade da direção") Then
        poSendMail.Recipient = "simac_telecom@cablena.com.br"
        poSendMail.CcRecipient = "cplavie@cablena.com.br"
        
    ElseIf VElemento = "6.2 - Manufatura de Fluxo" Then
        poSendMail.Recipient = "simac_telecom@cablena.com.br"
        poSendMail.CcRecipient = "cplavie@cablena.com.br;cmachado@cablena.com.br"
        
    ElseIf VElemento = "6.3 - Ambiente e envolvimento dos colaboradores" Then
        poSendMail.Recipient = "simac_telecom@cablena.com.br"
        poSendMail.CcRecipient = "rnelus@cablena.com.br;cplavie@cablena.com.br"
    
    ElseIf VElemento = "6.4 - Administração visual" Then
        poSendMail.Recipient = "simac_telecom@cablena.com.br"
        poSendMail.CcRecipient = "cmachado@cablena.com.br"
    
    ElseIf VElemento = "6.5 - Qualidade" Then
        poSendMail.Recipient = "simac_telecom@cablena.com.br"
        poSendMail.CcRecipient = "fsantos@cablena.com.br;r.venturar@cablena.com.br;wanderley.silva@cablena.com.br"
        
    ElseIf VElemento = "6.6 - Disponibilidade Operacional" Then
        poSendMail.Recipient = "simac_telecom@cablena.com.br"
        poSendMail.CcRecipient = "arbarbeiro@cablena.com.br;erocha@cablena.com.br;fmaia@cablena.com.br"
        
    ElseIf VElemento = "6.7 - Movimento de materiais" Then
        poSendMail.Recipient = "simac_telecom@cablena.com.br"
        poSendMail.CcRecipient = "josealopes@cablena.com.br;jefferson.silva@cablena.com.br"
    Else
        MsgBox "Elemento não encontrado!", vbExclamation, "Suporte Simac"
    
    End If
 
    Call suRelatorio(vAtendimento)
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
    MsgBox "Erro: " & Err.Description, vbCritical, "Suporte Simac"
    Set poSendMail = Nothing
End Function

Private Sub suRelatorio(ByVal vAtendimento As Long)

'On Error GoTo Erro
    strSQL = "SELECT * FROM vw_Chamados WHERE ATID = " & vAtendimento
    
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
'Erro:
'    MsgBox "erro " & Err.Description, vbCritical, "Suporte Simac relatorio"
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
