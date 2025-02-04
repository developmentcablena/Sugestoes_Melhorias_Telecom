VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFinalizarOS 
   Caption         =   "Reporte Tecnico - (Finalizar Atendimento)"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   Icon            =   "frmFinalizarOS_T.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   6405
   StartUpPosition =   1  'CenterOwner
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
      Left            =   5040
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "&Finalizar"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
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
      TabIndex        =   2
      Top             =   960
      Width           =   6255
   End
   Begin VB.TextBox txtOS1 
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
      TabIndex        =   0
      Text            =   "         "
      Top             =   120
      Width           =   975
   End
   Begin MSMask.MaskEdBox mkdHorario 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4200
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
   Begin VB.Label Label3 
      Caption         =   "Data/ Hora Final."
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
      TabIndex        =   5
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Reporte Técnico:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2295
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
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmFinalizarOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim strSQL As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Tempo()
Sleep 5000
End Sub

Private Sub cmdFinalizar_Click()

If Len(Trim(Me.txtDecricao.Text)) = 0 Then
    MsgBox "Digite o Reporte Tecnico! ", vbInformation, "Suporte Simac"
    Me.txtDecricao.SetFocus
    Exit Sub
End If


If IsDate(Me.mkdHorario) = False Then
    MsgBox "Digite Data/ Hora de finalização do Atendimento! ", vbOKOnly + vbExclamation, "Suporte Simac"
    Me.mkdHorario.SetFocus
    Exit Sub
End If

If Me.cmdFinalizar.Caption = "&Finalizar" Then
    If fnFinalizarOS(gintOSID, Me.txtDecricao, Me.mkdHorario) = True Then
        MsgBox "Atendimento N° " & Format(gintOSID, "0000") & " Finalizada!", vbInformation, "Suporte Simac"
        Call LimparDados
        gintOSID = 0
    End If
Else
    If fnCancelarOS(gintOSID, Me.txtDecricao, Me.mkdHorario) = True Then
        MsgBox "Atendimento N° " & Format(gintOSID, "0000") & " Cancelada!", vbInformation, "Suporte Simac"
        Call LimparDados
        gintOSID = 0
    End If
End If


End Sub

Private Sub Command1_Click()
Call Unload(Me)

End Sub

Private Sub Form_Load()
Me.txtOS1.Enabled = False
Me.txtOS1 = Format(gintOSID, "0000")
Me.mkdHorario.Text = Format(Now, "dd/MM/yyyy HH:mm")
End Sub
Private Sub LimparDados()
Me.txtOS1 = ""
Me.txtDecricao = ""
End Sub
Private Function fnCancelarOS(ByVal vOSID3 As Long, ByVal vDescricao3 As String, ByVal vData3 As Date)
On Error GoTo Erro
Call ConectarBD
    fnCancelarOS = False
    Me.Enabled = flase
    Screen.MousePointer = 11
    
    strSQL = "UPDATE tb_Atendimento SET " & _
    "MotivoCancelamento = '" & vDescricao3 & "', " & _
    "DataCancelamento = '" & Format(vData3, "yyyy-MM-dd HH:mm") & "', " & _
    "Status = 3" & _
    "WHERE ATID = " & vOSID3 & " "

    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    fnCancelarOS = True
    Me.Enabled = True
    Call Tempo
    Screen.MousePointer = 0
    
    Set rs = Nothing
    Exit Function
Erro:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Simac"
    
End Function
Private Function fnFinalizarOS(ByVal vOSID As Long, ByVal vDescricao2 As String, ByVal vData As Date)
Call ConectarBD
On Error GoTo Erro
    fnFinalizarOS = False
    Me.Enabled = False
    Screen.MousePointer = 11
    
    strSQL = "UPDATE tb_Atendimento set " & _
    "FinalizacaoOS = '" & vDescricao2 & "', " & _
    "DataBaixaAtual = '" & Format(vData, "yyyy-MM-dd HH:mm") & "', " & _
    "Status = 2 " & _
    "WHERE ATID = " & vOSID & " "
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    fnFinalizarOS = True
    Me.Enabled = True
    Call Tempo
    Screen.MousePointer = 0
    
    Set rs = Nothing
    Exit Function
Erro:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    MsgBox "Erro: " & Err.Description, vbCritical, "Suporte Simac"

End Function

