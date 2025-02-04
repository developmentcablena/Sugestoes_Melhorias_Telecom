VERSION 5.00
Begin VB.Form frmAlterarSenha 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alterar Senha"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2895
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtConfirmar 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtSenhaAtual 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox txtNovaSenha 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblConfirmar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar Senha"
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
      TabIndex        =   7
      Top             =   1320
      Width           =   1635
   End
   Begin VB.Label lblSenhaAtual 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha Atual"
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
      TabIndex        =   6
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label lblNovaSenha 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nova Senha"
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
      TabIndex        =   5
      Top             =   720
      Width           =   1155
   End
   Begin VB.Image imgSeguranca 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2280
      Picture         =   "frmAlterarSenha_T.frx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmAlterarSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String

Private Sub cmdAlterar_Click()
'    If Len(Trim(txtSenhaAtual.Text)) = 0 Then
'        MsgBox "Digite a senha atual!", vbOKOnly + vbExclamation, "Suporte Manutenção"
'        txtSenhaAtual.SetFocus
'        Exit Sub
'    End If
    
    If Len(Trim(txtNovaSenha.Text)) = 0 Then
        MsgBox "Digite uma nova senha!", vbOKOnly + vbExclamation, "Suporte Simac"
        txtNovaSenha.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtConfirmar.Text)) = 0 Then
        MsgBox "Digite a confirmação da senha!", vbOKOnly + vbExclamation, "Suporte Simac"
        txtConfirmar.SetFocus
        Exit Sub
    End If
    
    If CodDec(Trim(txtSenhaAtual.Text)) <> gstrSenha Then
        MsgBox "Senha atual não confere!", vbOKOnly + vbCritical, "Suporte Simac"
        txtSenhaAtual.Text = ""
        txtSenhaAtual.SetFocus
        Exit Sub
    End If
    
    If Trim(txtNovaSenha.Text) <> Trim(txtConfirmar.Text) Then
        MsgBox "Nova senha diferente da confirmação!", vbOKOnly + vbExclamation, "Suporte Simac"
        txtNovaSenha.Text = ""
        txtConfirmar.Text = ""
        txtNovaSenha.SetFocus
        Exit Sub
    Else
     
        Call suAlterarSenha(gintUsuarioID, CodDec(Trim(txtNovaSenha.Text)))
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    Call ConectarBD
End Sub

Private Sub txtConfirmar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtConfirmar.Text)) > 0 Then
            Call cmdAlterar_Click
        Else
            MsgBox "Digite a confirmação da senha!", vbOKOnly + vbExclamation, "Suporte Simac"
            txtConfirmar.SetFocus
        End If
    End If
End Sub

Private Sub txtNovaSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtNovaSenha.Text)) > 0 Then
            txtConfirmar.SetFocus
        Else
            MsgBox "Digite uma nova senha!", vbOKOnly + vbExclamation, "Suporte Simac"
            txtNovaSenha.SetFocus
        End If
    End If
End Sub

Private Sub txtSenhaAtual_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtNovaSenha.SetFocus
    End If
End Sub

Private Sub suAlterarSenha(ByVal vUsuarioID As Integer, ByVal vSenha As String)
On Error GoTo Erro

    strSQL = "UPDATE tb_Usuario SET Senha = '" & vSenha & "',AlterarSenha = 0 WHERE UsuarioID = " & vUsuarioID & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    MsgBox "Senha alterada com sucesso!", vbOKOnly + vbInformation, "Suporte Simac"
    Set rs = Nothing
    Call Unload(Me)
    Call frmPrincipal.Show
    Exit Sub
    
Erro:
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Simac"
    Set rs = Nothing
End Sub
